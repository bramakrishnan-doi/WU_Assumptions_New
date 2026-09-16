from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
import zipfile
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from docx_writer import build_docx_bytes
from policy import PolicyError, load_policy
from report_builder import build_report_context
from report_service import BatchValidationError, prepare_batch


ROOT = Path(__file__).resolve().parent

# Small accounting tolerances used only for post-generation QA.
# These allow for display-level rounding while still failing on meaningful
# accounting mismatches.
ICS_ROUNDING_TOLERANCE_AF = Decimal("1")
CONSERVATION_ROUNDING_TOLERANCE_AF = Decimal("2")


def repo_path(value: str, field: str) -> Path:
    """
    Resolve a repository-relative path and prevent the configuration
    from referring to files outside the project directory.
    """
    p = Path(str(value or "").strip())

    if not str(p):
        raise RuntimeError(f"{field} is required")

    if p.is_absolute():
        raise RuntimeError(f"{field} must be repository-relative")

    out = (ROOT / p).resolve()

    try:
        out.relative_to(ROOT)
    except ValueError as exc:
        raise RuntimeError(
            f"{field} must stay inside the repository"
        ) from exc

    return out


def sha256(path: Path) -> str:
    """
    Calculate a SHA-256 hash for run provenance.
    """
    h = hashlib.sha256()

    with path.open("rb") as f:
        for chunk in iter(
            lambda: f.read(1024 * 1024),
            b"",
        ):
            h.update(chunk)

    return h.hexdigest()


def num(value):
    """
    Convert a displayed report value into Decimal.
    Commas are removed before conversion.
    """
    try:
        return Decimal(
            str(value or "")
            .replace(",", "")
            .strip()
        )
    except InvalidOperation:
        return None


def maf(text):
    """
    Extract a displayed MAF value from text such as:
        Mexico's Scheduled Water Delivery: 1.250 maf
    """
    m = re.search(
        r"([0-9]+(?:\.[0-9]+)?)\s+maf\b",
        str(text),
        re.I,
    )

    return Decimal(m.group(1)) if m else None


def qa(context):
    """
    Run deterministic post-generation accounting checks against
    the same report context used to create the Word document.

    These checks are intentionally narrow. They validate internal
    arithmetic relationships without changing any report values.
    """
    checks = []

    def add(name, passed, detail):
        checks.append(
            {
                "name": name,
                "passed": bool(passed),
                "detail": detail,
            }
        )

    # Basic context checks.
    add(
        "source_map_nonempty",
        bool(context.source_map),
        f"{len(context.source_map)} source entries",
    )

    add(
        "report_has_years",
        bool(context.years),
        f"{len(context.years)} year sections",
    )

    # ------------------------------------------------------------
    # Annual state / U.S. / Mexico accounting
    # ------------------------------------------------------------
    #
    # These are displayed in MAF and may be rounded to three
    # decimal places. A tolerance of 0.003 MAF is therefore used.
    #
    for sec in context.years:
        states = [
            num(s.maf_value)
            for s in sec.states[:3]
        ]

        us = num(sec.us_contractors_maf)
        total = num(sec.total_use_maf)
        mx = maf(sec.mexico_heading)

        if (
            len(states) == 3
            and all(
                v is not None
                for v in states
            )
            and us is not None
        ):
            expected = sum(
                states,
                Decimal(0),
            )

            difference = abs(
                expected - us
            )

            add(
                f"{sec.year}_us_equals_states",
                difference <= Decimal("0.003"),
                (
                    f"states={expected}; "
                    f"US={us}; "
                    f"difference={difference} MAF"
                ),
            )

        if (
            total is not None
            and us is not None
            and mx is not None
        ):
            expected = us + mx
            difference = abs(
                expected - total
            )

            add(
                f"{sec.year}_total_equals_us_plus_mexico",
                difference <= Decimal("0.003"),
                (
                    f"US+Mexico={expected}; "
                    f"total={total}; "
                    f"difference={difference} MAF"
                ),
            )

    # ------------------------------------------------------------
    # ICS totals
    # ------------------------------------------------------------
    #
    # State ICS values are displayed as whole acre-feet.
    # Individual state values and the total can round independently.
    #
    # For example:
    #
    #   AZ + CA + NV = 1,912,104 AF
    #   Report Total = 1,912,103 AF
    #
    # A 1 AF difference is acceptable display-level rounding.
    #
    rows = {
        r.state: r
        for r in context.ics_table.rows
    }

    if all(
        k in rows
        for k in (
            "AZ",
            "CA",
            "NV",
            "Total",
        )
    ):
        for i, year in enumerate(
            context.ics_table.year_labels
        ):
            vals = [
                num(rows[k].values[i])
                for k in (
                    "AZ",
                    "CA",
                    "NV",
                )
            ]

            total = num(
                rows["Total"].values[i]
            )

            if (
                total is not None
                and all(
                    v is not None
                    for v in vals
                )
            ):
                expected = sum(
                    vals,
                    Decimal(0),
                )

                difference = abs(
                    expected - total
                )

                add(
                    f"{year}_ics_total",
                    (
                        difference
                        <= ICS_ROUNDING_TOLERANCE_AF
                    ),
                    (
                        f"AZ+CA+NV={expected}; "
                        f"total={total}; "
                        f"difference={difference} AF; "
                        f"tolerance="
                        f"{ICS_ROUNDING_TOLERANCE_AF} AF"
                    ),
                )

    # ------------------------------------------------------------
    # Conservation summary totals
    # ------------------------------------------------------------
    #
    # Conservation components are also displayed as whole
    # acre-feet. Because three independently rounded values are
    # summed, allow up to 2 AF of difference.
    #
    if context.show_conservation_summary:
        rows = {
            r.state: r
            for r
            in context.conservation_summary.rows
        }

        if all(
            k in rows
            for k in (
                "AZ",
                "CA",
                "NV",
                "Annual Total",
            )
        ):
            for i, year in enumerate(
                context
                .conservation_summary
                .year_labels
            ):
                vals = [
                    num(rows[k].values[i])
                    for k in (
                        "AZ",
                        "CA",
                        "NV",
                    )
                ]

                total = num(
                    rows[
                        "Annual Total"
                    ].values[i]
                )

                if (
                    total is not None
                    and all(
                        v is not None
                        for v in vals
                    )
                ):
                    expected = sum(
                        vals,
                        Decimal(0),
                    )

                    difference = abs(
                        expected - total
                    )

                    add(
                        (
                            f"{year}_"
                            "conservation_total"
                        ),
                        (
                            difference
                            <= CONSERVATION_ROUNDING_TOLERANCE_AF
                        ),
                        (
                            f"AZ+CA+NV={expected}; "
                            f"total={total}; "
                            f"difference="
                            f"{difference} AF; "
                            f"tolerance="
                            f"{CONSERVATION_ROUNDING_TOLERANCE_AF} AF"
                        ),
                    )

    return checks


def run(
    config_file: str,
    output_dir: str,
    preflight_only: bool,
):
    """
    Execute one configured monthly Water Use report run.
    """

    # ------------------------------------------------------------
    # Load monthly configuration
    # ------------------------------------------------------------
    cfg_path = repo_path(
        config_file,
        "config",
    )

    cfg = json.loads(
        cfg_path.read_text(
            encoding="utf-8"
        )
    )

    report_month = str(
        cfg.get(
            "report_month",
            "",
        )
    ).strip()

    # Validate YYYY-MM format.
    datetime.strptime(
        report_month,
        "%Y-%m",
    )

    mon_year = datetime.strptime(
        report_month,
        "%Y-%m",
    ).strftime(
        "%B %Y"
    )

    # ------------------------------------------------------------
    # Load policy and workbook
    # ------------------------------------------------------------
    policy_path = repo_path(
        cfg.get(
            "policy_file",
            "config/report_policy.json",
        ),
        "policy_file",
    )

    workbook = repo_path(
        cfg.get(
            "workbook",
            "",
        ),
        "workbook",
    )

    policy = load_policy(
        policy_path
    )

    if (
        not workbook.is_file()
        or workbook.suffix.lower()
        != ".xlsx"
    ):
        raise RuntimeError(
            "Workbook not found or "
            f"not .xlsx: {workbook}"
        )

    # ------------------------------------------------------------
    # Resolve enabled model scenarios
    # ------------------------------------------------------------
    models = {}

    for key, item in (
        cfg.get("scenarios")
        or {}
    ).items():

        if (
            not isinstance(
                item,
                dict,
            )
            or not item.get(
                "enabled"
            )
        ):
            continue

        path = repo_path(
            item.get(
                "model",
                "",
            ),
            f"scenarios.{key}.model",
        )

        if (
            not path.is_file()
            or not path.name.lower().endswith(
                (
                    ".mdl",
                    ".mdl.gz",
                )
            )
        ):
            raise RuntimeError(
                f"Invalid {key} model: "
                f"{path}"
            )

        models[key] = path

    if not models:
        raise RuntimeError(
            "Enable at least one scenario "
            "in config/monthly_run.json"
        )

    # ------------------------------------------------------------
    # Use the same application preflight used by the web app
    # ------------------------------------------------------------
    opened = []

    try:
        uploads = {}

        for key, path in models.items():
            fh = path.open(
                "rb"
            )

            opened.append(
                fh
            )

            uploads[key] = (
                path.name,
                fh,
            )

        prepared = prepare_batch(
            model_uploads=uploads,
            workbook_filename=
                workbook.name,
            workbook_bytes=
                workbook.read_bytes(),
            policy=policy,
        )

    finally:
        for fh in opened:
            fh.close()

    options = (
        cfg.get("options")
        or {}
    )

    # ------------------------------------------------------------
    # Make sure the configured report month matches each model
    # ------------------------------------------------------------
    if options.get(
        (
            "validate_report_month_"
            "against_model_start"
        ),
        True,
    ):
        wrong = []

        for key in prepared.scenarios:
            actual = (
                prepared
                .annual_by_scenario[key]
                .run_start
                .strftime("%Y-%m")
            )

            if actual != report_month:
                wrong.append(
                    (
                        f"{key} model "
                        f"starts {actual}"
                    )
                )

        if wrong:
            raise RuntimeError(
                "report_month/model "
                "mismatch: "
                + "; ".join(
                    wrong
                )
            )

    # ------------------------------------------------------------
    # Prepare output folder
    # ------------------------------------------------------------
    out = repo_path(
        output_dir,
        "output_dir",
    )

    out.mkdir(
        parents=True,
        exist_ok=True,
    )

    # Remove products from the previous local run.
    for pattern in (
        "*.docx",
        "*.zip",
        "run_summary.json",
        "run_summary.md",
    ):
        for old in out.glob(
            pattern
        ):
            old.unlink()

    # ------------------------------------------------------------
    # Build run summary metadata
    # ------------------------------------------------------------
    summary = {
        "status": (
            "preflight_only"
            if preflight_only
            else "generated"
        ),
        "report_month":
            report_month,
        "report_month_display":
            mon_year,
        "policy_version":
            policy.version,
        "git_sha":
            os.getenv(
                "GITHUB_SHA",
                "",
            ),
        "report_years":
            list(
                prepared.years
            ),
        "warnings":
            list(
                prepared.warnings
            ),
        "preflight":
            prepared.preflight_dict(
                policy
            ),
        "workbook": {
            "filename":
                workbook.name,
            "sha256":
                sha256(
                    workbook
                ),
        },
        "scenarios": [],
        "outputs": [],
        "qa_checks": [],
    }

    for key in prepared.scenarios:
        p = models[key]

        summary[
            "scenarios"
        ].append(
            {
                "key":
                    key,
                "label":
                    policy
                    .scenario(key)
                    .display_name,
                "filename":
                    p.name,
                "sha256":
                    sha256(p),
            }
        )

    # ------------------------------------------------------------
    # Generate reports unless --preflight-only was requested
    # ------------------------------------------------------------
    if not preflight_only:
        docs = []

        for key in prepared.scenarios:

            context = (
                build_report_context(
                    annual=
                        prepared
                        .annual_by_scenario[key],
                    state_use=
                        prepared
                        .state_use_by_scenario[key],
                    scenario=key,
                    mon_year=mon_year,
                    policy=policy,
                )
            )

            # Run deterministic QA before writing the document.
            checks = qa(
                context
            )

            summary[
                "qa_checks"
            ].extend(
                {
                    "scenario":
                        key,
                    **c,
                }
                for c in checks
            )

            failed = [
                c
                for c in checks
                if not c[
                    "passed"
                ]
            ]

            if failed:
                raise RuntimeError(
                    "Automated QA failed: "
                    + "; ".join(
                        (
                            f"{key}:"
                            f"{c['name']} "
                            f"("
                            f"{c['detail']}"
                            f")"
                        )
                        for c in failed
                    )
                )

            scenario_label = (
                context
                .scenario_label
            )

            name = (
                "24-MS LB Water Use "
                "Projections - "
                f"{mon_year} "
                f"{scenario_label}.docx"
            )

            path = out / name

            data = build_docx_bytes(
                context
            )

            # DOCX files are ZIP-based and therefore
            # begin with the ZIP PK signature.
            if not data.startswith(
                b"PK"
            ):
                raise RuntimeError(
                    "Generated file is not "
                    "a valid DOCX: "
                    f"{name}"
                )

            path.write_bytes(
                data
            )

            docs.append(
                path
            )

            summary[
                "outputs"
            ].append(
                {
                    "scenario":
                        key,
                    "filename":
                        name,
                    "sha256":
                        sha256(path),
                    "size_bytes":
                        path.stat().st_size,
                }
            )

        # --------------------------------------------------------
        # Optional batch ZIP
        # --------------------------------------------------------
        #
        # For GitHub Actions this is normally unnecessary because
        # GitHub already packages all output files as an artifact.
        #
        # Set:
        #
        #   "create_batch_zip": false
        #
        # in monthly_run.json to disable it.
        #
        if (
            options.get(
                "create_batch_zip",
                False,
            )
            and len(docs) > 1
        ):
            zpath = (
                out
                / (
                    "LB Water Use "
                    "Projections - "
                    f"{mon_year} "
                    "- Batch.zip"
                )
            )

            with zipfile.ZipFile(
                zpath,
                "w",
                zipfile.ZIP_DEFLATED,
            ) as z:

                for p in docs:
                    z.write(
                        p,
                        p.name,
                    )

            summary[
                "outputs"
            ].append(
                {
                    "scenario":
                        "batch",
                    "filename":
                        zpath.name,
                    "sha256":
                        sha256(
                            zpath
                        ),
                    "size_bytes":
                        zpath
                        .stat()
                        .st_size,
                }
            )

    # ------------------------------------------------------------
    # Optionally treat warnings as errors
    # ------------------------------------------------------------
    if (
        options.get(
            "fail_on_warnings",
            False,
        )
        and summary[
            "warnings"
        ]
    ):
        raise RuntimeError(
            "Warnings are configured "
            "as fatal: "
            + "; ".join(
                summary[
                    "warnings"
                ]
            )
        )

    # ------------------------------------------------------------
    # Create GitHub/local run summary
    # ------------------------------------------------------------
    md = [
        (
            "# Water Use Assumptions "
            "Generator - Monthly Run"
        ),
        "",
        (
            f"- Status: "
            f"**{summary['status']}**"
        ),
        (
            f"- Study month: "
            f"**{mon_year}**"
        ),
        (
            f"- Policy version: "
            f"**{policy.version}**"
        ),
        (
            "- Report years: **"
            + ", ".join(
                map(
                    str,
                    prepared.years,
                )
            )
            + "**"
        ),
        "",
        "## Inputs",
        "",
        (
            f"- Workbook: "
            f"`{workbook.name}`"
        ),
    ]

    md += [
        (
            f"- {s['label']}: "
            f"`{s['filename']}`"
        )
        for s
        in summary[
            "scenarios"
        ]
    ]

    md += [
        "",
        "## Validation",
        "",
        (
            "- Warnings: "
            f"**"
            f"{len(summary['warnings'])}"
            f"**"
        ),
    ]

    if summary[
        "qa_checks"
    ]:
        passed = sum(
            1
            for c
            in summary[
                "qa_checks"
            ]
            if c[
                "passed"
            ]
        )

        md.append(
            (
                "- Automated QA checks: "
                f"**{passed}/"
                f"{len(summary['qa_checks'])} "
                "passed**"
            )
        )

    if summary[
        "outputs"
    ]:
        md += [
            "",
            "## Generated Files",
            "",
        ] + [
            (
                f"- "
                f"`{o['filename']}`"
            )
            for o
            in summary[
                "outputs"
            ]
        ]

    text = (
        "\n".join(md)
        + "\n"
    )

    (
        out
        / "run_summary.json"
    ).write_text(
        json.dumps(
            summary,
            indent=2,
        ),
        encoding="utf-8",
    )

    (
        out
        / "run_summary.md"
    ).write_text(
        text,
        encoding="utf-8",
    )

    print(
        text
    )


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Generate Water Use "
            "Assumptions reports "
            "without the Flask UI."
        )
    )

    parser.add_argument(
        "--config",
        default=(
            "config/"
            "monthly_run.json"
        ),
        help=(
            "Repository-relative "
            "monthly configuration "
            "JSON file."
        ),
    )

    parser.add_argument(
        "--output-dir",
        default="output",
        help=(
            "Repository-relative "
            "output directory."
        ),
    )

    parser.add_argument(
        "--preflight-only",
        action="store_true",
        help=(
            "Validate inputs without "
            "creating Word reports."
        ),
    )

    args = (
        parser.parse_args()
    )

    try:
        run(
            args.config,
            args.output_dir,
            args.preflight_only,
        )

        return 0

    except (
        RuntimeError,
        ValueError,
        json.JSONDecodeError,
        PolicyError,
        BatchValidationError,
    ) as exc:

        print(
            f"ERROR: {exc}",
            file=sys.stderr,
        )

        if isinstance(
            exc,
            BatchValidationError,
        ):
            for item in exc.errors:
                print(
                    f"  - {item}",
                    file=sys.stderr,
                )

            for item in exc.warnings:
                print(
                    f"  warning: {item}",
                    file=sys.stderr,
                )

        return 2


if __name__ == "__main__":
    raise SystemExit(
        main()
    )