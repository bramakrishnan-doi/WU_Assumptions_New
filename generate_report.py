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


def repo_path(value: str, field: str) -> Path:
    p = Path(str(value or "").strip())
    if not str(p):
        raise RuntimeError(f"{field} is required")
    if p.is_absolute():
        raise RuntimeError(f"{field} must be repository-relative")

    out = (ROOT / p).resolve()

    try:
        out.relative_to(ROOT)
    except ValueError as exc:
        raise RuntimeError(f"{field} must stay inside the repository") from exc

    return out


def sha256(path: Path) -> str:
    h = hashlib.sha256()

    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)

    return h.hexdigest()


def num(value):
    try:
        return Decimal(str(value or "").replace(",", "").strip())
    except InvalidOperation:
        return None


def maf(text):
    m = re.search(
        r"([0-9]+(?:\.[0-9]+)?)\s+maf\b",
        str(text),
        re.I,
    )
    return Decimal(m.group(1)) if m else None


def qa(context):
    checks = []

    def add(name, passed, detail):
        checks.append(
            {
                "name": name,
                "passed": bool(passed),
                "detail": detail,
            }
        )

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

    for sec in context.years:
        states = [num(s.maf_value) for s in sec.states[:3]]
        us = num(sec.us_contractors_maf)
        total = num(sec.total_use_maf)
        mx = maf(sec.mexico_heading)

        if (
            len(states) == 3
            and all(v is not None for v in states)
            and us is not None
        ):
            expected = sum(states, Decimal(0))

            add(
                f"{sec.year}_us_equals_states",
                abs(expected - us) <= Decimal("0.003"),
                f"states={expected}; US={us}",
            )

        if total is not None and us is not None and mx is not None:
            add(
                f"{sec.year}_total_equals_us_plus_mexico",
                abs((us + mx) - total) <= Decimal("0.003"),
                f"US+Mexico={us + mx}; total={total}",
            )

    rows = {r.state: r for r in context.ics_table.rows}

    if all(k in rows for k in ("AZ", "CA", "NV", "Total")):
        for i, year in enumerate(context.ics_table.year_labels):
            vals = [
                num(rows[k].values[i])
                for k in ("AZ", "CA", "NV")
            ]

            total = num(rows["Total"].values[i])

            if total is not None and all(v is not None for v in vals):
                expected = sum(vals, Decimal(0))

                add(
                    f"{year}_ics_total",
                    expected == total,
                    f"AZ+CA+NV={expected}; total={total}",
                )

    if context.show_conservation_summary:
        rows = {
            r.state: r
            for r in context.conservation_summary.rows
        }

        if all(
            k in rows
            for k in ("AZ", "CA", "NV", "Annual Total")
        ):
            for i, year in enumerate(
                context.conservation_summary.year_labels
            ):
                vals = [
                    num(rows[k].values[i])
                    for k in ("AZ", "CA", "NV")
                ]

                total = num(rows["Annual Total"].values[i])

                if (
                    total is not None
                    and all(v is not None for v in vals)
                ):
                    expected = sum(vals, Decimal(0))

                    add(
                        f"{year}_conservation_total",
                        abs(expected - total) <= Decimal(2),
                        f"AZ+CA+NV={expected}; total={total}",
                    )

    return checks


def run(
    config_file: str,
    output_dir: str,
    preflight_only: bool,
):
    cfg_path = repo_path(config_file, "config")

    cfg = json.loads(
        cfg_path.read_text(encoding="utf-8")
    )

    report_month = str(
        cfg.get("report_month", "")
    ).strip()

    datetime.strptime(report_month, "%Y-%m")

    mon_year = datetime.strptime(
        report_month,
        "%Y-%m",
    ).strftime("%B %Y")

    policy_path = repo_path(
        cfg.get(
            "policy_file",
            "config/report_policy.json",
        ),
        "policy_file",
    )

    workbook = repo_path(
        cfg.get("workbook", ""),
        "workbook",
    )

    policy = load_policy(policy_path)

    if (
        not workbook.is_file()
        or workbook.suffix.lower() != ".xlsx"
    ):
        raise RuntimeError(
            f"Workbook not found or not .xlsx: {workbook}"
        )

    models = {}

    for key, item in (
        cfg.get("scenarios") or {}
    ).items():

        if (
            not isinstance(item, dict)
            or not item.get("enabled")
        ):
            continue

        path = repo_path(
            item.get("model", ""),
            f"scenarios.{key}.model",
        )

        if (
            not path.is_file()
            or not path.name.lower().endswith(
                (".mdl", ".mdl.gz")
            )
        ):
            raise RuntimeError(
                f"Invalid {key} model: {path}"
            )

        models[key] = path

    if not models:
        raise RuntimeError(
            "Enable at least one scenario in "
            "config/monthly_run.json"
        )

    opened = []

    try:
        uploads = {}

        for key, path in models.items():
            fh = path.open("rb")
            opened.append(fh)

            uploads[key] = (
                path.name,
                fh,
            )

        prepared = prepare_batch(
            model_uploads=uploads,
            workbook_filename=workbook.name,
            workbook_bytes=workbook.read_bytes(),
            policy=policy,
        )

    finally:
        for fh in opened:
            fh.close()

    options = cfg.get("options") or {}

    if options.get(
        "validate_report_month_against_model_start",
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
                    f"{key} model starts {actual}"
                )

        if wrong:
            raise RuntimeError(
                "report_month/model mismatch: "
                + "; ".join(wrong)
            )

    out = repo_path(
        output_dir,
        "output_dir",
    )

    out.mkdir(
        parents=True,
        exist_ok=True,
    )

    for pattern in (
        "*.docx",
        "*.zip",
        "run_summary.json",
        "run_summary.md",
    ):
        for old in out.glob(pattern):
            old.unlink()

    summary = {
        "status":
            "preflight_only"
            if preflight_only
            else "generated",
        "report_month": report_month,
        "report_month_display": mon_year,
        "policy_version": policy.version,
        "git_sha": os.getenv(
            "GITHUB_SHA",
            "",
        ),
        "report_years": list(
            prepared.years
        ),
        "warnings": list(
            prepared.warnings
        ),
        "preflight":
            prepared.preflight_dict(policy),
        "workbook": {
            "filename": workbook.name,
            "sha256": sha256(workbook),
        },
        "scenarios": [],
        "outputs": [],
        "qa_checks": [],
    }

    for key in prepared.scenarios:
        p = models[key]

        summary["scenarios"].append(
            {
                "key": key,
                "label":
                    policy
                    .scenario(key)
                    .display_name,
                "filename": p.name,
                "sha256": sha256(p),
            }
        )

    if not preflight_only:
        docs = []

        for key in prepared.scenarios:
            context = build_report_context(
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

            checks = qa(context)

            summary["qa_checks"].extend(
                {
                    "scenario": key,
                    **c,
                }
                for c in checks
            )

            failed = [
                c
                for c in checks
                if not c["passed"]
            ]

            if failed:
                raise RuntimeError(
                    "Automated QA failed: "
                    + "; ".join(
                        f"{key}:{c['name']} "
                        f"({c['detail']})"
                        for c in failed
                    )
                )

            name = (
                "24-MS LB Water Use Projections - "
                f"{mon_year} "
                f"{context.scenario_label}.docx"
            )

            path = out / name

            data = build_docx_bytes(
                context
            )

            if not data.startswith(b"PK"):
                raise RuntimeError(
                    "Generated file is not "
                    f"a valid DOCX: {name}"
                )

            path.write_bytes(data)
            docs.append(path)

            summary["outputs"].append(
                {
                    "scenario": key,
                    "filename": name,
                    "sha256": sha256(path),
                    "size_bytes":
                        path.stat().st_size,
                }
            )

        if (
            options.get(
                "create_batch_zip",
                True,
            )
            and len(docs) > 1
        ):
            zpath = (
                out
                / (
                    "LB Water Use Projections - "
                    f"{mon_year} - Batch.zip"
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

            summary["outputs"].append(
                {
                    "scenario": "batch",
                    "filename": zpath.name,
                    "sha256": sha256(zpath),
                    "size_bytes":
                        zpath.stat().st_size,
                }
            )

    if (
        options.get(
            "fail_on_warnings",
            False,
        )
        and summary["warnings"]
    ):
        raise RuntimeError(
            "Warnings are configured as fatal: "
            + "; ".join(
                summary["warnings"]
            )
        )

    md = [
        "# Water Use Assumptions Generator - Monthly Run",
        "",
        f"- Status: **{summary['status']}**",
        f"- Study month: **{mon_year}**",
        f"- Policy version: **{policy.version}**",
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
        f"- Workbook: `{workbook.name}`",
    ]

    md += [
        (
            f"- {s['label']}: "
            f"`{s['filename']}`"
        )
        for s in summary["scenarios"]
    ]

    md += [
        "",
        "## Validation",
        "",
        (
            "- Warnings: "
            f"**{len(summary['warnings'])}**"
        ),
    ]

    if summary["qa_checks"]:
        passed = sum(
            1
            for c in summary["qa_checks"]
            if c["passed"]
        )

        md.append(
            "- Automated QA checks: "
            f"**{passed}/"
            f"{len(summary['qa_checks'])} passed**"
        )

    if summary["outputs"]:
        md += [
            "",
            "## Generated Files",
            "",
        ] + [
            f"- `{o['filename']}`"
            for o in summary["outputs"]
        ]

    text = "\n".join(md) + "\n"

    (
        out / "run_summary.json"
    ).write_text(
        json.dumps(
            summary,
            indent=2,
        ),
        encoding="utf-8",
    )

    (
        out / "run_summary.md"
    ).write_text(
        text,
        encoding="utf-8",
    )

    print(text)


def main():
    p = argparse.ArgumentParser()

    p.add_argument(
        "--config",
        default="config/monthly_run.json",
    )

    p.add_argument(
        "--output-dir",
        default="output",
    )

    p.add_argument(
        "--preflight-only",
        action="store_true",
    )

    a = p.parse_args()

    try:
        run(
            a.config,
            a.output_dir,
            a.preflight_only,
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
    raise SystemExit(main())