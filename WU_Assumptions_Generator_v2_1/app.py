"""Local Flask application for Lower Basin Water Use report generation."""
from __future__ import annotations

import base64
import io
import os
import traceback
import zipfile
from datetime import datetime

from flask import Flask, jsonify, render_template, request
from werkzeug.exceptions import RequestEntityTooLarge

from docx_writer import build_docx_bytes
from policy import PolicyError, load_policy
from report_builder import build_report_context
from report_service import BatchValidationError, PreparedBatchCache, prepare_batch

app = Flask(__name__)
# This is a local-only application. Batch uploads can contain three compressed
# model files; Flask may spool multipart data to disk before this limit.
app.config["MAX_CONTENT_LENGTH"] = 1200 * 1024 * 1024  # 1.2 GB total request

POLICY = load_policy()
CACHE = PreparedBatchCache(ttl_seconds=30 * 60, max_entries=5)


def _json_error(message: str, status: int = 400, *, errors=None, warnings=None):
    return (
        jsonify(
            {
                "ok": False,
                "error": message,
                "errors": list(errors or []),
                "warnings": list(warnings or []),
            }
        ),
        status,
    )


def _normalize_month(value: str) -> str:
    value = (value or "").strip()
    if not value:
        raise ValueError("Select the 24-Month Study month.")
    try:
        return datetime.strptime(value, "%Y-%m").strftime("%B %Y")
    except ValueError:
        # Retain compatibility with API clients that already send 'August 2026'.
        try:
            return datetime.strptime(value, "%B %Y").strftime("%B %Y")
        except ValueError as exc:
            raise ValueError("Study month must be in YYYY-MM or 'Month YYYY' format.") from exc


def _output_filename(mon_year: str, scenario_label: str) -> str:
    return f"24-MS LB Water Use Projections - {mon_year} {scenario_label}.docx"


@app.get("/")
def index():
    return render_template(
        "index.html",
        policy_version=POLICY.version,
        default_enabled_scenarios=POLICY.default_enabled_scenarios,
        scenarios=[POLICY.scenario(k) for k in POLICY.scenario_keys],
    )


@app.get("/api/health")
def health():
    return jsonify({"ok": True, "policy_version": POLICY.version})


@app.get("/api/rules")
def rules():
    """Return the exact active policy used by this running application."""
    return jsonify(
        {
            "ok": True,
            "policy_version": POLICY.version,
            "description": str(POLICY.raw.get("description", "")),
            "config_file": "config/report_policy.json",
            "restart_required_after_edit": True,
            "policy": POLICY.raw,
        }
    )


@app.post("/api/preflight")
def preflight():
    excel = request.files.get("excel_file")
    if excel is None or not excel.filename:
        return _json_error("Upload the current Projected State Use .xlsx workbook.")

    model_uploads = {}
    for scenario in POLICY.scenario_keys:
        item = request.files.get(f"model_{scenario}")
        if item is not None and item.filename:
            model_uploads[scenario] = (item.filename, item.stream)

    if not model_uploads:
        return _json_error("Upload at least one selected RiverWare scenario model.")

    try:
        workbook_bytes = excel.read()
        prepared = prepare_batch(
            model_uploads=model_uploads,
            workbook_filename=excel.filename,
            workbook_bytes=workbook_bytes,
            policy=POLICY,
        )
        CACHE.put(prepared)
        return jsonify({"ok": True, "preflight": prepared.preflight_dict(POLICY)})
    except BatchValidationError as exc:
        return _json_error(
            "Preflight found issues that must be corrected before report generation.",
            errors=exc.errors,
            warnings=exc.warnings,
        )


@app.post("/api/generate")
def generate():
    payload = request.get_json(silent=True) or {}
    token = str(payload.get("token") or "")
    prepared = CACHE.get(token)
    if prepared is None:
        return _json_error(
            "The preflight session expired or the inputs changed. Run preflight again.",
            status=409,
        )
    try:
        mon_year = _normalize_month(str(payload.get("mon_year") or ""))
    except ValueError as exc:
        return _json_error(str(exc))
    reports = []
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for scenario in prepared.scenarios:
            context = build_report_context(
                annual=prepared.annual_by_scenario[scenario],
                state_use=prepared.state_use_by_scenario[scenario],
                scenario=scenario,
                mon_year=mon_year,
                policy=POLICY,
            )
            docx_bytes = build_docx_bytes(context)
            filename = _output_filename(mon_year, context.scenario_label)
            archive.writestr(filename, docx_bytes)
            preview_html = render_template("report_preview.html", report=context)
            reports.append(
                {
                    "scenario": scenario,
                    "scenario_label": context.scenario_label,
                    "filename": filename,
                    "docx_base64": base64.b64encode(docx_bytes).decode("ascii"),
                    "preview_html": preview_html,
                    "source_map": context.source_map,
                    "warnings": context.warnings,
                }
            )

    response = {"ok": True, "reports": reports}
    if len(reports) > 1:
        response["zip_filename"] = f"LB Water Use Projections - {mon_year} - Batch.zip"
        response["zip_base64"] = base64.b64encode(zip_buffer.getvalue()).decode("ascii")
    return jsonify(response)


@app.errorhandler(RequestEntityTooLarge)
def too_large(_error):
    return _json_error(
        "The combined upload is larger than the application's 1.2 GB request limit. "
        "Use .mdl.gz model files where possible or process fewer scenarios at once.",
        status=413,
    )


@app.errorhandler(PolicyError)
def policy_error(error):
    return _json_error(f"Report policy configuration error: {error}", status=500)


@app.errorhandler(Exception)
def unexpected_error(error):
    traceback.print_exc()
    return _json_error(f"Unexpected application error: {error}", status=500)


if __name__ == "__main__":
    debug = os.getenv("FLASK_DEBUG", "0") == "1"
    app.run(host="127.0.0.1", port=5051, debug=debug)
