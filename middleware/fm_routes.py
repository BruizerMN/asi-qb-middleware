"""
REST endpoints for FileMaker.

FileMaker uses Insert from URL to call these. All requests must include:
    X-API-Key: <shared secret from .env>

Endpoints:
    POST /fm/invoice          — queue a new invoice job
    GET  /fm/job/<job_id>     — check job status / retrieve QB invoice ID
"""

import json
from functools import wraps
from flask import Blueprint, request, jsonify

import middleware.db as db
from middleware.config import API_KEY
from middleware.qbxml_builder import build_invoice_add

bp = Blueprint("fm", __name__, url_prefix="/fm")


def require_api_key(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if request.headers.get("X-API-Key") != API_KEY:
            return jsonify({"error": "Unauthorized"}), 401
        return f(*args, **kwargs)
    return decorated


@bp.post("/invoice")
@require_api_key
def queue_invoice():
    """
    FM posts a JSON invoice payload here.

    Required top-level fields:
        company    — "acoustical" or "architectural"
        order_id   — FM order ID (for logging / deduplication)
        invoice    — the invoice dict (see qbxml_builder.py for schema)

    Returns:
        {"job_id": "<uuid>", "status": "pending"}
    """
    data = request.get_json(force=True, silent=True)
    if not data:
        return jsonify({"error": "Invalid JSON"}), 400

    company = data.get("company", "").lower()
    if company not in ("acoustical", "architectural"):
        return jsonify({"error": "company must be 'acoustical' or 'architectural'"}), 400

    order_id = data.get("order_id")
    if not order_id:
        return jsonify({"error": "order_id is required"}), 400

    invoice = data.get("invoice")
    if not invoice:
        return jsonify({"error": "invoice payload is required"}), 400

    qbxml = build_invoice_add(invoice)
    job_id = db.enqueue_job(company, str(order_id), qbxml)

    return jsonify({"job_id": job_id, "status": "pending"}), 202


@bp.get("/job/<job_id>")
@require_api_key
def get_job_status(job_id):
    """
    FM polls this to check if a job is done and retrieve the QB invoice ID.

    Returns:
        {
            "job_id": "...",
            "status": "pending|sent|done|error",
            "qb_invoice_id": "1042",   # present when status=done
            "error_msg": "...",        # present when status=error
            "order_id": "..."
        }
    """
    job = db.get_job(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404

    return jsonify({
        "job_id": job["id"],
        "status": job["status"],
        "order_id": job["order_id"],
        "qb_invoice_id": job.get("qb_invoice_id"),
        "error_msg": job.get("error_msg"),
    })
