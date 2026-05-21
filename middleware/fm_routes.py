"""
REST endpoints for FileMaker.

FileMaker uses Insert from URL to call these. All requests must include:
    X-API-Key: <shared secret from .env>

Endpoints:
    POST /fm/invoice        — submit an invoice to QuickBooks (synchronous via COM)
    POST /fm/ping           — verify QB is running and check which company is open
    POST /fm/sync-customers — return all active QB customers for FM to match and store
    POST /fm/sync-items     — return all active QB items for FM to match and store
    POST /fm/debug-invoice  — return generated qbXML without submitting (dev only)
"""

from functools import wraps
from flask import Blueprint, request, jsonify

from middleware.config import API_KEY
from middleware.qbxml_builder import build_invoice_add
import middleware.com_handler as com

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
def post_invoice():
    """
    FM posts a JSON invoice payload here. Submits synchronously to QB via COM.

    Required fields:
        company    — "acoustical" or "architectural"
        order_id   — FM order ID (included in response for FM to match)
        invoice    — invoice dict (see qbxml_builder.py for schema)

    Returns:
        {"status": "ok",    "qb_invoice_id": "1042", "order_id": "..."}
        {"status": "error", "error": "..."}
    """
    data = request.get_json(force=True, silent=True)
    if not data:
        return jsonify({"status": "error", "error": "Invalid JSON"}), 400

    company = data.get("company", "").lower()
    if company not in ("acoustical", "architectural"):
        return jsonify({"status": "error", "error": "company must be 'acoustical' or 'architectural'"}), 400

    order_id = data.get("order_id")
    if not order_id:
        return jsonify({"status": "error", "error": "order_id is required"}), 400

    invoice = data.get("invoice")
    if not invoice:
        return jsonify({"status": "error", "error": "invoice payload is required"}), 400

    try:
        qbxml = build_invoice_add(invoice)
        qb_invoice_id = com.submit_invoice(qbxml, company)
        return jsonify({
            "status": "ok",
            "qb_invoice_id": qb_invoice_id,
            "order_id": str(order_id),
        })
    except RuntimeError as e:
        return jsonify({"status": "error", "error": str(e)}), 422
    except Exception as e:
        return jsonify({"status": "error", "error": f"Unexpected error: {e}"}), 500


@bp.post("/sync-customers")
@require_api_key
def sync_customers():
    """
    Return all active top-level QB customers for FM to match against its CUSTOMER table.

    Required field:
        company — "acoustical" or "architectural"

    Returns:
        {"status": "ok", "count": N, "customers": [{"list_id": "...", "full_name": "...", "account_number": "..."}, ...]}
    """
    data    = request.get_json(force=True, silent=True) or {}
    company = data.get("company", "").lower()
    if company not in ("acoustical", "architectural"):
        return jsonify({"status": "error", "error": "company must be 'acoustical' or 'architectural'"}), 400

    try:
        customers = com.get_all_customers(company)
        return jsonify({"status": "ok", "count": len(customers), "customers": customers})
    except RuntimeError as e:
        return jsonify({"status": "error", "error": str(e)}), 422
    except Exception as e:
        return jsonify({"status": "error", "error": f"Unexpected error: {e}"}), 500


@bp.post("/sync-items")
@require_api_key
def sync_items():
    """
    Return all active QB non-inventory items for FM to match against its PRODUCTS table.

    Required field:
        company — "acoustical" or "architectural"

    Returns:
        {"status": "ok", "count": N, "items": [{"list_id": "...", "name": "..."}, ...]}
    """
    data    = request.get_json(force=True, silent=True) or {}
    company = data.get("company", "").lower()
    if company not in ("acoustical", "architectural"):
        return jsonify({"status": "error", "error": "company must be 'acoustical' or 'architectural'"}), 400

    try:
        items = com.get_all_items(company)
        return jsonify({"status": "ok", "count": len(items), "items": items})
    except RuntimeError as e:
        return jsonify({"status": "error", "error": str(e)}), 422
    except Exception as e:
        return jsonify({"status": "error", "error": f"Unexpected error: {e}"}), 500


@bp.post("/debug-invoice")
@require_api_key
def debug_invoice():
    """Return the generated qbXML without submitting — for debugging parse errors."""
    data = request.get_json(force=True, silent=True)
    if not data:
        return jsonify({"status": "error", "error": "Invalid JSON"}), 400
    invoice = data.get("invoice")
    if not invoice:
        return jsonify({"status": "error", "error": "invoice payload is required"}), 400
    try:
        qbxml = build_invoice_add(invoice)
        return jsonify({"status": "ok", "qbxml": qbxml})
    except Exception as e:
        return jsonify({"status": "error", "error": str(e)}), 500


@bp.post("/ping")
@require_api_key
def ping():
    """
    Verify QB is running and that the correct company file is open.

    Optional field:
        company — "acoustical" or "architectural". If provided, verifies
                  the open company file matches before returning ok.

    Returns:
        {"status": "ok",    "company_name": "...", "company_file": "..."}
        {"status": "error", "error": "..."}
    """
    data    = request.get_json(force=True, silent=True) or {}
    company = data.get("company", "").lower()

    try:
        info = com.get_open_company()

        if company in ("acoustical", "architectural"):
            com.verify_company(info, company)

        return jsonify({
            "status": "ok",
            "company_name": info["name"],
            "company_file": info["file"],
        })
    except RuntimeError as e:
        return jsonify({"status": "error", "error": str(e)}), 422
    except Exception as e:
        return jsonify({"status": "error", "error": f"Unexpected error: {e}"}), 500
