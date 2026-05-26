"""
REST endpoints for FileMaker.

FileMaker uses Insert from URL to call these. All requests must include:
    X-API-Key: <shared secret from .env>

Endpoints:
    POST /fm/invoice        — submit an invoice to QuickBooks (synchronous via COM)
    POST /fm/ping           — verify QB is running and check which company is open
    POST /fm/sync-customers — return all active QB customers for FM to match and store
    POST /fm/sync-items     — return all active QB items for FM to match and store
    POST /fm/sync-customer  — sync a single customer by AccountNumber
    POST /fm/sync-item      — sync a single item by Name (= FM productID)
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
        # Ensure Customer:Job exists before submitting — QB won't auto-create it.
        com.ensure_customer_job(invoice.get("customer_name", ""), company)

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

    Optional field:
        company — "acoustical" or "architectural". If omitted, auto-detected from
                  the open QB company file.

    Returns:
        {"status": "ok", "count": N, "customers": [{"list_id": "...", "full_name": "...", "account_number": "..."}, ...]}
    """
    data    = request.get_json(force=True, silent=True) or {}
    company = data.get("company", "").lower()

    try:
        if company not in ("acoustical", "architectural"):
            company = com.detect_open_slug()
        customers = com.get_all_customers(company)
        lookup = {c["account_number"]: {"list_id": c["list_id"], "full_name": c["full_name"]} for c in customers}
        return jsonify({"status": "ok", "count": len(customers), "lookup": lookup})
    except RuntimeError as e:
        return jsonify({"status": "error", "error": str(e)}), 422
    except Exception as e:
        return jsonify({"status": "error", "error": f"Unexpected error: {e}"}), 500


@bp.post("/sync-items")
@require_api_key
def sync_items():
    """
    Return all active QB non-inventory items for FM to match against its PRODUCTS table.

    Optional field:
        company — "acoustical" or "architectural". If omitted, auto-detected from
                  the open QB company file.

    Returns:
        {"status": "ok", "count": N, "items": [{"list_id": "...", "name": "..."}, ...]}
    """
    data    = request.get_json(force=True, silent=True) or {}
    company = data.get("company", "").lower()

    try:
        if company not in ("acoustical", "architectural"):
            company = com.detect_open_slug()
        items = com.get_all_items(company)
        lookup = {item["name"]: item["list_id"] for item in items}
        return jsonify({"status": "ok", "count": len(items), "lookup": lookup})
    except RuntimeError as e:
        return jsonify({"status": "error", "error": str(e)}), 422
    except Exception as e:
        return jsonify({"status": "error", "error": f"Unexpected error: {e}"}), 500


@bp.post("/sync-customer")
@require_api_key
def sync_customer():
    """
    Sync a single customer by AccountNumber.

    Required field:
        account_number — FM customerID value (e.g. "C-107391")

    Returns:
        {"status": "ok",        "customer": {"list_id": "...", "full_name": "...", "account_number": "..."}}
        {"status": "not_found", "account_number": "..."}
        {"status": "error",     "error": "..."}
    """
    data           = request.get_json(force=True, silent=True) or {}
    account_number = data.get("account_number", "").strip()
    if not account_number:
        return jsonify({"status": "error", "error": "account_number is required"}), 400

    try:
        customer = com.get_customer_by_account(account_number)
        if customer is None:
            return jsonify({"status": "not_found", "account_number": account_number})
        return jsonify({"status": "ok", "customer": customer})
    except RuntimeError as e:
        return jsonify({"status": "error", "error": str(e)}), 422
    except Exception as e:
        return jsonify({"status": "error", "error": f"Unexpected error: {e}"}), 500


@bp.post("/sync-item")
@require_api_key
def sync_item():
    """
    Sync a single non-inventory item by Name (matches FM productID).

    Required field:
        item_name — FM productID value (e.g. "XABAB12448BK")

    Returns:
        {"status": "ok",        "item": {"list_id": "...", "name": "..."}}
        {"status": "not_found", "item_name": "..."}
        {"status": "error",     "error": "..."}
    """
    data      = request.get_json(force=True, silent=True) or {}
    item_name = data.get("item_name", "").strip()
    if not item_name:
        return jsonify({"status": "error", "error": "item_name is required"}), 400

    try:
        item = com.get_item_by_name(item_name)
        if item is None:
            return jsonify({"status": "not_found", "item_name": item_name})
        return jsonify({"status": "ok", "item": item})
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


@bp.post("/list-terms")
@require_api_key
def list_terms():
    """Return all active QB payment terms."""
    try:
        terms = com.get_all_terms()
        return jsonify({"status": "ok", "count": len(terms), "terms": terms})
    except RuntimeError as e:
        return jsonify({"status": "error", "error": str(e)}), 422
    except Exception as e:
        return jsonify({"status": "error", "error": f"Unexpected error: {e}"}), 500


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
