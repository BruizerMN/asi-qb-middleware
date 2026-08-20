"""
REST endpoints for FileMaker.

FileMaker uses Insert from URL to call these. All requests must include:
    X-API-Key: <shared secret from .env>

Endpoints:
    POST /fm/invoice        — submit an invoice to QuickBooks (synchronous via COM)
    POST /fm/sales-order    — submit a sales order to QuickBooks (synchronous via COM)
    POST /fm/delete-transaction — permanently delete a Sales Order/Invoice by RefNumber (dev/reset tool)
    POST /fm/ping           — verify QB is running and check which company is open
    POST /fm/sync-customers — return all active QB customers for FM to match and store
    POST /fm/sync-items     — return all active QB items for FM to match and store
    POST /fm/sync-customer  — sync a single customer by AccountNumber
    POST /fm/warm-customer-cache — refresh the customer-list cache for the open company, no other side effects (dev only)
    POST /fm/customer-updatecreate — create or update a single customer by AccountNumber
    POST /fm/sync-item      — sync a single item by Name (= FM productID)
    POST /fm/debug-invoice  — return generated qbXML without submitting (dev only)
    GET  /fm/view-invoice   — render a posted invoice as HTML
    GET  /fm/view-sales-order — render a posted sales order as HTML
    GET  /fm/debug-query-raw — return raw qbXML for an existing invoice/SO (dev only)
    GET  /fm/debug-reactivate-customer — one-off diagnostic: does IsActive=true undelete? (dev only)
"""

import html as _html
from functools import wraps
from flask import Blueprint, request, jsonify

from middleware.config import API_KEY
from middleware.qbxml_builder import build_invoice_add, build_sales_order_add
from middleware import logger
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

    _log  = {"order_id": order_id, "company": company, "status": "error"}
    _qbxml = ""

    try:
        # Ensure Customer:Job exists before submitting — QB won't auto-create it.
        # ensure_customer_job returns the name to use in InvoiceAdd: normally the
        # same as customer_name, but corrected if a stale FM name was detected and
        # the customer was found via ListID fallback.
        customer_list_id    = invoice.get("customer_list_id", "")
        customer_account_no = invoice.get("customer_account_number", "")

        # If QB_CustomerID wasn't populated in FM (empty customer_list_id), look up
        # the customer's QB ListID fresh from QB using the account number.
        # Uses the same proven path as the sync: bypass cache, match by AccountNumber.
        if not customer_list_id and customer_account_no:
            customer_data = com.get_customer_by_account(
                customer_account_no, bypass_cache=True
            )
            if customer_data:
                customer_list_id = customer_data.get("list_id", "")

        try:
            corrected_name, job_list_id = com.ensure_customer_job(
                invoice.get("customer_name", ""), company, customer_list_id
            )
        except RuntimeError as e:
            list_id_debug = (
                "list_id=MISSING" if "customer_list_id" not in invoice
                else f"list_id={customer_list_id!r}"
            )
            acct_debug = (
                "acct=MISSING" if "customer_account_number" not in invoice
                else f"acct={customer_account_no!r}"
            )
            raise RuntimeError(
                f"{e} | route debug: {list_id_debug} {acct_debug}"
            )

        # Apply any corrections ensure_customer_job made.
        # customer_job_list_id lets build_invoice_add() use CustomerRef/ListID
        # instead of FullName, bypassing apostrophe-encoding mismatches (error 3140).
        invoice = dict(invoice)
        if corrected_name != invoice.get("customer_name", ""):
            invoice["customer_name"] = corrected_name
        if job_list_id:
            invoice["customer_job_list_id"] = job_list_id

        qbxml  = build_invoice_add(invoice)
        _qbxml = qbxml  # capture for error logging before submitting

        qb_invoice_id, warning = com.submit_invoice(qbxml, company, invoice.get("ship_date", ""))
        _log.update({"status": "ok", "qb_invoice_id": qb_invoice_id})
        if warning:
            _log["warning"] = warning
        response = {
            "status": "ok",
            "qb_invoice_id": qb_invoice_id,
            "order_id": str(order_id),
        }
        if warning:
            response["warning"] = warning
        return jsonify(response)
    except RuntimeError as e:
        _log["error"] = str(e)
        if _qbxml:
            logger.log_invoice_error(str(order_id), company, _qbxml, str(e))
        return jsonify({"status": "error", "error": str(e)}), 422
    except Exception as e:
        _log["error"] = f"Unexpected error: {e}"
        if _qbxml:
            logger.log_invoice_error(str(order_id), company, _qbxml, str(e))
        return jsonify({"status": "error", "error": f"Unexpected error: {e}"}), 500
    finally:
        logger.log_event("invoice_post", **_log)


@bp.post("/sales-order")
@require_api_key
def post_sales_order():
    """
    FM posts a JSON sales order payload here. Submits synchronously to QB via COM.

    Mirrors post_invoice() exactly, just builds/submits a SalesOrderAdd instead
    of an InvoiceAdd. Added 2026-07-28 per Cat's request — same flow as AE (QB
    Desktop Sales Order, converted to an Invoice by Cat's team inside QB
    afterward). Not yet tested against a live QB Desktop session.

    Required fields:
        company     — "acoustical" or "architectural"
        order_id    — FM order ID (included in response for FM to match)
        sales_order — sales order dict (see qbxml_builder.py build_invoice_add
                      docstring for schema — build_sales_order_add takes the
                      same shape)

    Returns:
        {"status": "ok",    "qb_so_number": "1042", "order_id": "..."}
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

    sales_order = data.get("sales_order")
    if not sales_order:
        return jsonify({"status": "error", "error": "sales_order payload is required"}), 400

    _log  = {"order_id": order_id, "company": company, "status": "error"}
    _qbxml = ""

    try:
        # Ensure Customer:Job exists before submitting — same requirement as
        # InvoiceAdd; QB won't auto-create it for SalesOrderAdd either.
        customer_list_id    = sales_order.get("customer_list_id", "")
        customer_account_no = sales_order.get("customer_account_number", "")

        if not customer_list_id and customer_account_no:
            customer_data = com.get_customer_by_account(
                customer_account_no, bypass_cache=True
            )
            if customer_data:
                customer_list_id = customer_data.get("list_id", "")

        try:
            corrected_name, job_list_id = com.ensure_customer_job(
                sales_order.get("customer_name", ""), company, customer_list_id
            )
        except RuntimeError as e:
            list_id_debug = (
                "list_id=MISSING" if "customer_list_id" not in sales_order
                else f"list_id={customer_list_id!r}"
            )
            acct_debug = (
                "acct=MISSING" if "customer_account_number" not in sales_order
                else f"acct={customer_account_no!r}"
            )
            raise RuntimeError(
                f"{e} | route debug: {list_id_debug} {acct_debug}"
            )

        sales_order = dict(sales_order)
        if corrected_name != sales_order.get("customer_name", ""):
            sales_order["customer_name"] = corrected_name
        if job_list_id:
            sales_order["customer_job_list_id"] = job_list_id

        qbxml  = build_sales_order_add(sales_order)
        _qbxml = qbxml  # capture for error logging before submitting

        qb_so_number, warning = com.submit_sales_order(qbxml, company, sales_order.get("ship_date", ""))
        _log.update({"status": "ok", "qb_so_number": qb_so_number})
        if warning:
            _log["warning"] = warning
        response = {
            "status": "ok",
            "qb_so_number": qb_so_number,
            "order_id": str(order_id),
        }
        if warning:
            response["warning"] = warning
        return jsonify(response)
    except RuntimeError as e:
        _log["error"] = str(e)
        if _qbxml:
            logger.log_invoice_error(str(order_id), company, _qbxml, str(e))
        return jsonify({"status": "error", "error": str(e)}), 422
    except Exception as e:
        _log["error"] = f"Unexpected error: {e}"
        if _qbxml:
            logger.log_invoice_error(str(order_id), company, _qbxml, str(e))
        return jsonify({"status": "error", "error": f"Unexpected error: {e}"}), 500
    finally:
        logger.log_event("sales_order_post", **_log)


@bp.post("/delete-transaction")
@require_api_key
def delete_transaction():
    """
    Permanently delete a Sales Order or Invoice from QuickBooks by RefNumber
    (Bill's design, 2026-08-20 -- a discrete "full reset" tool, separate
    from the existing FM unlink tool). Auto-detects SalesOrder vs Invoice --
    caller doesn't need to know which type it is.

    Refuses to delete a transaction that has any linked transaction (most
    commonly a Sales Order already converted to an Invoice) -- QuickBooks
    would likely block this anyway, but with a much less actionable error.

    Irreversible once it succeeds. FM is responsible for confirming with the
    user before calling this.

    Required field:
        ref_number — the ASI order number as stored in QB_InvoiceID (e.g. "ASI-113889")

    Optional field:
        company — "acoustical" or "architectural". If omitted, auto-detected from
                  the open QB company file.

    Returns:
        {"status": "ok", "action": "deleted", "txn_type": "SalesOrder", "ref_number": "...", "txn_id": "..."}
        {"status": "ok", "action": "not_found", "ref_number": "..."}
        {"status": "blocked", "action": "blocked_linked_txns", "txn_type": "...", "ref_number": "...",
         "linked": [{"txn_type": "...", "ref_number": "..."}, ...]}
        {"status": "error", "error": "..."}
    """
    data       = request.get_json(force=True, silent=True) or {}
    ref_number = data.get("ref_number", "").strip()
    company    = data.get("company", "").lower()
    if not ref_number:
        return jsonify({"status": "error", "error": "ref_number is required"}), 400

    _log = {"ref_number": ref_number, "company": company or "auto-detect", "status": "error"}

    try:
        if company not in ("acoustical", "architectural"):
            company = com.detect_open_slug()
        _log["company"] = company

        result = com.delete_transaction(ref_number, company)
        _log.update({"status": result["action"], **result})

        if result["action"] == "blocked_linked_txns":
            return jsonify({"status": "blocked", **result})
        return jsonify({"status": "ok", **result})
    except RuntimeError as e:
        _log["error"] = str(e)
        return jsonify({"status": "error", "error": str(e)}), 422
    except Exception as e:
        _log["error"] = f"Unexpected error: {e}"
        return jsonify({"status": "error", "error": f"Unexpected error: {e}"}), 500
    finally:
        logger.log_event("delete_transaction", **_log)


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
    _log    = {"company": company or "auto-detect", "status": "error"}

    try:
        if company not in ("acoustical", "architectural"):
            company = com.detect_open_slug()
        _log["company"] = company
        customers = com.get_all_customers(company)
        lookup = {c["account_number"]: {"list_id": c["list_id"], "full_name": c["full_name"]} for c in customers}
        _log.update({"status": "ok", "count": len(customers)})
        return jsonify({"status": "ok", "count": len(customers), "lookup": lookup})
    except RuntimeError as e:
        _log["error"] = str(e)
        return jsonify({"status": "error", "error": str(e)}), 422
    except Exception as e:
        _log["error"] = f"Unexpected error: {e}"
        return jsonify({"status": "error", "error": f"Unexpected error: {e}"}), 500
    finally:
        logger.log_event("sync_customers", **_log)


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
    _log    = {"company": company or "auto-detect", "status": "error"}

    try:
        if company not in ("acoustical", "architectural"):
            company = com.detect_open_slug()
        _log["company"] = company
        items  = com.get_all_items(company)
        lookup = {item["name"]: item["list_id"] for item in items}
        _log.update({"status": "ok", "count": len(items)})
        return jsonify({"status": "ok", "count": len(items), "lookup": lookup})
    except RuntimeError as e:
        _log["error"] = str(e)
        return jsonify({"status": "error", "error": str(e)}), 422
    except Exception as e:
        _log["error"] = f"Unexpected error: {e}"
        return jsonify({"status": "error", "error": f"Unexpected error: {e}"}), 500
    finally:
        logger.log_event("sync_items", **_log)


@bp.post("/sync-customer")
@require_api_key
def sync_customer():
    """
    Sync a single customer by AccountNumber.

    Required field:
        account_number — FM customerID value (e.g. "C-107391")

    Returns:
        {"status": "ok",        "customer": {"list_id": "...", "full_name": "...", "account_number": "...", "is_active": true}}
        {"status": "inactive",  "customer": {...same shape, "is_active": false}}
        {"status": "not_found", "account_number": "..."}
        {"status": "error",     "error": "..."}

    "inactive" means the customer already exists in QB (found by AccountNumber)
    but is marked inactive there -- distinct from "not_found" so FM can tell the
    user to reactivate it in QB rather than looping "not synced" forever with
    no way to tell the two apart.
    """
    data           = request.get_json(force=True, silent=True) or {}
    account_number = data.get("account_number", "").strip()
    if not account_number:
        return jsonify({"status": "error", "error": "account_number is required"}), 400

    _log = {"account_number": account_number, "status": "error"}

    try:
        # bypass_cache=True: a sync must always fetch current data from QB.
        # The cache exists for performance on the invoice path — not for syncs.
        customer = com.get_customer_by_account(account_number, bypass_cache=True)
        if customer is None:
            _log["status"] = "not_found"
            return jsonify({"status": "not_found", "account_number": account_number})
        if not customer.get("is_active", True):
            _log.update({"status": "inactive", "list_id": customer.get("list_id", "")})
            return jsonify({"status": "inactive", "customer": customer})
        _log.update({"status": "ok", "list_id": customer.get("list_id", "")})
        return jsonify({"status": "ok", "customer": customer})
    except RuntimeError as e:
        _log["error"] = str(e)
        return jsonify({"status": "error", "error": str(e)}), 422
    except Exception as e:
        _log["error"] = f"Unexpected error: {e}"
        return jsonify({"status": "error", "error": f"Unexpected error: {e}"}), 500
    finally:
        logger.log_event("sync_customer", **_log)


@bp.post("/warm-customer-cache")
@require_api_key
def warm_customer_cache():
    """
    Dev tool (dev only): force-refresh the customer-list cache for whichever
    QB company is currently open. Read-only against QB -- creates, updates,
    or deletes nothing. Exists so a tech can manually reset the cache from FM
    (e.g. before timing a customer-creation test) without needing to touch a
    real customer record just to warm it.

    No required fields.

    Returns:
        {"status": "ok", "company": "acoustical", "company_name": "...", "count": N}
        {"status": "error", "error": "..."}
    """
    _log = {"status": "error"}
    try:
        result = com.warm_customer_cache()
        _log.update({"status": "ok", "company": result["slug"], "count": result["count"]})
        return jsonify({
            "status":       "ok",
            "company":      result["slug"],
            "company_name": result["company_name"],
            "count":        result["count"],
        })
    except RuntimeError as e:
        _log["error"] = str(e)
        return jsonify({"status": "error", "error": str(e)}), 422
    except Exception as e:
        _log["error"] = f"Unexpected error: {e}"
        return jsonify({"status": "error", "error": f"Unexpected error: {e}"}), 500
    finally:
        logger.log_event("warm_customer_cache", **_log)


@bp.post("/customer-updatecreate")
@require_api_key
def customer_updatecreate():
    """
    Create the QB customer if it doesn't exist (matched by AccountNumber), or
    update it (CustomerMod) if it does.

    Field set is table-driven -- see qbxml_builder.CUSTOMER_FIELD_MAP for the
    authoritative current list (as of 2026-08-11: company_name, bill_addr1-4,
    bill_city, bill_state, bill_zip, email) and how to add more without a
    middleware code change (Bill, 2026-08-11: Cat's team will keep sending
    more fields over time, potentially "an hour from now or even days" apart
    -- this table exists specifically so that doesn't mean another deploy
    each time, as long as the new field is a standard QB Customer field).

    Required fields:
        account_number    — FM customerID value (e.g. "C-18042")
        fields.company_name — FM customerCompany_NEW value, used for QB Name + CompanyName
                            (NOT customerCompany -- that field is a stale legacy
                            duplicate as of 2026-08-11, confirmed by Bill)
    Optional:
        existing_list_id  — FM's currently-stored QB_CustomerID, if any. Enables
                             the stale-link check (see com_handler.create_or_update_customer).
        fields.*          — any other CUSTOMER_FIELD_MAP key; absent/empty
                             values are simply not sent to QB.

    Returns:
        {"status": "created", "customer": {"list_id":..., "full_name":..., "account_number":...}, "stale_link": {...} or null}
        {"status": "updated",  "customer": {...same shape}, "stale_link": {...} or null}
        {"status": "duplicate_name", "conflict": {"list_id":..., "full_name":..., "account_number":...}, "qb_message": "...", "stale_link": {...} or null}
        {"status": "error",    "error": "..."}

    "duplicate_name" means QB already has a different record using the same
    Name (error 3100) -- "conflict" identifies that record if it's a
    Customer/Job FM can see, but QB actually enforces Name uniqueness across
    Customers, Vendors, Employees, and Other Names together, so "conflict"
    may come back empty if the collision is with one of those other list
    types (confirmed to happen live, 2026-08-11). "qb_message" always carries
    QuickBooks' own statusMessage text regardless, since it may say more than
    our own lookup can determine.

    "stale_link" is non-null when existing_list_id was provided but no longer
    matches account_number in QB -- the link was cleared and normal matching
    ran fresh instead. Every response (including errors) is logged with the
    full step-by-step "trace" from com_handler -- Bill, 2026-08-11: verbose
    logging here on purpose, "if it were to go sideways" this is the one
    place to find out why.
    """
    data             = request.get_json(force=True, silent=True) or {}
    account_number   = data.get("account_number", "").strip()
    fields           = data.get("fields") or {}
    existing_list_id = (data.get("existing_list_id") or "").strip()
    company_name     = (fields.get("company_name") or "").strip()
    if not account_number or not company_name:
        return jsonify({"status": "error", "error": "account_number and fields.company_name are required"}), 400

    _log = {
        "account_number": account_number,
        "fields": fields,
        "existing_list_id": existing_list_id,
        "status": "error",
    }

    try:
        result = com.create_or_update_customer(account_number, fields, existing_list_id)
        _log["trace"] = result.get("trace", [])
        _log["stale_link"] = result.get("stale_link")
        if result["action"] == "duplicate_name":
            _log.update({
                "status": "duplicate_name",
                "conflict_account_number": result["conflict"].get("account_number", ""),
                "qb_message": result.get("qb_message", ""),
            })
            return jsonify({
                "status": "duplicate_name",
                "conflict": result["conflict"],
                "qb_message": result.get("qb_message", ""),
                "stale_link": result.get("stale_link"),
            })
        _log.update({"status": result["action"], "list_id": result["customer"].get("list_id", "")})
        return jsonify({
            "status": result["action"],
            "customer": result["customer"],
            "stale_link": result.get("stale_link"),
        })
    except RuntimeError as e:
        _log["error"] = str(e)
        return jsonify({"status": "error", "error": str(e)}), 422
    except Exception as e:
        _log["error"] = f"Unexpected error: {e}"
        return jsonify({"status": "error", "error": f"Unexpected error: {e}"}), 500
    finally:
        logger.log_event("customer_updatecreate", **_log)


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

    _log = {"item_name": item_name, "status": "error"}

    try:
        item = com.get_item_by_name(item_name)
        if item is None:
            _log["status"] = "not_found"
            return jsonify({"status": "not_found", "item_name": item_name})
        _log.update({"status": "ok", "list_id": item.get("list_id", "")})
        return jsonify({"status": "ok", "item": item})
    except RuntimeError as e:
        _log["error"] = str(e)
        return jsonify({"status": "error", "error": str(e)}), 422
    except Exception as e:
        _log["error"] = f"Unexpected error: {e}"
        return jsonify({"status": "error", "error": f"Unexpected error: {e}"}), 500
    finally:
        logger.log_event("sync_item", **_log)


@bp.post("/ping-customer")
@require_api_key
def ping_customer():
    """
    Diagnostic: look up a single customer by QB ListID and return what QB says.
    Useful for verifying the ListID stored in FM is still valid in QB.

    POST body: { "company": "acoustical"|"architectural", "list_id": "8000XXXX-..." }
    Returns full_name, account_number, is_active, and raw QB status info.
    """
    data = request.get_json(force=True, silent=True) or {}
    list_id = data.get("list_id", "").strip()
    if not list_id:
        return jsonify({"status": "error", "error": "list_id is required"}), 400
    try:
        result = com.lookup_customer_by_list_id(list_id)
        return jsonify({"status": "ok", "result": result})
    except RuntimeError as e:
        return jsonify({"status": "error", "error": str(e)}), 422
    except Exception as e:
        return jsonify({"status": "error", "error": f"Unexpected error: {e}"}), 500


@bp.post("/validate-rep")
@require_api_key
def validate_rep():
    """Validate that a sales rep name (initials) exists in QB."""
    data     = request.get_json(force=True, silent=True) or {}
    rep_name = (data.get("rep_name") or "").strip()
    if not rep_name:
        return jsonify({"status": "error", "error": "rep_name is required"}), 400

    _log = {"rep_name": rep_name, "status": "error"}

    try:
        reps     = com.get_all_sales_reps()
        initials = {r["initials"].upper() for r in reps}
        if rep_name.upper() in initials:
            _log["status"] = "ok"
            return jsonify({"status": "ok"})
        _log["error"] = f"Sales rep '{rep_name}' not found in QuickBooks."
        return jsonify({"status": "error", "error": _log["error"]}), 422
    except RuntimeError as e:
        _log["error"] = str(e)
        return jsonify({"status": "error", "error": str(e)}), 422
    except Exception as e:
        _log["error"] = f"Unexpected error: {e}"
        return jsonify({"status": "error", "error": f"Unexpected error: {e}"}), 500
    finally:
        logger.log_event("validate_rep", **_log)


@bp.post("/debug-invoice")
@require_api_key
def debug_invoice():
    """Return the generated qbXML without submitting — for debugging parse errors.

    Body: {"invoice": {...}} for InvoiceAdd (default), or
          {"sales_order": {...}} for SalesOrderAdd -- same payload shape,
          just routes to build_sales_order_add() instead. Added 2026-08-10
          because this endpoint predates the SalesOrder path and only ever
          built InvoiceAdd XML, even though production posts go through
          SalesOrderAdd."""
    data = request.get_json(force=True, silent=True)
    if not data:
        return jsonify({"status": "error", "error": "Invalid JSON"}), 400
    if "sales_order" in data:
        payload = data.get("sales_order")
        builder = build_sales_order_add
        kind = "sales_order"
    else:
        payload = data.get("invoice")
        builder = build_invoice_add
        kind = "invoice"
    if not payload:
        return jsonify({"status": "error", "error": "invoice or sales_order payload is required"}), 400
    try:
        qbxml = builder(payload)
        return jsonify({"status": "ok", "kind": kind, "qbxml": qbxml})
    except Exception as e:
        return jsonify({"status": "error", "error": str(e)}), 500


@bp.get("/view-invoice")
def view_invoice():
    """
    Render a QB invoice as HTML, styled to match QB Desktop's invoice form.

    Query params:
        invoice_id  — QB invoice number (FM QB_InvoiceID field)
        company     — "acoustical" or "architectural"
        api_key     — shared API key (query string since GET can't send custom headers)
    """
    if request.args.get("api_key", "") != API_KEY:
        return "<h1>Unauthorized</h1>", 401

    invoice_id = request.args.get("invoice_id", "").strip()
    company    = request.args.get("company", "").lower()

    if not invoice_id:
        return "<h1>Missing invoice_id parameter</h1>", 400
    if company not in ("acoustical", "architectural"):
        return "<h1>Missing or invalid company parameter</h1>", 400

    try:
        inv = com.get_invoice(invoice_id, company)
        return _render_invoice_html(inv)
    except RuntimeError as e:
        return f"<h1 style='font-family:Arial'>Error</h1><p style='font-family:Arial'>{_html.escape(str(e))}</p>", 422
    except Exception as e:
        return f"<h1 style='font-family:Arial'>Unexpected Error</h1><p style='font-family:Arial'>{_html.escape(str(e))}</p>", 500


@bp.get("/view-sales-order")
def view_sales_order():
    """
    Render a QB sales order as HTML, styled to match QB Desktop's SO form.

    Query params:
        so_number — QB sales order number (FM QB_InvoiceID field)
        company   — "acoustical" or "architectural"
        api_key   — shared API key (query string since GET can't send custom headers)
    """
    if request.args.get("api_key", "") != API_KEY:
        return "<h1>Unauthorized</h1>", 401

    so_number = request.args.get("so_number", "").strip()
    company   = request.args.get("company", "").lower()

    if not so_number:
        return "<h1>Missing so_number parameter</h1>", 400
    if company not in ("acoustical", "architectural"):
        return "<h1>Missing or invalid company parameter</h1>", 400

    try:
        so = com.get_sales_order(so_number, company)
        return _render_sales_order_html(so)
    except RuntimeError as e:
        return f"<h1 style='font-family:Arial'>Error</h1><p style='font-family:Arial'>{_html.escape(str(e))}</p>", 422
    except Exception as e:
        return f"<h1 style='font-family:Arial'>Unexpected Error</h1><p style='font-family:Arial'>{_html.escape(str(e))}</p>", 500


@bp.get("/debug-reactivate-customer")
def debug_reactivate_customer():
    """
    DIAGNOSTIC ONLY, not used by any live FM script or production path.
    Built 2026-08-12 to answer one specific question empirically: does
    CustomerMod's IsActive=true also resolve QB's "deleted" state (distinct
    from plain inactive -- confirmed by Bill, 2026-08-12: QB Desktop shows a
    red-X marker and a separate "would you like to undelete it?" prompt for
    deleted customers), or does it only reactivate plain-inactive ones?

    Query params:
        list_id — QB ListID to test against (e.g. a customer currently
                  showing the red-X "deleted" marker in QB Desktop)
        api_key — shared API key (query string since GET can't send custom headers)
    """
    if request.args.get("api_key", "") != API_KEY:
        return "<h1>Unauthorized</h1>", 401

    list_id = request.args.get("list_id", "").strip()
    if not list_id:
        return jsonify({"status": "error", "error": "list_id is required"}), 400

    try:
        result = com.diagnostic_reactivate_customer(list_id)
        return jsonify(result)
    except RuntimeError as e:
        return jsonify({"status": "error", "error": str(e)}), 422
    except Exception as e:
        return jsonify({"status": "error", "error": f"Unexpected error: {e}"}), 500


@bp.get("/debug-query-raw")
def debug_query_raw():
    """
    Return the RAW, unparsed qbXML response for an existing Invoice, Sales Order,
    or the full Sales Tax Item/Group list.

    One-off diagnostic tool -- not used by FM or any live posting path. Purpose:
    inspect exactly how QuickBooks represents sales tax (both on a real transaction
    and in its Sales Tax Item list) before wiring up automated tax handling
    (2026-08-07).

    Query params:
        type        — "invoice", "sales-order", or "sales-tax-items"
        ref_number  — QB invoice # or SO # (required for invoice/sales-order;
                      ignored for sales-tax-items, which is a list query)
        company     — "acoustical" or "architectural"
        api_key     — shared API key (query string since GET can't send custom headers)
    """
    if request.args.get("api_key", "") != API_KEY:
        return "<h1>Unauthorized</h1>", 401

    kind       = request.args.get("type", "").strip().lower()
    ref_number = request.args.get("ref_number", "").strip()
    company    = request.args.get("company", "").lower()

    if kind not in ("invoice", "sales-order", "sales-tax-items"):
        return "<h1>type must be 'invoice', 'sales-order', or 'sales-tax-items'</h1>", 400
    if kind in ("invoice", "sales-order") and not ref_number:
        return "<h1>Missing ref_number parameter</h1>", 400
    if company not in ("acoustical", "architectural"):
        return "<h1>Missing or invalid company parameter</h1>", 400

    try:
        raw = com.get_raw_query_response(kind, ref_number, company)
        return raw, 200, {"Content-Type": "text/xml; charset=utf-8"}
    except RuntimeError as e:
        return f"<h1>Error</h1><p>{_html.escape(str(e))}</p>", 422
    except Exception as e:
        return f"<h1>Unexpected Error</h1><p>{_html.escape(str(e))}</p>", 500


def _render_invoice_html(inv: dict) -> str:
    """Render a QB invoice dict as an HTML page styled like QB Desktop."""
    h = _html.escape

    def fmt_date(d):
        if d and len(d) == 10:
            return f"{d[5:7]}/{d[8:10]}/{d[:4]}"
        return d or ""

    def fmt_money(v):
        try:
            return f"{float(v):,.2f}"
        except (ValueError, TypeError):
            return str(v) if v else "0.00"

    def fmt_addr(a):
        parts = [a.get("addr1",""), a.get("addr2",""), a.get("addr3","")]
        parts = [p for p in parts if p]
        csz = " ".join(filter(None, [a.get("city",""), a.get("state",""), a.get("zip","")]))
        if csz:
            parts.append(csz)
        return "<br>".join(h(p) for p in parts)

    # Line item rows
    line_rows = ""
    for i, li in enumerate(inv["line_items"]):
        bg = "#EAF3FB" if i % 2 == 1 else "#FFFFFF"
        desc = h(li["description"]).replace("\n", "<br>")
        line_rows += f"""
        <tr style="background:{bg}">
          <td>{h(li['uom'])}</td>
          <td>{h(li['item'])}</td>
          <td class="num">{h(li['quantity'])}</td>
          <td class="desc">{desc}</td>
          <td class="num">{fmt_money(li['rate'])}</td>
          <td class="num">{fmt_money(li['amount'])}</td>
        </tr>"""

    memo_html = (
        f'<div class="memo"><strong>Memo:</strong> {h(inv["memo"])}</div>'
        if inv.get("memo") else ""
    )

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Invoice #{h(inv['ref_number'])}</title>
  <style>
    *{{box-sizing:border-box;margin:0;padding:0}}
    body{{font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#333;background:#f0f0f0;padding:24px 28px}}
    .page{{max-width:1060px;margin:0 auto;background:#fff;padding:24px 28px;box-shadow:0 1px 4px rgba(0,0,0,.15)}}
    .inv-title{{font-size:34px;font-weight:bold;color:#111;line-height:1}}
    .hdr{{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:16px}}
    .hdr-right{{display:flex;gap:14px}}
    .hdr-block{{min-width:130px}}
    .hdr-block.wide{{min-width:180px}}
    .lbl{{font-size:9px;color:#666;text-transform:uppercase;letter-spacing:.5px;font-weight:bold;margin-bottom:2px}}
    .box{{border:1px solid #c0c0c0;background:#f9f9f9;padding:5px 8px;min-height:26px;font-size:12px;line-height:1.45}}
    .box.addr{{min-height:64px}}
    .fields-bar{{display:flex;border:1px solid #c0c0c0;border-bottom:none}}
    .fc{{flex:1;padding:5px 10px;border-right:1px solid #c0c0c0}}
    .fc:last-child{{border-right:none}}
    .fc.wide{{flex:2}}
    table{{width:100%;border-collapse:collapse;border:1px solid #c0c0c0}}
    thead tr{{background:#edf1f5}}
    th{{padding:5px 8px;font-size:9px;text-transform:uppercase;color:#555;letter-spacing:.4px;border-right:1px solid #c0c0c0;border-bottom:2px solid #b0c4d4;font-weight:bold;text-align:left;white-space:nowrap}}
    th:last-child{{border-right:none}}
    td{{padding:5px 8px;vertical-align:top;border-right:1px solid #dce8f0}}
    td:last-child{{border-right:none}}
    tr{{border-bottom:1px solid #dce8f0}}
    td.num{{text-align:right;white-space:nowrap}}
    td.desc{{white-space:pre-wrap}}
    .footer{{display:flex;justify-content:flex-end;padding:10px 14px;border:1px solid #c0c0c0;border-top:none}}
    .totals{{width:270px}}
    .tr{{display:flex;justify-content:space-between;padding:3px 0;font-size:12px}}
    .tr.bal{{font-weight:bold;font-size:14px;border-top:2px solid #333;margin-top:5px;padding-top:6px}}
    .memo{{font-size:11px;color:#666;margin-top:10px;padding:6px 8px;border-left:3px solid #b0c4d4}}
  </style>
</head>
<body>
<div class="page">

<div class="hdr">
  <div class="inv-title">Invoice</div>
  <div class="hdr-right">
    <div class="hdr-block">
      <div class="lbl">Date</div>
      <div class="box">{fmt_date(inv['txn_date'])}</div>
      <div class="lbl" style="margin-top:8px">Invoice #</div>
      <div class="box">{h(inv['ref_number'])}</div>
    </div>
    <div class="hdr-block wide">
      <div class="lbl">Bill To</div>
      <div class="box addr">{fmt_addr(inv['bill_address'])}</div>
    </div>
    <div class="hdr-block wide">
      <div class="lbl">Ship To</div>
      <div class="box addr">{fmt_addr(inv['ship_address'])}</div>
    </div>
  </div>
</div>

<div class="fields-bar">
  <div class="fc wide"><div class="lbl">Customer · Job</div><div>{h(inv['customer'])}</div></div>
  <div class="fc"><div class="lbl">P.O. Number</div><div>{h(inv['po_number'])}</div></div>
  <div class="fc"><div class="lbl">Terms</div><div>{h(inv['terms'])}</div></div>
  <div class="fc"><div class="lbl">Rep</div><div>{h(inv['rep'])}</div></div>
  <div class="fc"><div class="lbl">Via</div><div>{h(inv['ship_method'])}</div></div>
  <div class="fc"><div class="lbl">Class</div><div>{h(inv['class_name'])}</div></div>
</div>

<table>
  <thead>
    <tr>
      <th style="width:60px">U/M</th>
      <th style="width:130px">Item Code</th>
      <th style="width:75px;text-align:right">Quantity</th>
      <th>Description</th>
      <th style="width:95px;text-align:right">Price Each</th>
      <th style="width:95px;text-align:right">Amount</th>
    </tr>
  </thead>
  <tbody>{line_rows}
  </tbody>
</table>

<div class="footer">
  <div class="totals">
    <div class="tr"><span>Total</span><span>{fmt_money(inv['total'])}</span></div>
    <div class="tr"><span>Payments Applied</span><span>{fmt_money(inv['applied'])}</span></div>
    <div class="tr bal"><span>Balance Due</span><span>{fmt_money(inv['balance'])}</span></div>
  </div>
</div>

{memo_html}
</div>
</body>
</html>"""


def _render_sales_order_html(so: dict) -> str:
    """Render a QB sales order dict as an HTML page styled like QB Desktop.

    Mirrors _render_invoice_html() — a SalesOrderRet has no AppliedAmount/
    BalanceRemaining (invoice/payment concepts that don't exist pre-conversion),
    so the footer shows Total only.
    """
    h = _html.escape

    def fmt_date(d):
        if d and len(d) == 10:
            return f"{d[5:7]}/{d[8:10]}/{d[:4]}"
        return d or ""

    def fmt_money(v):
        try:
            return f"{float(v):,.2f}"
        except (ValueError, TypeError):
            return str(v) if v else "0.00"

    def fmt_addr(a):
        parts = [a.get("addr1",""), a.get("addr2",""), a.get("addr3","")]
        parts = [p for p in parts if p]
        csz = " ".join(filter(None, [a.get("city",""), a.get("state",""), a.get("zip","")]))
        if csz:
            parts.append(csz)
        return "<br>".join(h(p) for p in parts)

    # Line item rows
    line_rows = ""
    for i, li in enumerate(so["line_items"]):
        bg = "#EAF3FB" if i % 2 == 1 else "#FFFFFF"
        desc = h(li["description"]).replace("\n", "<br>")
        line_rows += f"""
        <tr style="background:{bg}">
          <td>{h(li['uom'])}</td>
          <td>{h(li['item'])}</td>
          <td class="num">{h(li['quantity'])}</td>
          <td class="desc">{desc}</td>
          <td class="num">{fmt_money(li['rate'])}</td>
          <td class="num">{fmt_money(li['amount'])}</td>
        </tr>"""

    memo_html = (
        f'<div class="memo"><strong>Memo:</strong> {h(so["memo"])}</div>'
        if so.get("memo") else ""
    )

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Sales Order #{h(so['ref_number'])}</title>
  <style>
    *{{box-sizing:border-box;margin:0;padding:0}}
    body{{font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#333;background:#f0f0f0;padding:24px 28px}}
    .page{{max-width:1060px;margin:0 auto;background:#fff;padding:24px 28px;box-shadow:0 1px 4px rgba(0,0,0,.15)}}
    .inv-title{{font-size:34px;font-weight:bold;color:#111;line-height:1}}
    .hdr{{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:16px}}
    .hdr-right{{display:flex;gap:14px}}
    .hdr-block{{min-width:130px}}
    .hdr-block.wide{{min-width:180px}}
    .lbl{{font-size:9px;color:#666;text-transform:uppercase;letter-spacing:.5px;font-weight:bold;margin-bottom:2px}}
    .box{{border:1px solid #c0c0c0;background:#f9f9f9;padding:5px 8px;min-height:26px;font-size:12px;line-height:1.45}}
    .box.addr{{min-height:64px}}
    .fields-bar{{display:flex;border:1px solid #c0c0c0;border-bottom:none}}
    .fc{{flex:1;padding:5px 10px;border-right:1px solid #c0c0c0}}
    .fc:last-child{{border-right:none}}
    .fc.wide{{flex:2}}
    table{{width:100%;border-collapse:collapse;border:1px solid #c0c0c0}}
    thead tr{{background:#edf1f5}}
    th{{padding:5px 8px;font-size:9px;text-transform:uppercase;color:#555;letter-spacing:.4px;border-right:1px solid #c0c0c0;border-bottom:2px solid #b0c4d4;font-weight:bold;text-align:left;white-space:nowrap}}
    th:last-child{{border-right:none}}
    td{{padding:5px 8px;vertical-align:top;border-right:1px solid #dce8f0}}
    td:last-child{{border-right:none}}
    tr{{border-bottom:1px solid #dce8f0}}
    td.num{{text-align:right;white-space:nowrap}}
    td.desc{{white-space:pre-wrap}}
    .footer{{display:flex;justify-content:flex-end;padding:10px 14px;border:1px solid #c0c0c0;border-top:none}}
    .totals{{width:270px}}
    .tr{{display:flex;justify-content:space-between;padding:3px 0;font-size:12px}}
    .tr.bal{{font-weight:bold;font-size:14px;border-top:2px solid #333;margin-top:5px;padding-top:6px}}
    .memo{{font-size:11px;color:#666;margin-top:10px;padding:6px 8px;border-left:3px solid #b0c4d4}}
  </style>
</head>
<body>
<div class="page">

<div class="hdr">
  <div class="inv-title">Sales Order</div>
  <div class="hdr-right">
    <div class="hdr-block">
      <div class="lbl">Date</div>
      <div class="box">{fmt_date(so['txn_date'])}</div>
      <div class="lbl" style="margin-top:8px">S.O. #</div>
      <div class="box">{h(so['ref_number'])}</div>
    </div>
    <div class="hdr-block wide">
      <div class="lbl">Bill To</div>
      <div class="box addr">{fmt_addr(so['bill_address'])}</div>
    </div>
    <div class="hdr-block wide">
      <div class="lbl">Ship To</div>
      <div class="box addr">{fmt_addr(so['ship_address'])}</div>
    </div>
  </div>
</div>

<div class="fields-bar">
  <div class="fc wide"><div class="lbl">Customer · Job</div><div>{h(so['customer'])}</div></div>
  <div class="fc"><div class="lbl">P.O. Number</div><div>{h(so['po_number'])}</div></div>
  <div class="fc"><div class="lbl">Terms</div><div>{h(so['terms'])}</div></div>
  <div class="fc"><div class="lbl">Rep</div><div>{h(so['rep'])}</div></div>
  <div class="fc"><div class="lbl">Via</div><div>{h(so['ship_method'])}</div></div>
  <div class="fc"><div class="lbl">Class</div><div>{h(so['class_name'])}</div></div>
</div>

<table>
  <thead>
    <tr>
      <th style="width:60px">U/M</th>
      <th style="width:130px">Item Code</th>
      <th style="width:75px;text-align:right">Quantity</th>
      <th>Description</th>
      <th style="width:95px;text-align:right">Price Each</th>
      <th style="width:95px;text-align:right">Amount</th>
    </tr>
  </thead>
  <tbody>{line_rows}
  </tbody>
</table>

<div class="footer">
  <div class="totals">
    <div class="tr bal"><span>Total</span><span>{fmt_money(so['total'])}</span></div>
  </div>
</div>

{memo_html}
</div>
</body>
</html>"""


@bp.post("/list-ship-methods")
@require_api_key
def list_ship_methods():
    """Return all active QB shipping methods."""
    try:
        methods = com.get_all_ship_methods()
        return jsonify({"status": "ok", "count": len(methods), "ship_methods": methods})
    except RuntimeError as e:
        return jsonify({"status": "error", "error": str(e)}), 422
    except Exception as e:
        return jsonify({"status": "error", "error": f"Unexpected error: {e}"}), 500


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
