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

import html as _html
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
