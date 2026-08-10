"""
Build qbXML InvoiceAdd requests from FM invoice data.

FM sends a JSON payload with the invoice. This module turns that into
the qbXML string that QuickBooks Web Connector feeds to QB Desktop.

Expected payload structure (all fields are strings unless noted):
{
    "customer_name": "Acme Corp",
    "po_number": "12345",             # required — actually the FM order ID, becomes QB RefNumber
                                       # (Ref No. / Invoice # / SO #). Misleadingly named for
                                       # historical reasons; NOT the customer's PO (see cust_po).
    "order_date": "2026-05-05",       # YYYY-MM-DD
    "ship_date": "2026-05-10",        # YYYY-MM-DD, optional -- writes QB's "Promise Date"
                                       # custom field (a Data Extension), NOT native ShipDate
    "cust_po": "PO-9876",             # optional — QB PONumber (P.O. No. field)
    "class_id": "80000001-1234567",   # optional — QB Class ListID for segment reporting
    "ship_via": "UPS Ground",         # optional
    "ship_to_name": "Acme Corp",
    "ship_to_addr1": "123 Main St",
    "ship_to_addr2": "",              # optional
    "ship_to_city": "Minneapolis",
    "ship_to_state": "MN",
    "ship_to_zip": "55401",
    "tax_amount": "45.00",            # pre-calculated by FM/Avalara
    "freight_amount": "12.50",
    "salesperson": "Jane Smith",      # optional — QB rep initials
    "bill_to_name":  "Acme Corp",       # optional — billing address
    "bill_to_addr1": "123 Main St",
    "bill_to_addr2": "",
    "bill_to_city":  "Minneapolis",
    "bill_to_state": "MN",
    "bill_to_zip":   "55401",
    "terms":     "Net 30",             # optional — must match QB Terms list exactly
    "memo": "PP Note text",           # optional
    "line_items": [
        {
            "item_name": "PROD-001",
            "description": "Acoustical panel 2x4",
            "quantity": "2",
            "unit_price": "250.00",
            "exclude": false          # bool — respects AE_Exclude flag
        }
    ]
}
"""

import re
import xml.etree.ElementTree as ET
from datetime import datetime


def _mmddyyyy(iso_date: str) -> str:
    """Convert an FM-supplied YYYY-MM-DD date to QB's MM/DD/YYYY display convention.

    Used only for the "Promise Date" custom field (a plain STR255TYPE data
    extension, not a native QB date type) -- QB doesn't reformat/validate it,
    so we match the format already used by every other Promise Date value
    already in the company file (e.g. "09/16/2026"), confirmed 2026-08-07 via
    a real SalesOrderQuery response. Falls back to the raw input on parse
    failure rather than crash the whole push over a display-only field."""
    try:
        return datetime.strptime(iso_date, "%Y-%m-%d").strftime("%m/%d/%Y")
    except ValueError:
        return iso_date


# XML 1.0 only permits tab (0x09), LF (0x0A), and CR (0x0D) from the C0
# control range -- every other 0x00-0x1F byte is illegal in an XML document,
# even though it's perfectly valid ASCII. These can slip into FM text fields
# (Memo, notes, addresses) via copy-paste from email/PDF sources and produce
# qbXML that QuickBooks' own parser rejects outright ("found an error when
# parsing the provided XML text stream"), unrelated to the ASCII-range replacements
# below. Root-caused 2026-08-10: Katie Anderson hit this on ASI-113748 right after
# Preflight passed; Bill reproduced the identical error independently on the same
# order, confirming it's a deterministic data issue, not a transient QB glitch.
_ILLEGAL_XML_CHARS = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f]')


def _ascii_safe(value: str) -> str:
    """Replace Unicode typographic characters with ASCII equivalents.
    QB Desktop's XML parser does not support non-ASCII characters.

    Also maps Windows-1252 C1 control characters (U+0080-U+009F) to their
    correct Unicode equivalents before converting to ASCII. QB Desktop is a
    Win32 application whose qbXML output may contain raw CP1252 bytes (e.g.
    0x92 for the right single quote) even when the XML header declares UTF-8.
    Python's XML parser reads those as C1 control characters, so we remap
    them here before the ASCII conversion step.

    Also collapses any embedded line break (CRLF, lone CR, or lone LF) into
    a single space. FM multi-line item descriptions (e.g. spec sheets with
    one attribute per line) commonly use bare \\r as the line separator.
    A first attempt just normalized \\r -> \\n (both are legal XML 1.0
    characters) but QuickBooks Desktop's own parser rejected that too --
    confirmed 2026-08-10 by inspecting the actual in-memory string handed
    to ProcessRequest (verified to contain only \\n, no \\r) and getting the
    identical "found an error when parsing the provided XML text stream"
    failure on retry. QB's parser apparently doesn't tolerate ANY embedded
    line break inside element text, not just \\r specifically -- so instead
    of picking a different line-break character, don't send one at all.
    Root-caused from ASI-113748 (multi-line panel spec Desc fields on 2 of
    5 line items).

    Finally, strips any XML-illegal C0 control character (see
    _ILLEGAL_XML_CHARS) that would otherwise produce malformed qbXML."""
    result = (str(value)
        # Step 0: collapse all line breaks to a single space.
        .replace('\r\n', ' ').replace('\r', ' ').replace('\n', ' ')
        # Step 1: remap Windows-1252 C1 bytes that QB may emit as raw bytes.
        .replace('', '‘')  # CP1252 0x91 -> LEFT SINGLE QUOTATION MARK
        .replace('', '’')  # CP1252 0x92 -> RIGHT SINGLE QUOTATION MARK
        .replace('', '“')  # CP1252 0x93 -> LEFT DOUBLE QUOTATION MARK
        .replace('', '”')  # CP1252 0x94 -> RIGHT DOUBLE QUOTATION MARK
        .replace('', '–')  # CP1252 0x96 -> EN DASH
        .replace('', '—')  # CP1252 0x97 -> EM DASH
        .replace('', '…')  # CP1252 0x85 -> HORIZONTAL ELLIPSIS
        # Step 2: convert typographic Unicode characters to ASCII equivalents.
        .replace('“', '"').replace('”', '"')   # curly double quotes
        .replace('‘', "'").replace('’', "'")   # curly single quotes
        .replace('–', '-').replace('—', '-')   # en/em dash
        .replace(' ', ' ')                             # non-breaking space
        .replace('…', '...')                         # ellipsis
        # Step 3: replace any remaining non-ASCII with ?
        .encode('ascii', 'replace').decode('ascii')
    )
    # Step 4: replace any illegal XML control character (see _ILLEGAL_XML_CHARS)
    # with a space -- stripping to '' would silently join adjacent words.
    result = _ILLEGAL_XML_CHARS.sub(' ', result)
    # Step 5: collapse runs of spaces (e.g. a trailing space before a
    # collapsed line break in Step 0) down to one, and trim ends.
    return re.sub(r' {2,}', ' ', result).strip()


def build_invoice_add(payload: dict) -> str:
    """Return a complete qbXML InvoiceAdd request string."""
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="stopOnError")
    req = ET.SubElement(msgs, "InvoiceAddRq", requestID="1")
    inv = ET.SubElement(req, "InvoiceAdd")

    # QB SDK requires strict element order in InvoiceAdd.
    # Use ListID for CustomerRef when available — it's encoding-agnostic and
    # avoids apostrophe mismatches (straight vs curly) that cause error 3140.
    if payload.get("customer_job_list_id"):
        _text(inv, "CustomerRef/ListID", payload["customer_job_list_id"])
    else:
        _text(inv, "CustomerRef/FullName", _ascii_safe(payload["customer_name"]))

    if payload.get("class_id"):
        _text(inv, "ClassRef/FullName", payload["class_id"])

    _text(inv, "TxnDate", payload["order_date"])

    # RefNumber is the QB-visible Invoice #/Ref No. — must be the FM Order ID,
    # not the customer's PO, so ASI's order number and QB's doc number match.
    # NOTE: payload key is "po_number" despite the name — FM sends the order ID
    # under that key; there is no "order_id" key in the actual payload.
    _text(inv, "RefNumber", _ascii_safe(str(payload["po_number"])))

    _build_bill_to(inv, payload)
    _build_ship_to(inv, payload)

    if payload.get("cust_po"):
        _text(inv, "PONumber", _ascii_safe(payload["cust_po"]))

    if payload.get("terms"):
        _text(inv, "TermsRef/FullName", _ascii_safe(payload["terms"]))

    if payload.get("rep_name"):
        _text(inv, "SalesRepRef/FullName", _ascii_safe(payload["rep_name"]))

    if payload.get("ship_via"):
        _text(inv, "ShipMethodRef/FullName", _ascii_safe(payload["ship_via"]))

    if payload.get("memo"):
        _text(inv, "Memo", _ascii_safe(payload["memo"]))

    # Line items (skip any marked exclude=True)
    for item in payload.get("line_items", []):
        if item.get("exclude"):
            continue
        li = ET.SubElement(inv, "InvoiceLineAdd")
        _text(li, "ItemRef/FullName", _ascii_safe(item["item_name"]))
        if item.get("description"):
            _text(li, "Desc", _ascii_safe(item["description"]))
        _text(li, "Quantity", str(item["quantity"]))
        _text(li, "Rate", str(item["unit_price"]))

    # Freight as a separate line item (QB Desktop standard approach)
    if payload.get("freight_amount") and float(payload["freight_amount"]) != 0:
        freight_li = ET.SubElement(inv, "InvoiceLineAdd")
        _text(freight_li, "ItemRef/FullName", "Freight")
        _text(freight_li, "Desc", "Shipping & Handling")
        _text(freight_li, "Quantity", "1")
        _text(freight_li, "Rate", str(payload["freight_amount"]))

    # STOPGAP (2026-08-07): tax_amount is intentionally NOT pushed right now.
    # The old approach (an InvoiceLineAdd referencing an ItemRef named "Sales Tax")
    # crashes with QB error 3140 -- that Item was never created in QB, and Item
    # type/list references can't be used this way for tax regardless. The real
    # fix (ItemSalesTaxRef + per-line SalesTaxCodeRef, matched against a synced
    # QB_SalesTaxItems table) is in progress but not ready. Until then this drops
    # tax from the QB posting entirely rather than crash the whole push. Do not
    # re-add a "Sales Tax" ItemRef line -- see qbxml_builder.py history/PR notes.

    # "Promise Date" (2026-08-07 correction): this is a QB custom field (Data
    # Extension), NOT the native ShipDate element -- confirmed via a real
    # SalesOrderQuery response where ShipDate just mirrored TxnDate while the
    # actual "Promise Date" DataExtRet held a distinct, meaningful date. The
    # first attempt at this wired ship_date into native ShipDate; wrong target.
    if payload.get("ship_date"):
        de = ET.SubElement(inv, "DataExt")
        _text(de, "OwnerID", "0")
        _text(de, "DataExtName", "Promise Date")
        _text(de, "DataExtType", "STR255TYPE")
        _text(de, "DataExtValue", _mmddyyyy(payload["ship_date"]))

    return _wrap_qbxml(ET.tostring(root, encoding="unicode"))


def build_sales_order_add(payload: dict) -> str:
    """Return a complete qbXML SalesOrderAdd request string.

    Same payload shape as build_invoice_add(). SalesOrderAdd mirrors
    InvoiceAdd's schema closely (CustomerRef, ClassRef, TxnDate, RefNumber,
    BillAddress, ShipAddress, PONumber, TermsRef, SalesRepRef, ShipDate,
    ShipMethodRef, Memo, line adds) — the element order below follows the
    same sequence already confirmed working for InvoiceAdd. Not yet
    verified live against QB Desktop; confirm element order/acceptance on
    first real test before relying on it.
    """
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="stopOnError")
    req = ET.SubElement(msgs, "SalesOrderAddRq", requestID="1")
    so = ET.SubElement(req, "SalesOrderAdd")

    if payload.get("customer_job_list_id"):
        _text(so, "CustomerRef/ListID", payload["customer_job_list_id"])
    else:
        _text(so, "CustomerRef/FullName", _ascii_safe(payload["customer_name"]))

    if payload.get("class_id"):
        _text(so, "ClassRef/FullName", payload["class_id"])

    # ASI-specific SO template, so QB prints/emails from the right form instead
    # of defaulting to the built-in pro forma invoice template. Per Cat's
    # request (2026-08-07). TemplateRef must appear here -- after ClassRef,
    # before TxnDate -- per the QB SDK's required SalesOrderAdd element order.
    _text(so, "TemplateRef/FullName", "ASI Sales Order")

    _text(so, "TxnDate", payload["order_date"])

    # RefNumber is the QB-visible SO # — the FM Order ID, so it matches the
    # ASI order number Cat sees on her side (same convention as InvoiceAdd).
    # NOTE: payload key is "po_number" despite the name — FM sends the order ID
    # under that key; there is no "order_id" key in the actual payload.
    _text(so, "RefNumber", _ascii_safe(str(payload["po_number"])))

    _build_bill_to(so, payload)
    _build_ship_to(so, payload)

    if payload.get("cust_po"):
        _text(so, "PONumber", _ascii_safe(payload["cust_po"]))

    if payload.get("terms"):
        _text(so, "TermsRef/FullName", _ascii_safe(payload["terms"]))

    if payload.get("rep_name"):
        _text(so, "SalesRepRef/FullName", _ascii_safe(payload["rep_name"]))

    if payload.get("ship_via"):
        _text(so, "ShipMethodRef/FullName", _ascii_safe(payload["ship_via"]))

    if payload.get("memo"):
        _text(so, "Memo", _ascii_safe(payload["memo"]))

    # Line items (skip any marked exclude=True)
    for item in payload.get("line_items", []):
        if item.get("exclude"):
            continue
        li = ET.SubElement(so, "SalesOrderLineAdd")
        _text(li, "ItemRef/FullName", _ascii_safe(item["item_name"]))
        if item.get("description"):
            _text(li, "Desc", _ascii_safe(item["description"]))
        _text(li, "Quantity", str(item["quantity"]))
        _text(li, "Rate", str(item["unit_price"]))

    # Freight as a separate line item (same convention as InvoiceAdd)
    if payload.get("freight_amount") and float(payload["freight_amount"]) != 0:
        freight_li = ET.SubElement(so, "SalesOrderLineAdd")
        _text(freight_li, "ItemRef/FullName", "Freight")
        _text(freight_li, "Desc", "Shipping & Handling")
        _text(freight_li, "Quantity", "1")
        _text(freight_li, "Rate", str(payload["freight_amount"]))

    # STOPGAP (2026-08-07): tax_amount is intentionally NOT pushed right now.
    # Same reasoning as build_invoice_add() -- see its comment for detail. This
    # is what was crashing on ASI-113705 with QB error 3140. Real fix pending.

    # "Promise Date" -- QB custom field (Data Extension), not native ShipDate.
    # See build_invoice_add()'s comment for how we know that.
    if payload.get("ship_date"):
        de = ET.SubElement(so, "DataExt")
        _text(de, "OwnerID", "0")
        _text(de, "DataExtName", "Promise Date")
        _text(de, "DataExtType", "STR255TYPE")
        _text(de, "DataExtValue", _mmddyyyy(payload["ship_date"]))

    return _wrap_qbxml(ET.tostring(root, encoding="unicode"))


def build_customer_query(customer_name: str) -> str:
    """Check if a customer (or Customer:Job) exists in QB by FullName.
    _ascii_safe is applied so non-ASCII characters in the name don't cause
    QB's XML parser to reject or silently fail the query."""
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="stopOnError")
    req = ET.SubElement(msgs, "CustomerQueryRq", requestID="1")
    _text(req, "FullName", _ascii_safe(customer_name))
    _text(req, "OwnerID", "0")
    return _wrap_qbxml(ET.tostring(root, encoding="unicode"))


def build_customer_name_filter_query(name: str) -> str:
    """Find customers/jobs whose Name field contains the given string.
    Name is the last segment only (no parent prefix), so this works even
    when the full FullName has apostrophe encoding issues.
    Used to recover a job's ListID after a 3100 'already exists' error."""
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="stopOnError")
    req = ET.SubElement(msgs, "CustomerQueryRq", requestID="1")
    f = ET.SubElement(req, "NameFilter")
    _text(f, "MatchCriterion", "Contains")
    _text(f, "Name", _ascii_safe(name))
    _text(req, "OwnerID", "0")
    return _wrap_qbxml(ET.tostring(root, encoding="unicode"))


def build_customer_query_by_list_id(list_id: str) -> str:
    """Look up a customer by QB ListID. Used as a fallback when a FullName
    query fails due to a stale name stored in FM — the ListID never changes
    even when the customer's name is edited in QB.
    ActiveStatus=All so inactive customers are found too."""
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="stopOnError")
    req = ET.SubElement(msgs, "CustomerQueryRq", requestID="1")
    _text(req, "ListID", list_id)
    # Note: ActiveStatus cannot be combined with ListID in CustomerQueryRq --
    # they are mutually exclusive filter types per the QB SDK schema.
    # Querying by ListID returns the record regardless of active status.
    _text(req, "OwnerID", "0")
    return _wrap_qbxml(ET.tostring(root, encoding="unicode"))


def build_customer_add_job(parent_list_id: str, job_name: str) -> str:
    """Create a QB sub-customer (job) under the given parent customer."""
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="stopOnError")
    req = ET.SubElement(msgs, "CustomerAddRq", requestID="1")
    cust = ET.SubElement(req, "CustomerAdd")
    _text(cust, "Name", _ascii_safe(job_name))
    _text(cust, "ParentRef/ListID", parent_list_id)
    return _wrap_qbxml(ET.tostring(root, encoding="unicode"))


def _build_bill_to(parent: ET.Element, payload: dict):
    if not payload.get("bill_to_name"):
        return
    addr = ET.SubElement(parent, "BillAddress")
    _text(addr, "Addr1", _ascii_safe(payload.get("bill_to_name", ""))[:41])
    if payload.get("bill_to_addr1"):
        _text(addr, "Addr2", _ascii_safe(payload["bill_to_addr1"])[:41])
    if payload.get("bill_to_addr2"):
        _text(addr, "Addr3", _ascii_safe(payload["bill_to_addr2"])[:41])
    _text(addr, "City",       _ascii_safe(payload.get("bill_to_city",  ""))[:31])
    _text(addr, "State",      _ascii_safe(payload.get("bill_to_state", ""))[:21])
    _text(addr, "PostalCode", _ascii_safe(payload.get("bill_to_zip",   ""))[:13])


def _build_ship_to(parent: ET.Element, payload: dict):
    if not payload.get("ship_to_name"):
        return
    addr = ET.SubElement(parent, "ShipAddress")
    _text(addr, "Addr1", _ascii_safe(payload.get("ship_to_name", ""))[:41])
    if payload.get("ship_to_addr1"):
        _text(addr, "Addr2", _ascii_safe(payload["ship_to_addr1"])[:41])
    if payload.get("ship_to_addr2"):
        _text(addr, "Addr3", _ascii_safe(payload["ship_to_addr2"])[:41])
    _text(addr, "City", _ascii_safe(payload.get("ship_to_city", ""))[:31])
    _text(addr, "State", _ascii_safe(payload.get("ship_to_state", ""))[:21])
    _text(addr, "PostalCode", _ascii_safe(payload.get("ship_to_zip", ""))[:13])


def _text(parent: ET.Element, path: str, value: str):
    """Set a possibly-nested element's text, creating intermediates as needed."""
    parts = path.split("/")
    el = parent
    for part in parts:
        existing = el.find(part)
        if existing is not None:
            el = existing
        else:
            el = ET.SubElement(el, part)
    el.text = value


def build_customer_list_query(active_status: str = "ActiveOnly") -> str:
    """Return a CustomerQueryRq for all top-level customers (no jobs).
    active_status: "ActiveOnly" (default -- used by the bulk sync) or "All"
    (used by the individual sync-by-account lookup, so a customer that exists
    in QB but is marked inactive can be detected and reported instead of
    looking identical to "never existed in QB")."""
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="stopOnError")
    req = ET.SubElement(msgs, "CustomerQueryRq", requestID="1")
    _text(req, "ActiveStatus", active_status)
    _text(req, "OwnerID", "0")
    return _wrap_qbxml(ET.tostring(root, encoding="unicode"))


# All QB item types ASI uses. Each entry: (QueryRq element name, Ret element name).
# Used by both build functions and com_handler parsing.
ITEM_QUERY_TYPES = [
    ("ItemInventory",         "ItemInventoryRet"),
    ("ItemInventoryAssembly", "ItemInventoryAssemblyRet"),
    ("ItemNonInventory",      "ItemNonInventoryRet"),
    ("ItemService",           "ItemServiceRet"),
    ("ItemOtherCharge",       "ItemOtherChargeRet"),
    ("ItemGroup",             "ItemGroupRet"),
]


def build_item_list_query() -> str:
    """Return a batch query for all active items across all QB item types."""
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="continueOnError")
    for i, (req_type, _) in enumerate(ITEM_QUERY_TYPES, start=1):
        req = ET.SubElement(msgs, f"{req_type}QueryRq", requestID=str(i))
        _text(req, "ActiveStatus", "ActiveOnly")
        _text(req, "OwnerID", "0")
    return _wrap_qbxml(ET.tostring(root, encoding="unicode"))


def build_item_sales_tax_list_query() -> str:
    """Return a batch query for all active Sales Tax Items and Sales Tax Group Items.

    Diagnostic/sync tool (2026-08-07) -- pulls QB's actual tax-item list so FM can
    mirror it (rather than hand-maintaining a separate state->rate table) and
    auto-match an order's ship-to state + FM-calculated rate against a real,
    currently-valid QB Sales Tax Item.
    """
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="continueOnError")
    req1 = ET.SubElement(msgs, "ItemSalesTaxQueryRq", requestID="1")
    _text(req1, "ActiveStatus", "ActiveOnly")
    req2 = ET.SubElement(msgs, "ItemSalesTaxGroupQueryRq", requestID="2")
    _text(req2, "ActiveStatus", "ActiveOnly")
    return _wrap_qbxml(ET.tostring(root, encoding="unicode"))


def build_item_query_by_name(item_name: str) -> str:
    """Return a batch query for a single item by FullName across all QB item types."""
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="continueOnError")
    for i, (req_type, _) in enumerate(ITEM_QUERY_TYPES, start=1):
        req = ET.SubElement(msgs, f"{req_type}QueryRq", requestID=str(i))
        _text(req, "FullName", item_name)
        _text(req, "OwnerID", "0")
    return _wrap_qbxml(ET.tostring(root, encoding="unicode"))


def build_invoice_query(ref_number: str) -> str:
    """Return an InvoiceQueryRq by RefNumber (QB invoice number)."""
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="stopOnError")
    req = ET.SubElement(msgs, "InvoiceQueryRq", requestID="1")
    _text(req, "RefNumber", ref_number)
    _text(req, "IncludeLineItems", "true")
    _text(req, "OwnerID", "0")
    return _wrap_qbxml(ET.tostring(root, encoding="unicode"))


def build_sales_order_query(ref_number: str) -> str:
    """Return a SalesOrderQueryRq by RefNumber (QB sales order number)."""
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="stopOnError")
    req = ET.SubElement(msgs, "SalesOrderQueryRq", requestID="1")
    _text(req, "RefNumber", ref_number)
    _text(req, "IncludeLineItems", "true")
    _text(req, "OwnerID", "0")
    return _wrap_qbxml(ET.tostring(root, encoding="unicode"))


def build_ship_method_query() -> str:
    """Return a ShipMethodQueryRq for all active shipping methods."""
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="stopOnError")
    req = ET.SubElement(msgs, "ShipMethodQueryRq", requestID="1")
    _text(req, "ActiveStatus", "ActiveOnly")
    return _wrap_qbxml(ET.tostring(root, encoding="unicode"))


def build_terms_query() -> str:
    """Return a TermsQueryRq for all active payment terms."""
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="stopOnError")
    req = ET.SubElement(msgs, "TermsQueryRq", requestID="1")
    _text(req, "ActiveStatus", "ActiveOnly")
    return _wrap_qbxml(ET.tostring(root, encoding="unicode"))


def build_sales_rep_query() -> str:
    """Return a SalesRepQueryRq for all active sales reps."""
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="stopOnError")
    req = ET.SubElement(msgs, "SalesRepQueryRq", requestID="1")
    _text(req, "ActiveStatus", "ActiveOnly")
    return _wrap_qbxml(ET.tostring(root, encoding="unicode"))


def build_company_query() -> str:
    """Return a qbXML CompanyQueryRq — used to verify QB connectivity."""
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="stopOnError")
    ET.SubElement(msgs, "CompanyQueryRq", requestID="1")
    return _wrap_qbxml(ET.tostring(root, encoding="unicode"))


def _wrap_qbxml(inner: str) -> str:
    return (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<?qbxml version="16.0"?>'
        + inner
    )
