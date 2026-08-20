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

    Line breaks (bare \\r, \\n, or \\r\\n) are deliberately left untouched --
    they're legal XML 1.0 and QuickBooks accepts them fine in normal element
    text (e.g. multi-line item Desc fields with one spec attribute per
    line, which is standard practice across most of ASI's catalog and has
    always worked). An earlier investigation on 2026-08-10 wrongly suspected
    embedded line breaks as the cause of a parse crash on ASI-113748 and
    this function briefly collapsed them all to spaces; that was reverted
    once the real cause was found to be an unrelated bug (a top-level
    DataExt element that isn't valid on InvoiceAdd/SalesOrderAdd at all --
    see build_invoice_add()'s comment). Do not reintroduce line-break
    stripping here without new evidence.

    Finally, strips any XML-illegal C0 control character (see
    _ILLEGAL_XML_CHARS) that would otherwise produce malformed qbXML."""
    result = (str(value)
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
    return _ILLEGAL_XML_CHARS.sub(' ', result)


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

    # "Promise Date" is NOT included here. It's a QB custom field (Data
    # Extension) that must be set via a separate DataExtAdd request AFTER
    # this transaction is created, referencing its TxnID -- see
    # build_data_ext_add(). An earlier attempt embedded it inline as a
    # trailing <DataExt> child of InvoiceAdd/SalesOrderAdd; confirmed
    # 2026-08-10 via isolated live testing that DataExt is not a valid
    # top-level child of these Add requests at all (only within line items,
    # per the qbXML schema) -- QuickBooks rejected the ENTIRE request with a
    # raw "found an error when parsing the provided XML text stream" COM
    # exception, not a clean statusCode rejection. See com_handler.py's
    # submit_invoice()/submit_sales_order() for the follow-up call.

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

    # "Promise Date" is NOT included here -- see build_invoice_add()'s comment
    # for the full explanation. Must be set via a separate build_data_ext_add()
    # request after this transaction is created.

    return _wrap_qbxml(ET.tostring(root, encoding="unicode"))


def build_data_ext_add(txn_id: str, txn_data_ext_type: str, data_ext_name: str, data_ext_value: str) -> str:
    """Return a DataExtAddRq to set a transaction-level custom field (Data
    Extension) on an EXISTING transaction, referenced by TxnID.

    Must be called as a follow-up request after InvoiceAdd/SalesOrderAdd
    succeeds -- DataExt is not a valid top-level child of those Add requests
    (confirmed 2026-08-10 via isolated live testing; see build_invoice_add()'s
    comment for the full story). txn_data_ext_type is the QB transaction type
    name as a plain string, e.g. "Invoice" or "SalesOrder".

    Verified working live 2026-08-10 against a real existing Sales Order
    (queried by RefNumber to get its TxnID, then this call set Promise Date
    on it cleanly, statusCode 0)."""
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="stopOnError")
    req = ET.SubElement(msgs, "DataExtAddRq", requestID="1")
    de = ET.SubElement(req, "DataExtAdd")
    _text(de, "OwnerID", "0")
    _text(de, "DataExtName", _ascii_safe(data_ext_name))
    _text(de, "TxnDataExtType", txn_data_ext_type)
    _text(de, "TxnID", txn_id)
    _text(de, "DataExtValue", _ascii_safe(data_ext_value))
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
    Used to recover a job's ListID after a 3100 'already exists' error.

    NOTE (2026-08-11): an attempt to add ActiveStatus=All here (to also find
    inactive conflicting records after a duplicate-name CustomerAdd) made
    QuickBooks reject the whole request outright -- "found an error when
    parsing the provided XML text stream" (a raw COM exception, not a clean
    statusCode rejection), reproduced live by Bill. NameFilter + ActiveStatus
    is apparently not a valid combination in this qbXML version, unlike the
    documented ListID+ActiveStatus incompatibility. Do not re-add ActiveStatus
    to this function without testing directly against QB Desktop first --
    see com_handler.create_or_update_customer()'s duplicate-name handling for
    how that need is now met a different way (searching the customer list
    already fetched earlier in the same call, which already proved this
    exact combination is unsupported)."""
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


# ---------------------------------------------------------------------------
# QB Customer field map (2026-08-11)
#
# One reusable, order-aware table driving both CustomerAdd and CustomerMod --
# they share the identical field shape. Built to make adding a new QB
# Customer field a DATA change, not a code change: qbXML is strict about
# element order (confirmed the hard way, 2026-08-11 -- an unsupported
# NameFilter+ActiveStatus combination elsewhere in this file made QB reject
# a request outright rather than return a clean rejection), so a flat,
# order-preserving table is what lets a future field just be another row
# here instead of requiring careful re-verification of QB's schema order
# each time. Cat's team is expected to keep sending more fields over time
# (Bill, 2026-08-11) -- this table is sized to make that painless.
#
# Each entry is either:
#   (logical_key, xml_path, max_len_or_None) -- a simple field. logical_key
#     is what FM sends inside payload["fields"]; xml_path is the element name
#     directly under CustomerAdd/CustomerMod; max_len truncates (QB's real
#     limits, matching the same values already used for invoice/SO
#     addresses in _build_bill_to/_build_ship_to above) or None to skip
#     truncation.
#   "BILL_ADDRESS" -- a marker; the whole BillAddress group is built by
#     _build_customer_bill_address() using _CUSTOMER_BILL_ADDRESS_FIELDS'
#     own internal order below.
#   "CONTACT_NAME_SPLIT" -- a marker; FirstName/LastName are built by
#     _build_customer_contact_name() by splitting a single "contact_name"
#     field at its first space (Bill's design, 2026-08-12).
#   "ALWAYS_ACTIVE" -- a marker; always emits <IsActive>true</IsActive>,
#     unconditionally, regardless of the fields dict. Cat's explicit
#     instruction, 2026-08-12 (relayed by Bill): reactivate/undelete any
#     customer we touch "without a fuss or notice" -- she confirmed this is
#     safe since her team deletes/inactivates customers in QB exceedingly
#     rarely. Confirmed empirically the same day that IsActive=true resolves
#     BOTH plain-inactive and QB's separate "deleted" (red-X) state via the
#     same mechanism -- see com_handler.diagnostic_reactivate_customer().
#
# A field is emitted only when fields.get(logical_key) is truthy (present
# AND non-empty) -- so a currently-unpopulated field (e.g. bill_addr3/4,
# not yet exposed in FM's UI as of 2026-08-11) is silently skipped rather
# than emitted empty, and costs nothing to leave wired in ahead of need.
#
# name/company_name is duplicated on purpose -- QB's CustomerAdd requires
# both Name (the globally-unique internal list key) and CompanyName
# (display), and this solution always sets them to the same value.
#
# FirstName/LastName sit here (right after CompanyName, before BillAddress)
# per QB's actual Customer schema order -- confirmed by Bill's Name/
# CompanyName/BillAddress/Email/AccountNumber build succeeding live,
# 2026-08-12, though this specific position hasn't been live-tested yet.
#
# EXTENDING: add a new (key, path, max_len) tuple below in the correct QB
# schema position -- see QuickBooks' own CustomerAdd/CustomerMod SDK
# reference for where a field not listed here belongs. No other code
# changes needed once FM starts sending that key.
CUSTOMER_FIELD_MAP = [
    ("company_name", "Name", 41),
    "ALWAYS_ACTIVE",
    ("company_name", "CompanyName", 41),
    "CONTACT_NAME_SPLIT",
    "BILL_ADDRESS",
    ("email", "Email", None),
    ("account_number", "AccountNumber", 41),
]

# BillAddress's own internal element order (QB's standard Address aggregate,
# reused across Customer/Invoice/SalesOrder/etc.) -- Addr1-Addr5, then City/
# State/PostalCode. Company name goes in Addr1 (Cat's explicit instruction,
# 2026-08-12, from her annotated screenshot -- the Bill To block's first row
# should be the company name, not the street), pushing FM's four address
# lines to Addr2-Addr5 -- uses all five available lines, none wasted.
_CUSTOMER_BILL_ADDRESS_FIELDS = [
    ("company_name", "Addr1", 41),
    ("bill_addr1", "Addr2", 41),
    ("bill_addr2", "Addr3", 41),
    ("bill_addr3", "Addr4", 41),
    ("bill_addr4", "Addr5", 41),
    ("bill_city", "City", 31),
    ("bill_state", "State", 21),
    ("bill_zip", "PostalCode", 13),
]

# Which of _CUSTOMER_BILL_ADDRESS_FIELDS' keys actually indicate "there's a
# real address to report" -- company_name is deliberately excluded here even
# though it's one of the emitted lines, since company_name is present on
# nearly every customer and its presence alone shouldn't trigger an
# otherwise-empty BillAddress block.
_CUSTOMER_BILL_ADDRESS_PRESENCE_KEYS = (
    "bill_addr1", "bill_addr2", "bill_addr3", "bill_addr4",
    "bill_city", "bill_state", "bill_zip",
)


def _build_customer_bill_address(cust: ET.Element, fields: dict):
    """Emit <BillAddress>...</BillAddress> if at least one real address field
    is present -- an empty BillAddress element (or one with only the company
    name and no actual address) is pointless and QB doesn't need it."""
    if not any(fields.get(key) for key in _CUSTOMER_BILL_ADDRESS_PRESENCE_KEYS):
        return
    addr = ET.SubElement(cust, "BillAddress")
    for key, path, max_len in _CUSTOMER_BILL_ADDRESS_FIELDS:
        value = fields.get(key)
        if value:
            safe = _ascii_safe(str(value))
            _text(addr, path, safe[:max_len] if max_len else safe)


def _build_customer_contact_name(cust: ET.Element, fields: dict):
    """Emit <FirstName>/<LastName> split from a single "contact_name" field
    at its first space (Bill's design, 2026-08-12): "Steve Kerkvliet" ->
    FirstName "Steve", LastName "Kerkvliet"; a single-word name goes
    entirely into FirstName, LastName left unset. Source is FM's
    AcctPayableName field (same field this solution already uses for
    Email, confirmed via layout-mode screenshot, 2026-08-12) -- sent to FM
    under the logical key "contact_name" since the QB concept is a general
    contact, not specifically an AP one. 25-char max on both, QB's actual
    limit for these two fields (unlike the 41-char limit used elsewhere in
    this table for Name/CompanyName/Address lines)."""
    raw = fields.get("contact_name")
    if not raw:
        return
    raw = str(raw).strip()
    first, _, last = raw.partition(" ")
    _text(cust, "FirstName", _ascii_safe(first)[:25])
    last = last.strip()
    if last:
        _text(cust, "LastName", _ascii_safe(last)[:25])


def _build_customer_fields(cust: ET.Element, fields: dict):
    """Emit CUSTOMER_FIELD_MAP's fields onto a CustomerAdd/CustomerMod element,
    in the table's fixed order, from a flat `fields` dict (payload["fields"]
    merged with account_number by the caller -- see com_handler.create_or_update_customer)."""
    for entry in CUSTOMER_FIELD_MAP:
        if entry == "BILL_ADDRESS":
            _build_customer_bill_address(cust, fields)
            continue
        if entry == "CONTACT_NAME_SPLIT":
            _build_customer_contact_name(cust, fields)
            continue
        if entry == "ALWAYS_ACTIVE":
            _text(cust, "IsActive", "true")
            continue
        key, path, max_len = entry
        value = fields.get(key)
        if value:
            safe = _ascii_safe(str(value))
            _text(cust, path, safe[:max_len] if max_len else safe)


def build_customer_add(fields: dict) -> str:
    """Return a qbXML CustomerAdd request for a new top-level QB customer.

    `fields` is a flat dict keyed by CUSTOMER_FIELD_MAP's logical field names
    (see that table for the full current field set and how to extend it).
    Name is QB's globally-unique internal list key (<=41 chars) -- collisions
    raise QB error 3100, handled by the caller (com_handler.create_or_update_customer).
    """
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="stopOnError")
    req = ET.SubElement(msgs, "CustomerAddRq", requestID="1")
    cust = ET.SubElement(req, "CustomerAdd")
    _build_customer_fields(cust, fields)
    return _wrap_qbxml(ET.tostring(root, encoding="unicode"))


def build_customer_mod(fields: dict) -> str:
    """Return a qbXML CustomerMod request updating an existing top-level QB
    customer's fields (see CUSTOMER_FIELD_MAP for the full current set).

    Requires fields["list_id"] and fields["edit_sequence"] -- QB's
    optimistic-lock token, fetched immediately before this call (see
    com_handler.create_or_update_customer). A stale EditSequence causes QB
    to reject the request rather than silently overwrite a concurrent edit.
    ListID/EditSequence are handled directly here, not via CUSTOMER_FIELD_MAP,
    since they're QB's own identity/locking tokens, not synced customer data.
    """
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="stopOnError")
    req = ET.SubElement(msgs, "CustomerModRq", requestID="1")
    cust = ET.SubElement(req, "CustomerMod")
    _text(cust, "ListID", fields["list_id"])
    _text(cust, "EditSequence", fields["edit_sequence"])
    _build_customer_fields(cust, fields)
    return _wrap_qbxml(ET.tostring(root, encoding="unicode"))


def build_customer_reactivate(list_id: str, edit_sequence: str) -> str:
    """DIAGNOSTIC ONLY (2026-08-12) -- return a qbXML CustomerMod that sets
    IsActive=true and touches nothing else. Built to answer one specific
    open question empirically: does IsActive=true also resolve QB's
    separate "deleted" state (distinct from plain inactive -- confirmed by
    Bill, 2026-08-12: QB Desktop shows a red-X marker and a distinct "would
    you like to undelete it?" prompt for deleted customers, not just the
    plain inactive checkbox), or does it only reactivate plain-inactive
    ones? Not wired into any production create/update path -- see
    com_handler.diagnostic_reactivate_customer() / the
    /fm/debug-reactivate-customer route for how this gets used."""
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="stopOnError")
    req = ET.SubElement(msgs, "CustomerModRq", requestID="1")
    cust = ET.SubElement(req, "CustomerMod")
    _text(cust, "ListID", list_id)
    _text(cust, "EditSequence", edit_sequence)
    _text(cust, "IsActive", "true")
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


def build_invoice_query(ref_number: str, include_linked_txns: bool = False) -> str:
    """Return an InvoiceQueryRq by RefNumber (QB invoice number).

    include_linked_txns=True (2026-08-20, for delete_transaction's
    pre-delete safety check): asks QB to include each LinkedTxn (e.g. a
    Payment applied against this Invoice) in the response, so a caller can
    detect and refuse to delete a transaction that has dependents QB itself
    would otherwise block on with a less friendly error. Defaults False --
    existing callers (view/lookup paths) don't need this extra data."""
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="stopOnError")
    req = ET.SubElement(msgs, "InvoiceQueryRq", requestID="1")
    _text(req, "RefNumber", ref_number)
    _text(req, "IncludeLineItems", "true")
    if include_linked_txns:
        _text(req, "IncludeLinkedTxns", "true")
    _text(req, "OwnerID", "0")
    return _wrap_qbxml(ET.tostring(root, encoding="unicode"))


def build_sales_order_query(ref_number: str, include_linked_txns: bool = False) -> str:
    """Return a SalesOrderQueryRq by RefNumber (QB sales order number).

    include_linked_txns=True (2026-08-20, for delete_transaction's pre-delete
    safety check): asks QB to include each LinkedTxn (e.g. an Invoice
    generated from this Sales Order via Cat's team's manual SO->Invoice
    conversion) in the response -- QB will refuse to delete a Sales Order
    that still has a linked Invoice, so this lets the caller detect that and
    give the user a clear, specific message instead of a raw QB rejection.
    Defaults False -- existing callers (view/lookup/Promise-Date paths)
    don't need this extra data.

    NOT YET LIVE-TESTED against a real linked SO/Invoice pair as of
    2026-08-20 -- verify the LinkedTxn element actually appears in the
    response, and in the expected position in the request, before relying
    on it in production."""
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="stopOnError")
    req = ET.SubElement(msgs, "SalesOrderQueryRq", requestID="1")
    _text(req, "RefNumber", ref_number)
    _text(req, "IncludeLineItems", "true")
    if include_linked_txns:
        _text(req, "IncludeLinkedTxns", "true")
    _text(req, "OwnerID", "0")
    return _wrap_qbxml(ET.tostring(root, encoding="unicode"))


def build_txn_del(txn_type: str, txn_id: str) -> str:
    """Return a TxnDelRq to permanently delete an existing transaction.

    txn_type must be a valid qbXML TxnDelType value -- "SalesOrder" or
    "Invoice" for this project's purposes (the full QB SDK list is much
    longer, but those are the only two ASI's integration creates).

    Irreversible in QuickBooks once it succeeds. Callers are responsible for
    any pre-delete safety checks (e.g. refusing to delete a transaction with
    linked dependents) -- this function only builds the request."""
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="stopOnError")
    req = ET.SubElement(msgs, "TxnDelRq", requestID="1")
    _text(req, "TxnDelType", txn_type)
    _text(req, "TxnID", txn_id)
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
