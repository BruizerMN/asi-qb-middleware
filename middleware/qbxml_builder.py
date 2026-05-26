"""
Build qbXML InvoiceAdd requests from FM invoice data.

FM sends a JSON payload with the invoice. This module turns that into
the qbXML string that QuickBooks Web Connector feeds to QB Desktop.

Expected payload structure (all fields are strings unless noted):
{
    "customer_name": "Acme Corp",
    "order_id": "12345",
    "order_date": "2026-05-05",       # YYYY-MM-DD
    "ship_date": "2026-05-10",        # YYYY-MM-DD, optional
    "cust_po": "PO-9876",             # optional — QB RefNumber (Ref No. / Invoice #)
    "po_number": "ORD-12345",         # optional — QB PONumber (P.O. No. field); use FM OrderID
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
    "amount_paid": "557.50",
    "payment_method": "Visa",         # optional
    "salesperson": "Jane Smith",      # optional
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

import xml.etree.ElementTree as ET


def _ascii_safe(value: str) -> str:
    """Replace Unicode typographic characters with ASCII equivalents.
    QB Desktop's XML parser does not support non-ASCII characters."""
    return (str(value)
        .replace('“', '"').replace('”', '"')   # curly double quotes
        .replace('‘', "'").replace('’', "'")   # curly single quotes / apostrophe
        .replace('–', '-').replace('—', '-')   # en dash, em dash
        .replace(' ', ' ')                          # non-breaking space
        .replace('…', '...')                        # ellipsis
        .encode('ascii', 'replace').decode('ascii')      # replace any remaining non-ASCII with ?
    )


def build_invoice_add(payload: dict) -> str:
    """Return a complete qbXML InvoiceAdd request string."""
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="stopOnError")
    req = ET.SubElement(msgs, "InvoiceAddRq", requestID="1")
    inv = ET.SubElement(req, "InvoiceAdd")

    # QB SDK requires strict element order in InvoiceAdd
    _text(inv, "CustomerRef/FullName", _ascii_safe(payload["customer_name"]))

    if payload.get("class_id"):
        _text(inv, "ClassRef/ListID", payload["class_id"])

    _text(inv, "TxnDate", payload["order_date"])

    if payload.get("cust_po"):
        _text(inv, "RefNumber", _ascii_safe(payload["cust_po"]))

    _build_bill_to(inv, payload)
    _build_ship_to(inv, payload)

    if payload.get("po_number"):
        _text(inv, "PONumber", _ascii_safe(payload["po_number"]))

    if payload.get("terms"):
        _text(inv, "TermsRef/FullName", _ascii_safe(payload["terms"]))

    if payload.get("ship_date"):
        _text(inv, "ShipDate", payload["ship_date"])

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

    # Tax — passed as a pre-calculated amount via a tax item
    # QB Desktop requires a SalesTaxCodeRef or a tax line item.
    # We use a non-taxable item "Sales Tax" to record the Avalara-calculated amount.
    if payload.get("tax_amount") and float(payload["tax_amount"]) != 0:
        tax_li = ET.SubElement(inv, "InvoiceLineAdd")
        _text(tax_li, "ItemRef/FullName", "Sales Tax")
        _text(tax_li, "Desc", "Sales Tax (Avalara)")
        _text(tax_li, "Quantity", "1")
        _text(tax_li, "Rate", str(payload["tax_amount"]))

    return _wrap_qbxml(ET.tostring(root, encoding="unicode"))


def build_customer_query(customer_name: str) -> str:
    """Check if a customer (or Customer:Job) exists in QB by FullName."""
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="stopOnError")
    req = ET.SubElement(msgs, "CustomerQueryRq", requestID="1")
    _text(req, "FullName", customer_name)
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


def build_item_query(item_name: str) -> str:
    """Check if an item exists in QB by full name."""
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="stopOnError")
    req = ET.SubElement(msgs, "ItemNonInventoryQueryRq", requestID="1")
    _text(req, "FullName", item_name)
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


def build_customer_list_query() -> str:
    """Return a CustomerQueryRq for all active top-level customers (no jobs)."""
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="stopOnError")
    req = ET.SubElement(msgs, "CustomerQueryRq", requestID="1")
    _text(req, "ActiveStatus", "ActiveOnly")
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


def build_item_query_by_name(item_name: str) -> str:
    """Return a batch query for a single item by FullName across all QB item types."""
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="continueOnError")
    for i, (req_type, _) in enumerate(ITEM_QUERY_TYPES, start=1):
        req = ET.SubElement(msgs, f"{req_type}QueryRq", requestID=str(i))
        _text(req, "FullName", item_name)
        _text(req, "OwnerID", "0")
    return _wrap_qbxml(ET.tostring(root, encoding="unicode"))


def build_terms_query() -> str:
    """Return a TermsQueryRq for all active payment terms."""
    root = ET.Element("QBXML")
    msgs = ET.SubElement(root, "QBXMLMsgsRq", onError="stopOnError")
    req = ET.SubElement(msgs, "TermsQueryRq", requestID="1")
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
