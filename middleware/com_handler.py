"""
QuickBooks Desktop COM interface.

Sends qbXML requests directly to the open QB company file via win32com.
This module only works on Windows with QuickBooks Desktop installed.
"""

import time
import xml.etree.ElementTree as ET

try:
    import win32com.client
    import pythoncom
    _COM_AVAILABLE = True
except ImportError:
    _COM_AVAILABLE = False

from .config import QB_COMPANY_FILES
from .qbxml_builder import (
    build_company_query,
    build_customer_list_query, build_item_list_query,
    build_item_query_by_name,
)

APP_NAME = "ASI QB Middleware"

# ---------------------------------------------------------------------------
# Customer list cache
# Keyed by company slug. Avoids a full QB CustomerQueryRq on every individual
# sync — QB has no AccountNumber filter, so without this every lookup fetches
# all customers (~20 s). The All sync always repopulates; individual syncs
# reuse the cache for up to _CACHE_TTL seconds.
# ---------------------------------------------------------------------------
_CACHE_TTL = 1800  # 30 minutes

# { slug: {"customers": [...], "by_account": {lower_acct_no: dict}, "ts": float} }
_customer_cache: dict = {}


def _customer_cache_populate(slug: str, customers: list) -> None:
    """Store a customer list in the cache, indexed by lower-cased account_number."""
    _customer_cache[slug] = {
        "customers":  customers,
        "by_account": {c["account_number"].strip().lower(): c for c in customers},
        "ts":         time.time(),
    }


def _customer_cache_lookup(slug: str, account_number: str) -> tuple:
    """
    Return (hit, customer_or_None).
    hit=True  — cache was valid; customer_or_None is the result (may be None = not found).
    hit=False — cache stale/missing; caller should fetch from QB.
    """
    entry = _customer_cache.get(slug)
    if entry and (time.time() - entry["ts"]) < _CACHE_TTL:
        return True, entry["by_account"].get(account_number.strip().lower())
    return False, None


def _detect_slug(company_name: str) -> str | None:
    """Map an open QB company display name back to its configured slug."""
    name_lower = company_name.strip().lower()
    for slug, path in QB_COMPANY_FILES.items():
        if _expected_company_name(path) == name_lower:
            return slug
    return None


def _expected_company_name(configured_path: str) -> str:
    """
    Derive the expected QB company name from the configured file path.
    QB returns the company name, not the file path, in CompanyQueryRq.
    We extract it from the filename: 'Q:\\Test Parent 2.23.QBW' -> 'test parent 2.23'
    Comparison is case-insensitive.
    """
    import os
    basename = os.path.basename(configured_path.strip())
    # Strip .qbw extension (case-insensitive)
    if basename.lower().endswith(".qbw"):
        basename = basename[:-4]
    return basename.lower()


def _open_session():
    """
    Open a QB COM connection and session. Returns (rp, ticket).
    Caller is responsible for calling _close_session().
    """
    if not _COM_AVAILABLE:
        raise RuntimeError(
            "pywin32 is not installed. Run: pip install pywin32"
        )

    try:
        rp = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
    except Exception:
        raise RuntimeError(
            "QuickBooks is not running. Please open QuickBooks and a company file, then try again."
        )

    try:
        rp.OpenConnection2("", APP_NAME, 1)
    except Exception as e:
        raise RuntimeError(
            f"QuickBooks refused the connection. Make sure QuickBooks is open and try again. ({e})"
        )

    try:
        # "" = currently open file, 2 = omDontCare (single or multi-user)
        ticket = rp.BeginSession("", 2)
    except Exception:
        try:
            rp.CloseConnection()
        except Exception:
            pass
        raise RuntimeError(
            "No QuickBooks company file is open. "
            "Please open the correct company file and try again."
        )

    return rp, ticket


def _close_session(rp, ticket):
    """Close QB session and connection."""
    try:
        rp.EndSession(ticket)
    except Exception:
        pass
    try:
        rp.CloseConnection()
    except Exception:
        pass


def _get_company_info(rp, ticket) -> dict:
    """Returns {'name': '...', 'file': '...'} for the currently open company."""
    response = rp.ProcessRequest(ticket, build_company_query())
    root = ET.fromstring(response)

    ret = root.find(".//CompanyRet")
    if ret is None:
        raise RuntimeError("CompanyQueryRq returned no data from QuickBooks.")

    return {
        "name": ret.findtext("CompanyName") or "",
        "file": ret.findtext("CompanyFileName") or "",
    }


def verify_company(info: dict, expected_slug: str):
    """
    Raise RuntimeError if the open company doesn't match expected_slug.
    Pass the dict returned by get_open_company() or _get_company_info().
    """
    expected_file = QB_COMPANY_FILES.get(expected_slug, "")

    if not expected_file:
        raise RuntimeError(
            f"No QB company file configured for '{expected_slug}'. "
            "Check QB_FILE_ACOUSTICAL / QB_FILE_ARCHITECTURAL in .env."
        )

    expected_name = _expected_company_name(expected_file)
    actual_name   = info["name"].strip().lower()

    if actual_name != expected_name:
        raise RuntimeError(
            f"Wrong QuickBooks company file is open. "
            f"This invoice requires '{expected_file}', "
            f"but '{info['name']}' is currently open. "
            f"Please switch QuickBooks to the correct company file and try again."
        )


def get_open_company() -> dict:
    """
    Returns {'name': '...', 'file': '...'} for the currently open QB company.
    Raises RuntimeError if QB is not running or no file is open.
    """
    rp, ticket = _open_session()
    try:
        return _get_company_info(rp, ticket)
    finally:
        _close_session(rp, ticket)


def submit_invoice(qbxml: str, expected_slug: str) -> str:
    """
    Verify the correct QB company file is open, then submit an InvoiceAdd.
    Returns the QB invoice number (TxnNumber) on success.
    Raises RuntimeError with a user-friendly message on any failure.
    """
    rp, ticket = _open_session()
    try:
        # Verify the correct company file is open before submitting.
        info = _get_company_info(rp, ticket)
        verify_company(info, expected_slug)

        # Log the qbxml being sent to user home dir (always writable by task scheduler)
        import os
        _log = os.path.join(os.path.expanduser("~"), "qb_last_request.xml")
        with open(_log, "w", encoding="utf-8") as f:
            f.write(qbxml)

        # Submit the invoice
        response = rp.ProcessRequest(ticket, qbxml)

        # Log raw QB response
        log_path2 = os.path.normpath(os.path.join(os.path.dirname(__file__), "..", "qb_last_response.xml"))
        with open(log_path2, "w", encoding="utf-8") as f:
            f.write(response)

        # Parse the invoice number from the response
        root = ET.fromstring(response)
        rs = root.find(".//InvoiceAddRs")

        if rs is None:
            raise RuntimeError("No InvoiceAddRs in QB response.")

        status_code = rs.get("statusCode", "")
        status_msg  = rs.get("statusMessage", "")

        if status_code != "0":
            raise RuntimeError(
                f"QuickBooks rejected the invoice: {status_msg} (code {status_code})"
            )

        txn_number = rs.findtext(".//TxnNumber")
        ref_number = rs.findtext(".//RefNumber")
        txn_id     = rs.findtext(".//TxnID")

        # Return whichever identifier QB provided — prefer RefNumber (user-visible)
        invoice_number = ref_number or txn_number or txn_id
        if not invoice_number:
            raise RuntimeError(
                f"Invoice was created in QuickBooks but no invoice number was returned. "
                f"Response: {ET.tostring(rs, encoding='unicode')}"
            )

        return invoice_number

    finally:
        _close_session(rp, ticket)


def get_all_customers(expected_slug: str) -> list:
    """
    Return all active top-level QB customers as a list of dicts.
    Each dict: {list_id, full_name, account_number}.
    Sub-customers (jobs) are excluded — they have a ParentRef element.
    Customers without an AccountNumber are excluded (can't match to FM).
    """
    rp, ticket = _open_session()
    try:
        info = _get_company_info(rp, ticket)
        verify_company(info, expected_slug)

        response = rp.ProcessRequest(ticket, build_customer_list_query())
        root = ET.fromstring(response)

        rs = root.find(".//CustomerQueryRs")
        if rs is not None:
            status_code = rs.get("statusCode", "0")
            if status_code not in ("0", "1"):
                raise RuntimeError(
                    f"QB CustomerQuery failed: {rs.get('statusMessage')} (code {status_code})"
                )

        customers = []
        for cust in root.findall(".//CustomerRet"):
            if cust.find("ParentRef") is not None:
                continue
            account_number = cust.findtext("AccountNumber") or ""
            if not account_number:
                continue
            customers.append({
                "list_id":        cust.findtext("ListID") or "",
                "full_name":      cust.findtext("FullName") or "",
                "account_number": account_number,
            })

        # Warm the cache so subsequent individual lookups skip the QB fetch.
        _customer_cache_populate(expected_slug, customers)
        return customers

    finally:
        _close_session(rp, ticket)


def get_all_items(expected_slug: str) -> list:
    """
    Return all active non-inventory QB items as a list of dicts.
    Each dict: {list_id, name}.
    name matches FM's productID field (confirmed by Cat).
    """
    rp, ticket = _open_session()
    try:
        info = _get_company_info(rp, ticket)
        verify_company(info, expected_slug)

        response = rp.ProcessRequest(ticket, build_item_list_query())
        root = ET.fromstring(response)

        rs = root.find(".//ItemNonInventoryQueryRs")
        if rs is not None:
            status_code = rs.get("statusCode", "0")
            if status_code not in ("0", "1"):
                raise RuntimeError(
                    f"QB ItemQuery failed: {rs.get('statusMessage')} (code {status_code})"
                )

        items = []
        for item in root.findall(".//ItemNonInventoryRet"):
            name = item.findtext("Name") or ""
            if not name:
                continue
            items.append({
                "list_id": item.findtext("ListID") or "",
                "name":    name,
            })
        return items

    finally:
        _close_session(rp, ticket)


def get_customer_by_account(account_number: str) -> dict | None:
    """
    Return a single customer matching account_number, or None if not found.

    QB CustomerQueryRq has no AccountNumber filter, so a full fetch is needed
    when the cache is cold. On a cache hit (within _CACHE_TTL), the QB fetch is
    skipped entirely — only a cheap CompanyQueryRq is made to detect the slug.
    """
    rp, ticket = _open_session()
    try:
        # Detect which company file is open so we can key the cache correctly.
        info = _get_company_info(rp, ticket)
        slug = _detect_slug(info["name"])

        # Cache hit — no customer list query needed.
        if slug:
            hit, customer = _customer_cache_lookup(slug, account_number)
            if hit:
                return customer

        # Cache miss — fetch all customers, populate cache, return match.
        response = rp.ProcessRequest(ticket, build_customer_list_query())
        root = ET.fromstring(response)
        target = account_number.strip().lower()

        customers = []
        result = None
        for cust in root.findall(".//CustomerRet"):
            if cust.find("ParentRef") is not None:
                continue
            acct = cust.findtext("AccountNumber") or ""
            if not acct:
                continue
            c = {
                "list_id":        cust.findtext("ListID") or "",
                "full_name":      cust.findtext("FullName") or "",
                "account_number": acct,
            }
            customers.append(c)
            if acct.strip().lower() == target:
                result = c

        if slug:
            _customer_cache_populate(slug, customers)

        return result

    finally:
        _close_session(rp, ticket)


def get_item_by_name(item_name: str) -> dict | None:
    """Return a single non-inventory item matching name, or None if not found."""
    rp, ticket = _open_session()
    try:
        response = rp.ProcessRequest(ticket, build_item_query_by_name(item_name))
        root = ET.fromstring(response)
        item = root.find(".//ItemNonInventoryRet")
        if item is None:
            return None
        return {
            "list_id": item.findtext("ListID") or "",
            "name":    item.findtext("Name") or "",
        }
    finally:
        _close_session(rp, ticket)
