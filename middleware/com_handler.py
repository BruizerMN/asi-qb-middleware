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
    _ascii_safe,
    build_company_query,
    build_customer_query, build_customer_query_by_list_id,
    build_customer_name_filter_query, build_customer_add_job,
    build_customer_list_query, build_item_list_query,
    build_item_query_by_name, build_terms_query, build_ship_method_query,
    build_sales_rep_query, build_invoice_query, ITEM_QUERY_TYPES,
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


def detect_open_slug() -> str:
    """
    Return the slug for the currently open QB company ("acoustical" or "architectural").
    Raises RuntimeError if QB is not running, no file is open, or the open file
    doesn't match any configured company.
    """
    rp, ticket = _open_session()
    try:
        info = _get_company_info(rp, ticket)
        slug = _detect_slug(info["name"])
        if not slug:
            raise RuntimeError(
                f"The open QuickBooks company '{info['name']}' is not configured in the middleware. "
                "Expected the Acoustical Surfaces or Architectural Surfaces company file."
            )
        return slug
    finally:
        _close_session(rp, ticket)


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


def _find_job_list_id(rp, ticket, parent_list_id: str, job_name: str) -> str:
    """Find an existing Customer:Job's ListID when FullName matching fails.

    Uses a NameFilter query (matches the short Name field, not the full path)
    and then filters results in Python by ParentRef/ListID. This sidesteps the
    apostrophe-encoding mismatch that makes FullName queries unreliable.
    """
    safe_job = _ascii_safe(job_name).lower()
    resp = rp.ProcessRequest(ticket, build_customer_name_filter_query(job_name))
    root = ET.fromstring(resp)
    for cust in root.findall(".//CustomerRet"):
        if (cust.findtext("ParentRef/ListID") or "") == parent_list_id:
            cust_name = _ascii_safe(cust.findtext("Name") or "").lower()
            if cust_name == safe_job:
                return cust.findtext("ListID") or ""
    return ""


def ensure_customer_job(
    customer_job_fullname: str,
    expected_slug: str,
    customer_list_id: str = "",
) -> tuple[str, str]:
    """
    Ensure a Customer:Job (sub-customer) exists in QB. Creates it if missing.

    customer_job_fullname is "CustomerFullName:JobName" as sent in the invoice payload.
    If there is no colon, the name is a top-level customer and no action is needed.

    customer_list_id is the QB ListID (QB_CustomerID from FM). When provided, it is
    used as a fallback if the FullName query returns nothing — this handles stale names
    in FM caused by name edits in QB since the last sync, or apostrophes and other
    characters that behave differently across the sync and invoice paths.

    Returns the customer_job_fullname to use in InvoiceAdd. Normally this is the
    same as the input, but when the ListID fallback fires and finds the parent under a
    different name, the corrected "ActualName:JobName" string is returned so the
    InvoiceAdd uses the current QB name rather than the stale FM one.

    Raises RuntimeError if the parent customer cannot be found or job creation fails.
    """
    if ":" not in customer_job_fullname:
        return customer_job_fullname

    rp, ticket = _open_session()
    try:
        info = _get_company_info(rp, ticket)
        verify_company(info, expected_slug)

        # _ascii_safe is required: QB's XML parser rejects non-ASCII characters.
        # (build_customer_query also applies it, but we need the safe version here
        # to construct the corrected return value consistently.)
        safe_name = _ascii_safe(customer_job_fullname)
        parent_safe, job_name = safe_name.split(":", 1)

        # Check whether the job already exists under the (safe) name FM has.
        resp = rp.ProcessRequest(ticket, build_customer_query(safe_name))
        root = ET.fromstring(resp)
        job_ret = root.find(".//CustomerRet")
        if job_ret is not None:
            # Job found — return its ListID so InvoiceAdd can use CustomerRef/ListID
            # and avoid apostrophe-encoding mismatches with CustomerRef/FullName.
            return customer_job_fullname, (job_ret.findtext("ListID") or "")

        # Job not found. Look up parent customer by FullName.
        resp = rp.ProcessRequest(ticket, build_customer_query(parent_safe))
        root = ET.fromstring(resp)
        parent_ret = root.find(".//CustomerRet")

        # FullName lookup failed — try ListID fallback if we have one.
        # This fires when the customer's name in QB differs from what FM has stored
        # (e.g. name corrected in QB since last sync, apostrophe added/removed).
        corrected_parent = None
        if parent_ret is None and customer_list_id:
            resp = rp.ProcessRequest(ticket, build_customer_query_by_list_id(customer_list_id))
            root = ET.fromstring(resp)
            parent_ret = root.find(".//CustomerRet")
            if parent_ret is not None:
                corrected_parent = _ascii_safe(parent_ret.findtext("FullName") or "")
                # Check whether the job already exists under the corrected parent name.
                corrected_job = corrected_parent + ":" + job_name
                resp2 = rp.ProcessRequest(ticket, build_customer_query(corrected_job))
                root2 = ET.fromstring(resp2)
                job_ret2 = root2.find(".//CustomerRet")
                if job_ret2 is not None:
                    return corrected_job, (job_ret2.findtext("ListID") or "")

        if parent_ret is None:
            list_id_info = (
                f"ListID tried: {customer_list_id!r}" if customer_list_id
                else "no customer_list_id in payload"
            )
            raise RuntimeError(
                f"Customer '{parent_safe}' was not found in QuickBooks "
                f"({list_id_info}). "
                "Run QB - Sync Customers to link this customer, then try again."
            )

        parent_list_id = parent_ret.findtext("ListID") or ""
        if not parent_list_id:
            raise RuntimeError(
                f"Could not retrieve the QuickBooks ListID for customer '{parent_safe}'."
            )

        # Create the job under the parent.
        resp = rp.ProcessRequest(ticket, build_customer_add_job(parent_list_id, job_name))
        root = ET.fromstring(resp)
        rs = root.find(".//CustomerAddRs")
        job_name_used = (corrected_parent + ":" + job_name) if corrected_parent else customer_job_fullname

        if rs is not None:
            code = rs.get("statusCode", "0")
            if code == "0":
                # Job created — extract its ListID from the response.
                new_job_ret = root.find(".//CustomerRet")
                new_job_list_id = new_job_ret.findtext("ListID") if new_job_ret is not None else ""
                return job_name_used, (new_job_list_id or "")
            elif code == "3100":
                # Job already exists in QB but our FullName query couldn't find it
                # (apostrophe encoding mismatch between FM and QB). Recover by
                # searching for the job by Name field + parent ListID match.
                existing_id = _find_job_list_id(rp, ticket, parent_list_id, job_name)
                return job_name_used, existing_id
            else:
                msg = rs.get("statusMessage", "Unknown error")
                raise RuntimeError(
                    f"QuickBooks rejected job creation for '{job_name}': {msg} (code {code})"
                )

        return job_name_used, ""

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

        # Submit the invoice
        response = rp.ProcessRequest(ticket, qbxml)

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


def submit_sales_order(qbxml: str, expected_slug: str) -> str:
    """
    Verify the correct QB company file is open, then submit a SalesOrderAdd.
    Returns the QB sales order number (RefNumber) on success.
    Raises RuntimeError with a user-friendly message on any failure.

    Mirrors submit_invoice() exactly, just checking SalesOrderAddRs instead
    of InvoiceAddRs. Not yet tested against a live QB Desktop session.
    """
    rp, ticket = _open_session()
    try:
        info = _get_company_info(rp, ticket)
        verify_company(info, expected_slug)

        response = rp.ProcessRequest(ticket, qbxml)

        root = ET.fromstring(response)
        rs = root.find(".//SalesOrderAddRs")

        if rs is None:
            raise RuntimeError("No SalesOrderAddRs in QB response.")

        status_code = rs.get("statusCode", "")
        status_msg  = rs.get("statusMessage", "")

        if status_code != "0":
            raise RuntimeError(
                f"QuickBooks rejected the sales order: {status_msg} (code {status_code})"
            )

        txn_number = rs.findtext(".//TxnNumber")
        ref_number = rs.findtext(".//RefNumber")
        txn_id     = rs.findtext(".//TxnID")

        # Return whichever identifier QB provided — prefer RefNumber (user-visible)
        so_number = ref_number or txn_number or txn_id
        if not so_number:
            raise RuntimeError(
                f"Sales order was created in QuickBooks but no SO number was returned. "
                f"Response: {ET.tostring(rs, encoding='unicode')}"
            )

        return so_number

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
                "full_name":      _ascii_safe(cust.findtext("FullName") or ""),
                "account_number": account_number,
            })

        # Warm the cache so subsequent individual lookups skip the QB fetch.
        _customer_cache_populate(expected_slug, customers)
        return customers

    finally:
        _close_session(rp, ticket)


def get_all_items(expected_slug: str) -> list:
    """
    Return all active QB items as a list of dicts across all item types.
    Each dict: {list_id, name}.
    name matches FM's productID field (confirmed by Cat).
    Covers: Inventory, InventoryAssembly, NonInventory, Service, OtherCharge, Group.
    """
    rp, ticket = _open_session()
    try:
        info = _get_company_info(rp, ticket)
        verify_company(info, expected_slug)

        response = rp.ProcessRequest(ticket, build_item_list_query())
        root = ET.fromstring(response)

        items = []
        for _, ret_type in ITEM_QUERY_TYPES:
            for item in root.findall(f".//{ret_type}"):
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


def get_customer_by_account(account_number: str, bypass_cache: bool = False) -> dict | None:
    """
    Return a single customer matching account_number, or None if not found.

    QB CustomerQueryRq has no AccountNumber filter, so a full fetch is needed
    when the cache is cold. On a cache hit (within _CACHE_TTL), the QB fetch is
    skipped entirely — only a cheap CompanyQueryRq is made to detect the slug.

    bypass_cache=True forces a fresh QB fetch regardless of cache state. Always
    pass True from sync endpoints — the purpose of a sync is to get current data
    from QB, so a cache hit would defeat the point and return stale values.
    """
    rp, ticket = _open_session()
    try:
        # Detect which company file is open so we can key the cache correctly.
        info = _get_company_info(rp, ticket)
        slug = _detect_slug(info["name"])

        # Cache hit — skip QB fetch unless caller explicitly wants fresh data.
        if slug and not bypass_cache:
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
                "full_name":      _ascii_safe(cust.findtext("FullName") or ""),
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


def lookup_customer_by_list_id(list_id: str) -> dict:
    """
    Diagnostic helper: look up a single customer by ListID and return whatever
    QB says. Used by the /fm/ping-customer endpoint to debug stale-name issues.
    Returns a dict with keys: found (bool), list_id, full_name, active_status,
    account_number, raw_status_code, raw_status_message.
    """
    rp, ticket = _open_session()
    try:
        resp = rp.ProcessRequest(ticket, build_customer_query_by_list_id(list_id))
        root = ET.fromstring(resp)
        rs = root.find(".//CustomerQueryRs")
        status_code = rs.get("statusCode", "?") if rs is not None else "?"
        status_msg  = rs.get("statusMessage", "") if rs is not None else ""
        cust = root.find(".//CustomerRet")
        if cust is None:
            return {
                "found": False,
                "list_id": list_id,
                "raw_status_code": status_code,
                "raw_status_message": status_msg,
            }
        return {
            "found": True,
            "list_id": cust.findtext("ListID") or "",
            "full_name": cust.findtext("FullName") or "",
            "full_name_ascii": _ascii_safe(cust.findtext("FullName") or ""),
            "account_number": cust.findtext("AccountNumber") or "",
            "is_active": cust.findtext("IsActive") or "",
            "raw_status_code": status_code,
            "raw_status_message": status_msg,
        }
    finally:
        _close_session(rp, ticket)


def get_all_ship_methods() -> list:
    """Return all active QB shipping methods as a list of {name} dicts."""
    rp, ticket = _open_session()
    try:
        response = rp.ProcessRequest(ticket, build_ship_method_query())
        root = ET.fromstring(response)
        return [
            {"name": m.findtext("Name") or ""}
            for m in root.findall(".//ShipMethodRet")
        ]
    finally:
        _close_session(rp, ticket)


def get_all_terms() -> list:
    """Return all active QB payment terms as a list of {name, days_due, discount_days, discount_pct} dicts."""
    rp, ticket = _open_session()
    try:
        response = rp.ProcessRequest(ticket, build_terms_query())
        root = ET.fromstring(response)
        terms = []
        for t in root.findall(".//StandardTermsRet") + root.findall(".//DateDrivenTermsRet"):
            terms.append({
                "name":           t.findtext("Name") or "",
                "days_due":       t.findtext("StdDueDays") or t.findtext("DayOfMonthDue") or "",
                "discount_days":  t.findtext("StdDiscountDays") or "",
                "discount_pct":   t.findtext("DiscountPct") or "",
            })
        return terms
    finally:
        _close_session(rp, ticket)


def get_all_sales_reps() -> list:
    """Return all active QB sales reps as a list of {initials} dicts."""
    rp, ticket = _open_session()
    try:
        response = rp.ProcessRequest(ticket, build_sales_rep_query())
        root = ET.fromstring(response)
        return [
            {"initials": r.findtext("Initial") or ""}
            for r in root.findall(".//SalesRepRet")
        ]
    finally:
        _close_session(rp, ticket)


def get_invoice(ref_number: str, expected_slug: str) -> dict:
    """
    Fetch a QB invoice by RefNumber (the QB invoice number stored in FM QB_InvoiceID).
    Returns a structured dict suitable for HTML rendering.
    Raises RuntimeError if QB is not running, wrong company, or invoice not found.
    """
    rp, ticket = _open_session()
    try:
        info = _get_company_info(rp, ticket)
        verify_company(info, expected_slug)

        resp = rp.ProcessRequest(ticket, build_invoice_query(ref_number))
        root = ET.fromstring(resp)

        rs = root.find(".//InvoiceQueryRs")
        if rs is not None:
            code = rs.get("statusCode", "0")
            if code not in ("0", "1"):
                raise RuntimeError(
                    f"QB invoice query failed: {rs.get('statusMessage')} (code {code})"
                )

        ret = root.find(".//InvoiceRet")
        if ret is None:
            raise RuntimeError(f"Invoice #{ref_number} was not found in QuickBooks.")

        def _addr(tag: str) -> dict:
            a = ret.find(tag)
            if a is None:
                return {}
            return {
                "addr1": a.findtext("Addr1") or "",
                "addr2": a.findtext("Addr2") or "",
                "addr3": a.findtext("Addr3") or "",
                "city":  a.findtext("City")  or "",
                "state": a.findtext("State") or "",
                "zip":   a.findtext("PostalCode") or "",
            }

        lines = []
        for li in ret.findall("InvoiceLineRet"):
            lines.append({
                "item":        li.findtext("ItemRef/FullName") or "",
                "description": li.findtext("Desc") or "",
                "uom":         li.findtext("UnitOfMeasure") or "",
                "quantity":    li.findtext("Quantity") or "",
                "rate":        li.findtext("Rate") or "",
                "amount":      li.findtext("Amount") or "",
            })

        return {
            "txn_id":       ret.findtext("TxnID") or "",
            "txn_date":     ret.findtext("TxnDate") or "",
            "ref_number":   ret.findtext("RefNumber") or ref_number,
            "customer":     ret.findtext("CustomerRef/FullName") or "",
            "class_name":   ret.findtext("ClassRef/FullName") or "",
            "bill_address": _addr("BillAddress"),
            "ship_address": _addr("ShipAddress"),
            "po_number":    ret.findtext("PONumber") or "",
            "terms":        ret.findtext("TermsRef/FullName") or "",
            "rep":          ret.findtext("SalesRepRef/FullName") or "",
            "ship_date":    ret.findtext("ShipDate") or "",
            "ship_method":  ret.findtext("ShipMethodRef/FullName") or "",
            "memo":         ret.findtext("Memo") or "",
            "subtotal":     ret.findtext("Subtotal") or "0",
            "total":        ret.findtext("TotalAmount") or "0",
            "applied":      ret.findtext("AppliedAmount") or "0",
            "balance":      ret.findtext("BalanceRemaining") or "0",
            "line_items":   lines,
        }
    finally:
        _close_session(rp, ticket)


def get_item_by_name(item_name: str) -> dict | None:
    """
    Return a single item matching item_name, or None if not found.
    Searches across all QB item types (Inventory, NonInventory, Service, etc.).
    """
    rp, ticket = _open_session()
    try:
        response = rp.ProcessRequest(ticket, build_item_query_by_name(item_name))
        root = ET.fromstring(response)
        for _, ret_type in ITEM_QUERY_TYPES:
            item = root.find(f".//{ret_type}")
            if item is not None:
                return {
                    "list_id": item.findtext("ListID") or "",
                    "name":    item.findtext("Name") or "",
                }
        return None
    finally:
        _close_session(rp, ticket)
