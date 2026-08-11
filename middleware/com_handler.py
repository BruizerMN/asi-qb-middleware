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
    _ascii_safe, _mmddyyyy,
    build_company_query,
    build_customer_query, build_customer_query_by_list_id,
    build_customer_name_filter_query, build_customer_add_job,
    build_customer_add, build_customer_mod,
    build_customer_list_query, build_item_list_query,
    build_item_query_by_name, build_terms_query, build_ship_method_query,
    build_sales_rep_query, build_invoice_query, build_sales_order_query, ITEM_QUERY_TYPES,
    build_item_sales_tax_list_query, build_data_ext_add,
)

APP_NAME = "ASI QB Middleware"

# QB status code for "This list has been modified by another user" -- a
# transient multi-user list-lock conflict, not a real rejection. See
# ensure_customer_job() for where this is retried.
_LIST_CONFLICT_CODE = "3180"
_LIST_CONFLICT_MAX_ATTEMPTS = 3
_LIST_CONFLICT_RETRY_DELAY = 1.5  # seconds

# Recognized QB rejection message substrings (case-insensitive) -> a
# friendlier, actionable message to show instead of QuickBooks' raw nested
# error text. Matched on the MESSAGE, not the status code -- QB reuses the
# same generic codes (e.g. 3180) for unrelated rejection reasons, confirmed
# 2026-08-10 when a real order hit "credit limit exceeded" under the exact
# same code used elsewhere for list-lock conflicts. The original QB message
# is always appended, never hidden -- this only adds context on top.
_KNOWN_REJECTIONS = [
    (
        "credit limit",
        "This order could not be posted to QuickBooks because the customer's "
        "credit limit has been exceeded. Have your QuickBooks administrator "
        "review the customer's credit limit or apply an override, then "
        "resubmit this order.",
    ),
]


def _friendly_qb_message(status_msg: str) -> str:
    """Prefix a known-cause explanation onto a raw QB rejection message when
    recognized; otherwise return it unchanged. Never drops the original QB
    text -- see _KNOWN_REJECTIONS."""
    lowered = status_msg.lower()
    for needle, explanation in _KNOWN_REJECTIONS:
        if needle in lowered:
            return f"{explanation}\n\n(QuickBooks said: {status_msg})"
    return status_msg

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
        job_name_used = (corrected_parent + ":" + job_name) if corrected_parent else customer_job_fullname

        for attempt in range(1, _LIST_CONFLICT_MAX_ATTEMPTS + 1):
            resp = rp.ProcessRequest(ticket, build_customer_add_job(parent_list_id, job_name))
            root = ET.fromstring(resp)
            rs = root.find(".//CustomerAddRs")

            if rs is None:
                return job_name_used, ""

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
            elif code == _LIST_CONFLICT_CODE and attempt < _LIST_CONFLICT_MAX_ATTEMPTS:
                # "This list has been modified by another user" -- a transient
                # multi-user contention error, not a real rejection. Each
                # workstation runs its own independent QB session, so this can
                # happen when several people submit orders around the same
                # time. Retrying almost always succeeds within a second or
                # two. Root-caused 2026-08-10 (Cat Shoop's team, first day of
                # multi-workstation rollout with zero retry logic in place).
                time.sleep(_LIST_CONFLICT_RETRY_DELAY)
                continue
            else:
                msg = rs.get("statusMessage", "Unknown error")
                raise RuntimeError(
                    f"QuickBooks rejected job creation for '{job_name}': {msg} (code {code})"
                )

        return job_name_used, ""

    finally:
        _close_session(rp, ticket)


def _set_promise_date(rp, ticket, txn_id: str, txn_type: str, ship_date: str) -> str:
    """Best-effort follow-up: set the 'Promise Date' custom field on an
    already-created transaction via DataExtAdd. Must run AFTER the parent
    Add succeeds and reuses its open session -- DataExt is not valid inline
    within InvoiceAdd/SalesOrderAdd itself (see build_invoice_add()'s
    comment / qbxml_builder.py history, 2026-08-10).

    Never raises -- the parent transaction has already succeeded by the time
    this runs, so a Promise Date failure shouldn't fail the whole submission.
    But it also isn't silently swallowed: returns a warning string on
    failure ("" on success) for the caller to surface, since a hidden
    failure behind a reported success is exactly the bug class that bit
    this project before (see QB_Preflight Build 0028 history)."""
    try:
        de_xml = build_data_ext_add(txn_id, txn_type, "Promise Date", _mmddyyyy(ship_date))
        de_resp = rp.ProcessRequest(ticket, de_xml)
        de_root = ET.fromstring(de_resp)
        de_rs = de_root.find(".//DataExtAddRs")
        if de_rs is None or de_rs.get("statusCode", "") != "0":
            msg = de_rs.get("statusMessage", "no response") if de_rs is not None else "no response"
            return f"Promise Date was not set: {msg}"
        return ""
    except Exception as e:
        return f"Promise Date was not set: {e}"


def submit_invoice(qbxml: str, expected_slug: str, ship_date: str = "") -> tuple:
    """
    Verify the correct QB company file is open, then submit an InvoiceAdd.
    Returns (invoice_number, warning) -- warning is "" unless the Promise
    Date follow-up failed (the invoice itself still succeeded in that case).
    Raises RuntimeError with a user-friendly message on any failure of the
    invoice itself.
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
                f"QuickBooks rejected the invoice: {_friendly_qb_message(status_msg)} (code {status_code})"
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

        warning = ""
        if ship_date and txn_id:
            warning = _set_promise_date(rp, ticket, txn_id, "Invoice", ship_date)

        return invoice_number, warning

    finally:
        _close_session(rp, ticket)


def submit_sales_order(qbxml: str, expected_slug: str, ship_date: str = "") -> tuple:
    """
    Verify the correct QB company file is open, then submit a SalesOrderAdd.
    Returns (so_number, warning) -- warning is "" unless the Promise Date
    follow-up failed (the sales order itself still succeeded in that case).
    Raises RuntimeError with a user-friendly message on any failure of the
    sales order itself.

    Mirrors submit_invoice() exactly, just checking SalesOrderAddRs instead
    of InvoiceAddRs.
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
                f"QuickBooks rejected the sales order: {_friendly_qb_message(status_msg)} (code {status_code})"
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

        warning = ""
        if ship_date and txn_id:
            warning = _set_promise_date(rp, ticket, txn_id, "SalesOrder", ship_date)

        return so_number, warning

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

        # Cache miss — fetch all customers (including inactive -- see below),
        # populate cache, return match.
        #
        # ActiveStatus="All" here (unlike get_all_customers' bulk "ActiveOnly"
        # query): a customer that already exists in QB but is marked inactive
        # must still be found and reported as such, not treated identically to
        # a customer that has never been added to QB at all. Without this, a
        # sync retry loops forever with no way for the user to discover why.
        response = rp.ProcessRequest(ticket, build_customer_list_query("All"))
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
                "is_active":      (cust.findtext("IsActive") or "true").lower() != "false",
            }
            customers.append(c)
            if acct.strip().lower() == target:
                result = c

        if slug:
            _customer_cache_populate(slug, customers)

        return result

    finally:
        _close_session(rp, ticket)


def create_or_update_customer(account_number: str, company_name: str) -> dict:
    """
    "UpdateCreate" entry point for a single top-level QB customer -- unlike
    get_customer_by_account (read-only matching), this function writes to QB.
    Matched by AccountNumber, same convention as the rest of the customer sync
    family. Creates the customer if no match is found; updates it (CustomerMod)
    if one is.

    v1 minimal field set (Cat's first-round mapping, 2026-08-11): Name,
    CompanyName (both set from company_name), AccountNumber. Cat's team is
    still finalizing the rest of the field mapping -- more fields will be
    added to build_customer_add()/build_customer_mod() once that's ready.

    No expected_slug/company param, deliberately -- mirrors get_customer_by_account:
    operates against whichever QB company file is currently open on this
    workstation. When called automatically during a sales order push, the
    caller (QB_Preflight) has already verified the correct company file is
    open via its own earlier check, so a second verification here would be
    redundant.

    Returns one of:
      {"action": "created", "customer": {list_id, full_name, account_number}}
      {"action": "updated", "customer": {list_id, full_name, account_number}}
      {"action": "duplicate_name", "conflict": {list_id, full_name, account_number}}
        -- QB rejected the CustomerAdd because another customer/job already
           owns that Name (error 3100). "conflict" identifies the existing
           record so the FM user has enough information to find and resolve
           it in QB, rather than a bare "name already in use" message.
    Raises RuntimeError on any other QB rejection or connection failure.
    """
    rp, ticket = _open_session()
    try:
        info = _get_company_info(rp, ticket)
        slug = _detect_slug(info["name"])

        safe_name = _ascii_safe(company_name)[:41]
        safe_account = _ascii_safe(account_number)[:41]

        # Full fetch (ActiveStatus=All) -- QB has no AccountNumber filter, and
        # an existing-but-inactive match must still be treated as "exists"
        # (CustomerMod), not "not found" -- same reasoning as get_customer_by_account.
        response = rp.ProcessRequest(ticket, build_customer_list_query("All"))
        root = ET.fromstring(response)
        target = safe_account.strip().lower()
        existing = None
        for cust in root.findall(".//CustomerRet"):
            if cust.find("ParentRef") is not None:
                continue
            acct = cust.findtext("AccountNumber") or ""
            if acct.strip().lower() == target:
                existing = cust
                break

        if existing is not None:
            list_id = existing.findtext("ListID") or ""
            edit_sequence = existing.findtext("EditSequence") or ""
            mod_xml = build_customer_mod({
                "list_id": list_id,
                "edit_sequence": edit_sequence,
                "name": safe_name,
                "account_number": safe_account,
            })
            resp = rp.ProcessRequest(ticket, mod_xml)
            root2 = ET.fromstring(resp)
            rs = root2.find(".//CustomerModRs")
            if rs is None or rs.get("statusCode", "") != "0":
                msg  = rs.get("statusMessage", "no response") if rs is not None else "no response"
                code = rs.get("statusCode", "?") if rs is not None else "?"
                raise RuntimeError(f"QuickBooks rejected the customer update: {msg} (code {code})")
            ret = root2.find(".//CustomerRet")
            if slug:
                _customer_cache.pop(slug, None)
            return {
                "action": "updated",
                "customer": {
                    "list_id":        (ret.findtext("ListID") if ret is not None else "") or list_id,
                    "full_name":      _ascii_safe((ret.findtext("FullName") if ret is not None else "") or safe_name),
                    "account_number": (ret.findtext("AccountNumber") if ret is not None else "") or safe_account,
                },
            }

        # Not found -- create.
        add_xml = build_customer_add({"name": safe_name, "account_number": safe_account})
        resp = rp.ProcessRequest(ticket, add_xml)
        root2 = ET.fromstring(resp)
        rs = root2.find(".//CustomerAddRs")
        if rs is None:
            raise RuntimeError("No CustomerAddRs in QB response.")

        code = rs.get("statusCode", "")
        if code == "0":
            ret = root2.find(".//CustomerRet")
            if slug:
                _customer_cache.pop(slug, None)
            return {
                "action": "created",
                "customer": {
                    "list_id":        (ret.findtext("ListID") if ret is not None else "") or "",
                    "full_name":      _ascii_safe((ret.findtext("FullName") if ret is not None else "") or safe_name),
                    "account_number": (ret.findtext("AccountNumber") if ret is not None else "") or safe_account,
                },
            }
        elif code == "3100":
            # Name collision -- look up the conflicting record by Name so FM
            # can show the user its account number (Bill, 2026-08-11: "enough
            # information to find and address the issue"), instead of a bare
            # "name already in use" message.
            resp2 = rp.ProcessRequest(ticket, build_customer_name_filter_query(safe_name))
            root3 = ET.fromstring(resp2)
            conflict = None
            for cust in root3.findall(".//CustomerRet"):
                if _ascii_safe(cust.findtext("Name") or "").lower() == safe_name.lower():
                    conflict = cust
                    break
            if conflict is None:
                # Exact Name match not found (e.g. apostrophe/encoding drift) --
                # fall back to the first NameFilter hit rather than nothing.
                conflict = root3.find(".//CustomerRet")
            return {
                "action": "duplicate_name",
                "conflict": {
                    "list_id":        (conflict.findtext("ListID") if conflict is not None else "") or "",
                    "full_name":      _ascii_safe((conflict.findtext("FullName") if conflict is not None else "") or ""),
                    "account_number": (conflict.findtext("AccountNumber") if conflict is not None else "") or "",
                },
            }
        else:
            msg = rs.get("statusMessage", "Unknown error")
            raise RuntimeError(
                f"QuickBooks rejected customer creation for '{safe_name}': {msg} (code {code})"
            )

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


def get_sales_order(ref_number: str, expected_slug: str) -> dict:
    """
    Fetch a QB sales order by RefNumber (the QB SO number stored in FM QB_InvoiceID).
    Returns a structured dict suitable for HTML rendering.
    Raises RuntimeError if QB is not running, wrong company, or SO not found.

    Mirrors get_invoice() — a SalesOrderRet has no AppliedAmount/BalanceRemaining
    (those are invoice/payment concepts that don't exist until the SO is converted),
    so this returns "total" only, no applied/balance.
    """
    rp, ticket = _open_session()
    try:
        info = _get_company_info(rp, ticket)
        verify_company(info, expected_slug)

        resp = rp.ProcessRequest(ticket, build_sales_order_query(ref_number))
        root = ET.fromstring(resp)

        rs = root.find(".//SalesOrderQueryRs")
        if rs is not None:
            code = rs.get("statusCode", "0")
            if code not in ("0", "1"):
                raise RuntimeError(
                    f"QB sales order query failed: {rs.get('statusMessage')} (code {code})"
                )

        ret = root.find(".//SalesOrderRet")
        if ret is None:
            raise RuntimeError(f"Sales Order #{ref_number} was not found in QuickBooks.")

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
        for li in ret.findall("SalesOrderLineRet"):
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
            "line_items":   lines,
        }
    finally:
        _close_session(rp, ticket)


def get_raw_query_response(kind: str, ref_number: str, expected_slug: str) -> str:
    """
    Return the raw, unparsed qbXML response for an InvoiceQuery, SalesOrderQuery,
    or the full Sales Tax Item/Group list.

    One-off diagnostic tool -- not used by any live posting or rendering path.
    kind is "invoice", "sales-order", or "sales-tax-items" (ref_number ignored
    for "sales-tax-items" -- it's a list query, not a lookup by number).
    """
    rp, ticket = _open_session()
    try:
        info = _get_company_info(rp, ticket)
        verify_company(info, expected_slug)
        if kind == "invoice":
            qbxml = build_invoice_query(ref_number)
        elif kind == "sales-order":
            qbxml = build_sales_order_query(ref_number)
        else:
            qbxml = build_item_sales_tax_list_query()
        return rp.ProcessRequest(ticket, qbxml)
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
