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
    build_customer_add, build_customer_mod, build_customer_reactivate,
    build_customer_list_query, build_item_list_query,
    build_item_query_by_name, build_terms_query, build_ship_method_query,
    build_sales_rep_query, build_invoice_query, build_sales_order_query, ITEM_QUERY_TYPES,
    build_item_sales_tax_list_query, build_data_ext_add,
)

APP_NAME = "ASI QB Middleware"

# QB status codes for "This list has been modified by another user" -- a
# transient multi-user list-lock conflict, not a real rejection. See
# ensure_customer_job() and create_or_update_customer() for where this is
# retried. 3170 confirmed as another code for the same underlying condition
# 2026-08-12 -- reproduced live via CustomerMod (diagnostic_reactivate_customer),
# not just the CustomerAdd path 3180 was originally found on.
_LIST_CONFLICT_CODES = ("3180", "3170")
_LIST_CONFLICT_MAX_ATTEMPTS = 3
_LIST_CONFLICT_RETRY_DELAY = 1.5  # seconds

# 3200 -- "The provided edit sequence ... is out-of-date." Distinct from the
# transient list-lock codes above: the record was genuinely modified in QB
# since we fetched our EditSequence, so resending the same request fails
# again. Requires re-querying for a fresh EditSequence before retrying, not
# just a blind resend. See create_or_update_customer()'s CustomerMod retry
# loop, root-caused 2026-08-13 (Nicole Kalkes, ASI-113804 area) after the
# morning's "customer not found in FileMaker" incident produced several
# rapid repeat CustomerMod calls against the same customers.
_STALE_EDIT_SEQUENCE_CODE = "3200"

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


def _customer_cache_populate(slug: str, customers: list, complete: bool) -> None:
    """
    Store a customer list in the cache, indexed by lower-cased account_number.

    complete=True means this list came from an ActiveStatus="All" fetch (every
    customer, active or inactive) -- safe for any caller that needs to
    distinguish "genuinely never in QB" from "exists but inactive" (e.g.
    create_or_update_customer's reactivate-on-touch logic). complete=False
    means ActiveOnly (the bulk/get_all_customers query) -- a "not found" from
    this list only means "not found among active customers", NOT "never
    existed in QB", so callers that need the stronger guarantee should not
    trust a hit against an incomplete entry (see require_complete below).
    """
    _customer_cache[slug] = {
        "customers":  customers,
        "by_account": {c["account_number"].strip().lower(): c for c in customers},
        "ts":         time.time(),
        "complete":   complete,
    }


def _customer_cache_lookup(slug: str, account_number: str, require_complete: bool = False) -> tuple:
    """
    Return (hit, customer_or_None).
    hit=True  — cache was valid (and, if require_complete, was an ActiveStatus="All"
                fetch); customer_or_None is the result (may be None = not found).
    hit=False — cache stale/missing/not-complete-enough; caller should fetch from QB.

    require_complete=True: only accept a hit from a cache entry populated via
    an ActiveStatus="All" fetch (see _customer_cache_populate) -- pass this
    when "no match" must mean "does not exist in QB at all" and not merely
    "not currently active", e.g. before deciding whether to CustomerAdd.
    """
    entry = _customer_cache.get(slug)
    if entry and (time.time() - entry["ts"]) < _CACHE_TTL:
        if require_complete and not entry.get("complete", False):
            return False, None
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
            elif code in _LIST_CONFLICT_CODES and attempt < _LIST_CONFLICT_MAX_ATTEMPTS:
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
        # complete=False: this is an ActiveOnly fetch, so a miss here only
        # means "not active", not "never existed in QB" -- see
        # _customer_cache_populate's docstring.
        _customer_cache_populate(expected_slug, customers, complete=False)
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
            # complete=True: this fetch used ActiveStatus="All", so it covers
            # every customer/job -- a miss here safely means "never existed in
            # QB", not just "not active". See _customer_cache_populate's docstring.
            _customer_cache_populate(slug, customers, complete=True)

        return result

    finally:
        _close_session(rp, ticket)


def warm_customer_cache() -> dict:
    """
    Dev tool (Bill, 2026-08-20): force-refresh the customer-list cache for
    whichever QB company is currently open, with no other side effects --
    doesn't read, create, or modify any specific customer record. Exists so
    a tech can manually warm the cache from FM (e.g. to set up a clean speed
    test of create_or_update_customer's cache-hit path, without spending a
    real first-time customer sync just to warm it).

    Always does a live ActiveStatus="All" fetch -- never reads the existing
    cache -- and populates it complete=True, the same shape
    create_or_update_customer's cache-hit check trusts.

    Returns {"slug": "...", "company_name": "...", "count": N}.
    Raises RuntimeError if QB isn't open or the open file isn't a
    recognized ASI company (same as the rest of this module).
    """
    rp, ticket = _open_session()
    try:
        info = _get_company_info(rp, ticket)
        slug = _detect_slug(info["name"])
        if not slug:
            raise RuntimeError(
                f"QB company '{info['name']}' is open but not a recognized ASI company file."
            )

        response = rp.ProcessRequest(ticket, build_customer_list_query("All"))
        root = ET.fromstring(response)

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

        _customer_cache_populate(slug, customers, complete=True)

        return {"slug": slug, "company_name": info["name"], "count": len(customers)}

    finally:
        _close_session(rp, ticket)


def create_or_update_customer(account_number: str, fields: dict, existing_list_id: str = "") -> dict:
    """
    "UpdateCreate" entry point for a single top-level QB customer -- unlike
    get_customer_by_account (read-only matching), this function writes to QB.
    Matched by AccountNumber, same convention as the rest of the customer sync
    family. Creates the customer if no match is found; updates it (CustomerMod)
    if one is.

    `fields` is a flat dict keyed by qbxml_builder.CUSTOMER_FIELD_MAP's logical
    field names (company_name, bill_addr1-4, bill_city, bill_state, bill_zip,
    email, as of 2026-08-11 -- see that table for the authoritative current
    list and how to add more without a code change). company_name is
    effectively required -- QB's Name can't be blank. Values are raw/
    unsanitized here; build_customer_add()/build_customer_mod() ascii-safe
    and truncate each one internally.

    No expected_slug/company param, deliberately -- mirrors get_customer_by_account:
    operates against whichever QB company file is currently open on this
    workstation. When called automatically during a sales order push, the
    caller (QB_Preflight) has already verified the correct company file is
    open via its own earlier check, so a second verification here would be
    redundant.

    existing_list_id (2026-08-11, Bill's design): FM's currently-stored
    QB_CustomerID, if any. When provided, it's validated FIRST via a direct
    ListID query (cheap, exact, immune to the AccountNumber-scan finding a
    stale/wrong match): if the returned record's AccountNumber still matches
    `account_number`, the link is current and the expensive full-customer-list
    fetch below is skipped entirely -- proceeds straight to CustomerMod using
    the ListID query's own ListID/EditSequence (2026-08-11 perf fix: that full
    fetch costs ~20s+ and was previously always run regardless, even though a
    validated ListID already tells us everything the scan would have found).
    If it does NOT match -- the FM-side link has gone stale (e.g. someone
    manually corrected the account number in QB without re-syncing FM) --
    the mismatch is recorded in the returned "stale_link" dict and the
    function proceeds exactly as if no link had ever existed. Deliberately
    does NOT try to "repair" the stale record (e.g. push FM's account number
    onto it) -- that would risk silently overwriting a genuinely different
    QB customer if the old link was simply wrong to begin with. Worst case
    of this conservative approach: an orphaned old QB record plus a new one,
    which a human can notice and merge -- safer than guessing.

    Every step is appended to the returned "trace" list (Bill, 2026-08-11:
    "verbosely log everything... if it were to go sideways, we'll need as
    much info as we can get our hands on") -- callers should always log the
    full trace, not just the summary fields.

    Returns one of:
      {"action": "created", "customer": {list_id, full_name, account_number},
       "stale_link": {...} or None, "trace": [...]}
      {"action": "updated", "customer": {list_id, full_name, account_number},
       "stale_link": {...} or None, "trace": [...]}
      {"action": "duplicate_name", "conflict": {list_id, full_name, account_number},
       "stale_link": {...} or None, "trace": [...]}
        -- QB rejected the CustomerAdd because another customer/job already
           owns that Name (error 3100). "conflict" identifies the existing
           record so the FM user has enough information to find and resolve
           it in QB, rather than a bare "name already in use" message.
    Raises RuntimeError on any other QB rejection or connection failure --
    the trace-so-far is appended to the exception message so it isn't lost.
    """
    trace = []
    stale_link = None

    rp, ticket = _open_session()
    try:
        info = _get_company_info(rp, ticket)
        slug = _detect_slug(info["name"])
        trace.append(f"session opened; QB company='{info['name']}' (slug={slug!r})")

        safe_name = _ascii_safe(fields.get("company_name", ""))[:41]
        safe_account = _ascii_safe(account_number)[:41]
        trace.append(
            f"inputs: account_number={safe_account!r} company_name={safe_name!r} "
            f"existing_list_id={existing_list_id!r} fields={sorted(fields.keys())!r}"
        )

        # --- Stale-link pre-check (Bill's design, 2026-08-11) ---------------
        if existing_list_id:
            trace.append(f"validating existing_list_id={existing_list_id!r} via direct ListID query")
            val_resp = rp.ProcessRequest(ticket, build_customer_query_by_list_id(existing_list_id))
            val_root = ET.fromstring(val_resp)
            val_cust = val_root.find(".//CustomerRet")
            if val_cust is None:
                trace.append("ListID not found in QB (deleted/never existed?) -- treating as unlinked")
                stale_link = {
                    "cleared": True,
                    "reason": "list_id_not_found",
                    "old_list_id": existing_list_id,
                    "old_account_number": None,
                    "fm_account_number": safe_account,
                }
            else:
                qb_account = (val_cust.findtext("AccountNumber") or "").strip().lower()
                if qb_account == safe_account.strip().lower():
                    trace.append(f"ListID validated OK -- QB AccountNumber={qb_account!r} matches FM, link still good")
                else:
                    trace.append(
                        f"ListID STALE -- QB AccountNumber={qb_account!r} does not match FM "
                        f"account_number={safe_account!r}; clearing link, will re-derive via AccountNumber scan"
                    )
                    stale_link = {
                        "cleared": True,
                        "reason": "account_number_mismatch",
                        "old_list_id": existing_list_id,
                        "old_account_number": val_cust.findtext("AccountNumber") or "",
                        "fm_account_number": safe_account,
                    }
        else:
            trace.append("no existing_list_id provided -- skipping ListID validation (first-time link/create)")

        # --- Standard AccountNumber-based match-or-create --------------------
        # Full fetch (ActiveStatus=All) -- QB has no AccountNumber filter, and
        # an existing-but-inactive match must still be treated as "exists"
        # (CustomerMod), not "not found" -- same reasoning as get_customer_by_account.
        #
        # Skipped entirely when the ListID pre-check above already validated a
        # current link (2026-08-11 perf fix): val_cust IS the matching record
        # in that case, so the full fetch + scan would just re-derive what we
        # already have. This is the common case once a customer has synced
        # once -- only a genuinely new-to-FM customer (no existing_list_id) or
        # a stale link (cleared above) still needs the full scan.
        #
        # 2026-08-20 perf fix: for that remaining case (genuinely first sync),
        # check the shared customer cache before paying for the full scan --
        # require_complete=True means we only trust a hit that came from an
        # ActiveStatus="All" fetch (see _customer_cache_populate), so this is
        # exactly as reliable as a live fetch here, just free when warm. This
        # is the path Cat/Bill's ~1-3min "customer sync" complaints traced to
        # (2026-08-20): the 08-11 fix only sped up re-syncing an
        # already-linked customer, never this first-time-creation path.
        # `root` stays None on a cache hit -- nothing was fetched, so the
        # 3100-duplicate-name handler below falls back to a live fetch only
        # in that rare case.
        root = None
        if existing_list_id and stale_link is None:
            trace.append(f"ListID pre-check already confirmed a valid link -- skipping full customer-list fetch, reusing ListID={existing_list_id!r}")
            existing = val_cust
        else:
            cache_hit, cached_customer = (
                _customer_cache_lookup(slug, safe_account, require_complete=True) if slug else (False, None)
            )
            if cache_hit and cached_customer is not None:
                trace.append(f"customer cache hit (ListID={cached_customer['list_id']!r}) -- skipping full fetch, re-querying just this record for a current EditSequence")
                fresh_resp = rp.ProcessRequest(ticket, build_customer_query_by_list_id(cached_customer["list_id"]))
                fresh_root = ET.fromstring(fresh_resp)
                existing = fresh_root.find(".//CustomerRet")
            elif cache_hit:
                trace.append("customer cache hit -- no match, skipping full fetch, proceeding to CustomerAdd")
                existing = None
            else:
                trace.append("customer cache miss/stale -- fetching full customer list (ActiveStatus=All) for AccountNumber scan")
                response = rp.ProcessRequest(ticket, build_customer_list_query("All"))
                root = ET.fromstring(response)
                target = safe_account.strip().lower()
                existing = None
                fetched_customers = []
                for cust in root.findall(".//CustomerRet"):
                    if cust.find("ParentRef") is not None:
                        continue
                    acct = cust.findtext("AccountNumber") or ""
                    if acct.strip().lower() == target:
                        existing = cust
                    if acct:
                        fetched_customers.append({
                            "list_id":        cust.findtext("ListID") or "",
                            "full_name":      _ascii_safe(cust.findtext("FullName") or ""),
                            "account_number": acct,
                        })
                if slug:
                    _customer_cache_populate(slug, fetched_customers, complete=True)

        if existing is not None:
            list_id = existing.findtext("ListID") or ""
            edit_sequence = existing.findtext("EditSequence") or ""
            trace.append(f"issuing CustomerMod (ListID={list_id!r})")
            mod_xml = build_customer_mod({
                **fields,
                "list_id": list_id,
                "edit_sequence": edit_sequence,
                "account_number": account_number,
            })
            rs = None
            for attempt in range(1, _LIST_CONFLICT_MAX_ATTEMPTS + 1):
                resp = rp.ProcessRequest(ticket, mod_xml)
                root2 = ET.fromstring(resp)
                rs = root2.find(".//CustomerModRs")
                code = rs.get("statusCode", "") if rs is not None else ""
                if code == "0":
                    break
                if code in _LIST_CONFLICT_CODES and attempt < _LIST_CONFLICT_MAX_ATTEMPTS:
                    trace.append(f"CustomerMod hit transient list-lock conflict (code={code}), retrying (attempt {attempt}/{_LIST_CONFLICT_MAX_ATTEMPTS})")
                    time.sleep(_LIST_CONFLICT_RETRY_DELAY)
                    continue
                if code == _STALE_EDIT_SEQUENCE_CODE and attempt < _LIST_CONFLICT_MAX_ATTEMPTS:
                    # Unlike 3180/3170, a stale EditSequence (someone/something
                    # else modified this record after we fetched ours) can't be
                    # fixed by resending the same request -- QB will reject the
                    # same edit_sequence again. Re-query the record for its
                    # current EditSequence and rebuild the Mod before retrying.
                    trace.append(f"CustomerMod hit stale edit sequence (code={code}), re-querying ListID={list_id!r} for current EditSequence, retrying (attempt {attempt}/{_LIST_CONFLICT_MAX_ATTEMPTS})")
                    refresh_resp = rp.ProcessRequest(ticket, build_customer_query_by_list_id(list_id))
                    refresh_root = ET.fromstring(refresh_resp)
                    refresh_cust = refresh_root.find(".//CustomerRet")
                    edit_sequence = (refresh_cust.findtext("EditSequence") or "") if refresh_cust is not None else edit_sequence
                    mod_xml = build_customer_mod({
                        **fields,
                        "list_id": list_id,
                        "edit_sequence": edit_sequence,
                        "account_number": account_number,
                    })
                    time.sleep(_LIST_CONFLICT_RETRY_DELAY)
                    continue
                break
            if rs is None or rs.get("statusCode", "") != "0":
                msg  = rs.get("statusMessage", "no response") if rs is not None else "no response"
                code = rs.get("statusCode", "?") if rs is not None else "?"
                trace.append(f"CustomerMod FAILED: code={code} msg={msg!r}")
                raise RuntimeError(
                    f"QuickBooks rejected the customer update: {msg} (code {code}) | trace: {' > '.join(trace)}"
                )
            trace.append("CustomerMod succeeded")
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
                "stale_link": stale_link,
                "trace": trace,
            }

        # Not found -- create.
        trace.append("AccountNumber scan: no match -- issuing CustomerAdd")
        add_xml = build_customer_add({**fields, "account_number": account_number})
        rs = None
        for attempt in range(1, _LIST_CONFLICT_MAX_ATTEMPTS + 1):
            resp = rp.ProcessRequest(ticket, add_xml)
            root2 = ET.fromstring(resp)
            rs = root2.find(".//CustomerAddRs")
            code = rs.get("statusCode", "") if rs is not None else ""
            if code in _LIST_CONFLICT_CODES and attempt < _LIST_CONFLICT_MAX_ATTEMPTS:
                trace.append(f"CustomerAdd hit transient list-lock conflict (code={code}), retrying (attempt {attempt}/{_LIST_CONFLICT_MAX_ATTEMPTS})")
                time.sleep(_LIST_CONFLICT_RETRY_DELAY)
                continue
            break
        if rs is None:
            trace.append("CustomerAdd FAILED: no CustomerAddRs in response")
            raise RuntimeError(f"No CustomerAddRs in QB response. | trace: {' > '.join(trace)}")

        code = rs.get("statusCode", "")
        if code == "0":
            trace.append("CustomerAdd succeeded")
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
                "stale_link": stale_link,
                "trace": trace,
            }
        elif code == "3100":
            # Name collision -- look up the conflicting record by Name so FM
            # can show the user its account number (Bill, 2026-08-11: "enough
            # information to find and address the issue"), instead of a bare
            # "name already in use" message.
            #
            # Deliberately does NOT issue a new NameFilter query here (tried
            # once, 2026-08-11 -- NameFilter + ActiveStatus=All made QB reject
            # the whole request as unparseable, reproduced live by Bill; see
            # build_customer_name_filter_query()'s docstring). Instead this
            # searches `root` -- the full customer/job list (ActiveStatus=All)
            # already fetched above for the AccountNumber scan -- which uses
            # the same query shape build_customer_list_query() has always used
            # successfully. Correct as long as nothing else could have created
            # the conflicting record in QB between that fetch and this Add
            # attempt, which holds here since both happen within this same
            # single QB session with no intervening writes.
            #
            # Only searches Customers/Jobs -- QB actually enforces Name
            # uniqueness across Customers, Vendors, Employees, and Other Names
            # together, so a 3100 can come from any of those. Reproduced live
            # 2026-08-11 (Bill, testing against "KVC, Inc." -- likely an
            # existing Vendor). Deliberately NOT chasing that down with new
            # VendorQueryRq/EmployeeQueryRq/OtherNameQueryRq support -- Bill's
            # call, 2026-08-11: expected to be rare, and the person running
            # this middleware may not even have QB permission to see Vendors
            # (confirmed true for Bill's own login), which would undermine an
            # automated search anyway. QB's own statusMessage is included
            # below instead -- cheap, and often already says something useful.
            if root is None:
                # 2026-08-20 perf fix took the cache-hit path above, so no
                # full fetch has happened yet this call -- only pay for one
                # now, in this rare conflict case where it's actually needed.
                trace.append("CustomerAdd FAILED: code=3100 (duplicate name) -- no full fetch on hand (cache-hit path), fetching now to find the conflict")
                conflict_resp = rp.ProcessRequest(ticket, build_customer_list_query("All"))
                root = ET.fromstring(conflict_resp)
            else:
                trace.append("CustomerAdd FAILED: code=3100 (duplicate name) -- searching already-fetched customer/job list for the conflict")
            conflict = None
            for cust in root.findall(".//CustomerRet"):
                if _ascii_safe(cust.findtext("Name") or "").lower() == safe_name.lower():
                    conflict = cust
                    break
            if conflict is None:
                trace.append("no exact Name match found among Customers/Jobs -- likely a Vendor/Employee/Other Name collision instead; conflict details will be blank, qb_message carries QB's own text")
            return {
                "action": "duplicate_name",
                "conflict": {
                    "list_id":        (conflict.findtext("ListID") if conflict is not None else "") or "",
                    "full_name":      _ascii_safe((conflict.findtext("FullName") if conflict is not None else "") or ""),
                    "account_number": (conflict.findtext("AccountNumber") if conflict is not None else "") or "",
                },
                "qb_message": rs.get("statusMessage", ""),
                "stale_link": stale_link,
                "trace": trace,
            }
        else:
            msg = rs.get("statusMessage", "Unknown error")
            trace.append(f"CustomerAdd FAILED: code={code} msg={msg!r}")
            raise RuntimeError(
                f"QuickBooks rejected customer creation for '{safe_name}': {msg} (code {code}) | trace: {' > '.join(trace)}"
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


def diagnostic_reactivate_customer(list_id: str) -> dict:
    """
    DIAGNOSTIC ONLY (2026-08-12), not wired into any production create/update
    path. Queries a customer by ListID, reports its current state, attempts
    CustomerMod with IsActive=true, and reports the result -- built to answer
    one specific open question empirically: does IsActive=true also resolve
    QB's separate "deleted" state (distinct from plain inactive -- confirmed
    by Bill, 2026-08-12: QB Desktop shows a red-X marker and a distinct
    "would you like to undelete it?" prompt for deleted customers), or does
    it only reactivate plain-inactive ones? See build_customer_reactivate()
    and the /fm/debug-reactivate-customer route.

    Returns {"found": False} if the ListID doesn't resolve to anything at
    all, otherwise {"found": True, "before": {...}, "mod_status_code": "...",
    "mod_status_message": "...", "after": {...} or None, "raw_mod_response": "..."}.
    """
    rp, ticket = _open_session()
    try:
        before_resp = rp.ProcessRequest(ticket, build_customer_query_by_list_id(list_id))
        before_root = ET.fromstring(before_resp)
        before_cust = before_root.find(".//CustomerRet")
        if before_cust is None:
            return {"found": False, "list_id": list_id}

        edit_sequence = before_cust.findtext("EditSequence") or ""
        before_state = {
            "full_name": before_cust.findtext("FullName") or "",
            "account_number": before_cust.findtext("AccountNumber") or "",
            "is_active": before_cust.findtext("IsActive") or "",
        }

        mod_resp = rp.ProcessRequest(ticket, build_customer_reactivate(list_id, edit_sequence))
        mod_root = ET.fromstring(mod_resp)
        rs = mod_root.find(".//CustomerModRs")
        status_code = rs.get("statusCode", "?") if rs is not None else "?"
        status_msg  = rs.get("statusMessage", "") if rs is not None else ""

        after_cust = mod_root.find(".//CustomerRet")
        after_state = None
        if after_cust is not None:
            after_state = {
                "full_name": after_cust.findtext("FullName") or "",
                "account_number": after_cust.findtext("AccountNumber") or "",
                "is_active": after_cust.findtext("IsActive") or "",
            }

        return {
            "found": True,
            "before": before_state,
            "mod_status_code": status_code,
            "mod_status_message": status_msg,
            "after": after_state,
            "raw_mod_response": mod_resp,
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
