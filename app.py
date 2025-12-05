import os, io, datetime, traceback
from pathlib import Path
from typing import Dict, Any, List
from flask import Flask, request, render_template, jsonify, send_file, abort
from dotenv import load_dotenv
import requests
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from jinja2 import Environment
from num2words import num2words
from invoice_app import invoice_bp



# ──────────────────────────────────────────────────────────────────────────────
# Load env next to this file (works locally and on Render)
# ──────────────────────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
load_dotenv(dotenv_path=BASE_DIR / ".env")

app = Flask(__name__)
app.register_blueprint(invoice_bp)

# ─── Config from env ──────────────────────────────────────────────────────────
NOTION_TOKEN   = os.getenv("NOTION_TOKEN", "")
REMITTER_DB    = os.getenv("REMITTER_DATABASE_ID", "")
BENEFICIARY_DB = os.getenv("BENEFICIARY_DATABASE_ID", "")
TT_TEMPLATE_RAW = os.getenv("TT_TEMPLATE", "")     # relative or absolute; we resolve it
PORT           = int(os.getenv("PORT", "5055"))

NOTION_HEADERS = {
    "Authorization": f"Bearer {NOTION_TOKEN}",
    "Notion-Version": "2022-06-28",
    "Content-Type": "application/json",
}

# ──────────────────────────────────────────────────────────────────────────────
# Helpers: validate env + resolve template path + uppercase helper
# ──────────────────────────────────────────────────────────────────────────────
def _template_path() -> Path:
    """
    Accept relative ('TT_Form_template.docx') or absolute path.
    Resolve to an absolute Path and ensure it exists.
    """
    raw = TT_TEMPLATE_RAW.strip()
    if not raw:
        raise RuntimeError("TT_TEMPLATE missing.")
    p = Path(raw)
    if not p.is_absolute():
        p = (BASE_DIR / p).resolve()
    if not p.exists():
        raise RuntimeError(f"Template not found: {p}")
    return p

def _assert_env():
    if not NOTION_TOKEN:
        raise RuntimeError("NOTION_TOKEN missing.")
    if not (NOTION_TOKEN.startswith("secret_") or NOTION_TOKEN.startswith("ntn_")):
        raise RuntimeError("NOTION_TOKEN must start with 'secret_' or 'ntn_'.")
    if len(REMITTER_DB) != 32 or len(BENEFICIARY_DB) != 32:
        raise RuntimeError("Database IDs must be 32 characters (no dashes).")
    _ = _template_path()  # validates and resolves

def upper(v):
    """Uppercase helper for strings; leave non-strings untouched."""
    return v.upper() if isinstance(v, str) else v

# ──────────────────────────────────────────────────────────────────────────────
# Notion utilities
# ──────────────────────────────────────────────────────────────────────────────
def notion_get_database(dbid: str) -> Dict[str, Any]:
    r = requests.get(
        f"https://api.notion.com/v1/databases/{dbid}",
        headers=NOTION_HEADERS,
        timeout=30,
    )
    if r.status_code == 404:
        abort(404, f"Database not found or not shared: {dbid}")
    r.raise_for_status()
    return r.json()

def title_prop_name(db: Dict[str, Any]) -> str:
    for name, prop in db.get("properties", {}).items():
        if prop.get("type") == "title":
            return name
    return "Name"

def notion_query_all(dbid: str) -> List[Dict[str, Any]]:
    url, results, payload = f"https://api.notion.com/v1/databases/{dbid}/query", [], {}
    while True:
        r = requests.post(url, headers=NOTION_HEADERS, json=payload, timeout=30)
        r.raise_for_status()
        data = r.json()
        results.extend(data.get("results", []))
        if not data.get("has_more"):
            break
        payload["start_cursor"] = data.get("next_cursor")
    return results

def get_rich(p):
    arr = p.get("rich_text", [])
    return "".join([x.get("plain_text", "") for x in arr]) if arr else ""

def get_title(p):
    arr = p.get("title", [])
    return "".join([x.get("plain_text", "") for x in arr]) if arr else ""

def get_phone(p):
    return p.get("phone_number") or ""

def get_select(p):
    obj = p.get("select") or {}
    return obj.get("name", "")

def get_file_url(p):
    """
    Extract first file URL from a Notion 'files' property.
    Used for the signature image.
    """
    files = p.get("files", [])
    if not files:
        return ""
    f = files[0]
    t = f.get("type")
    if t == "file":
        return f.get("file", {}).get("url", "")
    if t == "external":
        return f.get("external", {}).get("url", "")
    return ""

def parse_remitter(page: Dict[str, Any], t: str) -> Dict[str, Any]:
    p = page.get("properties", {})
    return {
        "id": page.get("id"),
        "name": get_title(p.get(t, {})),
        "account_no": get_rich(p.get("Account No", {})),
        "address": get_rich(p.get("Address", {})),
        "phone": get_phone(p.get("Phone", {})),
        "id_type": get_select(p.get("ID Type", {})),
        "id_value": get_rich(p.get("ID Value", {})),
        # Signature image: Files & media property named "Signature"
        "signature_url": get_file_url(p.get("Signature", {})),
    }

def parse_beneficiary(page: Dict[str, Any], t: str) -> Dict[str, Any]:
    p = page.get("properties", {})
    return {
        "id": page.get("id"),
        "beneficiary_name": get_title(p.get(t, {})),
        "beneficiary_account_number": get_rich(p.get("Beneficiary Account Number", {})),
        "beneficiary_address": get_rich(p.get("Beneficiary Address", {})),
        "beneficiary_country": get_select(p.get("Beneficiary Country", {})),
        "beneficiary_bank_name": get_rich(p.get("Beneficiary Bank Name", {})),
        "beneficiary_bank_address": get_rich(p.get("Beneficiary Bank Address", {})),
        "beneficiary_bank_country": get_select(p.get("Beneficiary Bank Country", {})),
        "beneficiary_bank_swift": get_rich(p.get("Beneficiary Bank SWIFT", {})),
        "intermediary_bank_name": get_rich(p.get("Intermediary Bank Name", {})),
        "intermediary_bank_address": get_rich(p.get("Intermediary Bank Address", {})),
        "intermediary_bank_swift": get_rich(p.get("Intermediary Bank SWIFT", {})),
    }

def _title(v):  return {"title": [{"text": {"content": str(v)}}]}
def _rich(v):   return {"rich_text": [{"text": {"content": str(v)}}]}
def _phone(v):  return {"phone_number": str(v) if v is not None else ""}
def _select(v): return {"select": {"name": str(v)}} if v else {"select": None}

def build_remitter_properties(row: Dict[str, Any], title_prop: str) -> Dict[str, Any]:
    props = {}
    # Uppercase all relevant text fields (not numbers)
    props[title_prop] = _title(upper(row.get("name","")))
    props["Account No"] = _rich(row.get("account_no",""))  # keep numbers as-is
    props["Address"]    = _rich(upper(row.get("address","")))
    props["Phone"]      = _phone(row.get("phone",""))      # keep phone numeric
    props["ID Type"]    = _select(upper(row.get("id_type","")))
    props["ID Value"]   = _rich(upper(row.get("id_value","")))
    return props

def build_beneficiary_properties(row: Dict[str, Any], title_prop: str) -> Dict[str, Any]:
    props = {}
    # Uppercase title/name
    props[title_prop] = _title(upper(row.get("beneficiary_name","")))

    r = lambda k: _rich(upper(row.get(k,"")))
    s = lambda k: _select(upper(row.get(k,"")))

    # Keep pure numeric/ID fields as-is where appropriate
    props["Beneficiary Account Number"] = _rich(row.get("beneficiary_account_number",""))
    props["Beneficiary Address"]        = r("beneficiary_address")
    props["Beneficiary Country"]        = s("beneficiary_country")
    props["Beneficiary Bank Name"]      = r("beneficiary_bank_name")
    props["Beneficiary Bank Address"]   = r("beneficiary_bank_address")
    props["Beneficiary Bank Country"]   = s("beneficiary_bank_country")
    props["Beneficiary Bank SWIFT"]     = r("beneficiary_bank_swift")
    props["Intermediary Bank Name"]     = r("intermediary_bank_name")
    props["Intermediary Bank Address"]  = r("intermediary_bank_address")
    props["Intermediary Bank SWIFT"]    = r("intermediary_bank_swift")
    return props

# ──────────────────────────────────────────────────────────────────────────────
# Template utilities (version-safe missing var detection) + signature helpers
# ──────────────────────────────────────────────────────────────────────────────
def list_template_vars(docx_path: Path) -> List[str]:
    tpl = DocxTemplate(str(docx_path))
    try:
        env = Environment()
        return sorted(list(tpl.get_undeclared_template_variables(env)))
    except TypeError:
        return sorted(list(tpl.get_undeclared_template_variables()))
    except Exception:
        return []

def build_signature_image(tpl: DocxTemplate, url: str):
    """
    Download a signature image from Notion and wrap it as InlineImage
    for insertion into the Word template.
    """
    if not url:
        print("No signature URL found.")
        return ""
    try:
        print(f"Downloading signature from: {url}")
        r = requests.get(url, timeout=30)
        r.raise_for_status()
        img_bytes = io.BytesIO(r.content)
        # adjust width as needed
        return InlineImage(tpl, img_bytes, width=Mm(25))
    except Exception as e:
        print("SIGNATURE DOWNLOAD ERROR:", e)
        return ""

def find_remitter_signature_by_name(name: str) -> str:
    """
    Fallback when remitter.id is not present in the POST body.
    Searches the Remitter DB by title (name) and returns the Signature URL.
    """
    if not name:
        print("No remitter name provided; cannot search for signature.")
        return ""
    try:
        print(f"Searching Notion remitter DB for name: {name!r}")
        db = notion_get_database(REMITTER_DB)
        title_prop = title_prop_name(db)
        pages = notion_query_all(REMITTER_DB)
        target_upper = name.strip().upper()
        for page in pages:
            props = page.get("properties", {})
            page_name = get_title(props.get(title_prop, {}))
            if page_name and page_name.strip().upper() == target_upper:
                print(f"Matched remitter page by name: {page.get('id')}")
                return get_file_url(props.get("Signature", {}))
        print("No matching remitter name found in Notion DB.")
        return ""
    except Exception as e:
        print("ERROR searching remitter by name for signature:", e)
        return ""

def fetch_signature_url_from_notion(remitter: Dict[str, Any]) -> str:
    """
    Ensure we have a signature URL:
    1) Try remitter['signature_url'] from the frontend.
    2) If that is missing, try remitter['id'] to fetch the page.
    3) If id is missing, search the DB by remitter['name'].
    """
    # 1) Directly from body
    url = remitter.get("signature_url") or ""
    if url:
        print("Using signature_url from POST body.")
        return url

    # 2) Use page id if present
    page_id = remitter.get("id")
    if page_id:
        try:
            print(f"Fetching remitter page {page_id} from Notion for signature...")
            r = requests.get(
                f"https://api.notion.com/v1/pages/{page_id}",
                headers=NOTION_HEADERS,
                timeout=30,
            )
            r.raise_for_status()
            page = r.json()
            props = page.get("properties", {})
            url = get_file_url(props.get("Signature", {}))
            if url:
                print("Signature URL fetched from Notion page via id.")
            else:
                print("No Signature file found on Notion page (via id).")
            return url
        except Exception as e:
            print("SIGNATURE NOTION FETCH ERROR (via id):", e)

    # 3) Finally, fall back to name-based lookup
    print("No remitter.id, falling back to name-based lookup for signature.")
    name = remitter.get("name") or remitter.get("remitter_name") or ""
    return find_remitter_signature_by_name(name)

# ──────────────────────────────────────────────────────────────────────────────
# Amount-in-words helpers
# ──────────────────────────────────────────────────────────────────────────────
def amount_to_words(amount_str: str, currency_code: str = "USD") -> str:
    if not amount_str:
        return ""
    s = amount_str.strip().replace(",", "")
    try:
        parts = s.split(".")
        major = int(parts[0]) if parts[0] else 0
        minor = int((parts[1] + "00")[:2]) if len(parts) > 1 and parts[1].isdigit() else 0
    except Exception:
        return ""

    lang = "en_IN" if currency_code.upper() == "INR" else "en"
    major_words = num2words(major, to="cardinal", lang=lang).replace("-", " ")
    minor_words = num2words(minor, to="cardinal", lang=lang).replace("-", " ") if minor else ""

    cur = currency_code.upper()
    if cur == "INR":
        major_unit = "Rupee" if major == 1 else "Rupees"
        minor_unit = "Paisa" if minor == 1 else "Paise"
    elif cur in ("USD","CAD","AUD","NZD","SGD","HKD"):
        major_unit = "Dollar" if major == 1 else "Dollars"
        minor_unit = "Cent" if minor == 1 else "Cents"
    elif cur == "EUR":
        major_unit = "Euro" if major == 1 else "Euros"
        minor_unit = "Cent" if minor == 1 else "Cents"
    elif cur == "GBP":
        major_unit = "Pound" if major == 1 else "Pounds"
        minor_unit = "Pence"
    elif cur in ("AED","SAR","QAR","OMR","BHD","KWD"):
        major_unit = "Dirham" if cur == "AED" else ("Rial" if cur in ("SAR","QAR","OMR") else "Dinar")
        minor_unit = "Fils"
    else:
        major_unit, minor_unit = "", ""

    if minor and minor_unit:
        return f"{major_words.title()} {major_unit} And {minor_words.title()} {minor_unit} Only"
    if major_unit:
        return f"{major_words.title()} {major_unit} Only"
    # generic fallback
    return f"{major_words.title()} Only" if not minor else f"{major_words.title()} Point {minor_words.title()} Only"

# ──────────────────────────────────────────────────────────────────────────────
# UI
# ──────────────────────────────────────────────────────────────────────────────
@app.get("/")
def ui_home():
    try:
        _assert_env()
    except Exception as e:
        return render_template("index.html", env_error=str(e))
    return render_template("index.html", env_error=None)

# ──────────────────────────────────────────────────────────────────────────────
# API: load options / read / upsert
# ──────────────────────────────────────────────────────────────────────────────
@app.get("/api/options")
def api_options():
    _assert_env()
    rem_db = notion_get_database(REMITTER_DB)
    ben_db = notion_get_database(BENEFICIARY_DB)
    r_title = title_prop_name(rem_db)
    b_title = title_prop_name(ben_db)
    r_pages = notion_query_all(REMITTER_DB)
    b_pages = notion_query_all(BENEFICIARY_DB)
    remitters = [parse_remitter(p, r_title) for p in r_pages]
    beneficiaries = [parse_beneficiary(p, b_title) for p in b_pages]
    return jsonify({
        "remitters": [{"id": r["id"], "name": r["name"]} for r in remitters],
        "beneficiaries": [{"id": b["id"], "name": b["beneficiary_name"]} for b in beneficiaries],
        "title_props": {"remitter": r_title, "beneficiary": b_title}
    })

@app.get("/api/record/<rtype>/<page_id>")
def api_record(rtype, page_id):
    _assert_env()
    if rtype not in ("remitter","beneficiary"):
        abort(400, "rtype must be remitter or beneficiary")
    db = notion_get_database(REMITTER_DB if rtype=="remitter" else BENEFICIARY_DB)
    title = title_prop_name(db)
    r = requests.get(f"https://api.notion.com/v1/pages/{page_id}", headers=NOTION_HEADERS, timeout=30)
    r.raise_for_status()
    page = r.json()
    data = parse_remitter(page, title) if rtype=="remitter" else parse_beneficiary(page, title)
    return jsonify(data)

@app.post("/api/upsert/<rtype>")
def api_upsert(rtype):
    _assert_env()
    body = request.get_json(force=True)

    # Optional: normalize all incoming string fields in body to uppercase at the top level
    # (we still selectively handle numbers below)
    def upper_body(obj):
        if isinstance(obj, dict):
            return {k: (v.upper() if isinstance(v, str) else v) for k, v in obj.items()}
        return obj

    body = upper_body(body)

    if rtype == "remitter":
        db = notion_get_database(REMITTER_DB)
        title = title_prop_name(db)
        props = build_remitter_properties(body, title)
    elif rtype == "beneficiary":
        db = notion_get_database(BENEFICIARY_DB)
        title = title_prop_name(db)
        props = build_beneficiary_properties(body, title)
    else:
        abort(400, "rtype must be remitter or beneficiary")

    page_id = body.get("id")
    if page_id:
        res = requests.patch(
            f"https://api.notion.com/v1/pages/{page_id}",
            headers=NOTION_HEADERS,
            json={"properties": props},
            timeout=30,
        )
    else:
        target_db = REMITTER_DB if rtype=="remitter" else BENEFICIARY_DB
        payload = {"parent": {"database_id": target_db}, "properties": props}
        res = requests.post(
            "https://api.notion.com/v1/pages",
            headers=NOTION_HEADERS,
            json=payload,
            timeout=30,
        )

    if res.status_code >= 300:
        return jsonify({"ok": False, "error": res.text, "status": res.status_code}), 400
    return jsonify({"ok": True, "page": res.json()})

# ──────────────────────────────────────────────────────────────────────────────
# API: generate DOCX (download / overwrite / save to path)
# ──────────────────────────────────────────────────────────────────────────────
@app.post("/generate")
def generate():
    try:
        _assert_env()
    except Exception as e:
        return {"ok": False, "error": str(e)}, 400

    data        = request.get_json(force=True)
    beneficiary = data.get("beneficiary", {})
    remitter    = data.get("remitter", {})
    extra       = data.get("extra", {})

    currency       = (extra.get("currency") or "USD").strip().upper()
    amount_figures = (extra.get("amount_figures") or "").strip()
    amount_figures_text = upper(amount_to_words(amount_figures, currency))

    # Build context, uppercasing all textual user-entered data
    ctx = {
        # Beneficiary
        "beneficiary_account_number": beneficiary.get("beneficiary_account_number",""),
        "beneficiary_name": upper(beneficiary.get("beneficiary_name","")),
        "beneficiary_address": upper(beneficiary.get("beneficiary_address","")),
        "beneficiary_country": upper(beneficiary.get("beneficiary_country","")),
        "beneficiary_bank_name": upper(beneficiary.get("beneficiary_bank_name","")),
        "beneficiary_bank_address": upper(beneficiary.get("beneficiary_bank_address","")),
        "beneficiary_bank_country": upper(beneficiary.get("beneficiary_bank_country","")),
        "beneficiary_bank_swift": upper(beneficiary.get("beneficiary_bank_swift","")),
        "intermediary_bank_name": upper(beneficiary.get("intermediary_bank_name","")),
        "intermediary_bank_address": upper(beneficiary.get("intermediary_bank_address","")),
        "intermediary_bank_swift": upper(beneficiary.get("intermediary_bank_swift","")),
        # Remitter
        "remitter_name": upper(remitter.get("name","")),
        "remitter_account_no": remitter.get("account_no",""),
        "remitter_address": upper(remitter.get("address","")),
        "remitter_phone": remitter.get("phone",""),
        "remitter_id_type": upper(remitter.get("id_type","")),
        "remitter_id_value": upper(remitter.get("id_value","")),
        # TT extras
        "date": extra.get("date") or str(datetime.date.today()),
        "currency": currency,
        "amount_figures": amount_figures,
        "amount_figures_text": amount_figures_text,   # auto-generated words (kept as title case)
        "charges": upper(extra.get("charges","SHA")),
        "account_to_be_debited_number_currency": upper(extra.get("account_to_be_debited_number_currency","")),
        "notes": upper(extra.get("notes","")),
    }

    try:
        tpl_path = _template_path()
        tpl = DocxTemplate(str(tpl_path))

        # Get signature URL (from body, or Notion via id, or name-based lookup)
        signature_url = fetch_signature_url_from_notion(remitter)
        ctx["signature"] = build_signature_image(tpl, signature_url)

        # Optional: show missing template vars not present in ctx
        wanted = list_template_vars(tpl_path)
        print("Template expects variables:", wanted)
        missing = [m for m in wanted if m not in ctx]
        if missing:
            print("Missing variables for template:", missing)
            return {"ok": False, "error": "Missing variables for template", "missing_variables": missing}, 400

        tpl.render(ctx)

        overwrite = bool(extra.get("overwrite"))
        out_path  = extra.get("out_path")

        if overwrite:
            tpl.save(str(tpl_path))
            return {"ok": True, "message": f"Overwrote template: {tpl_path}"}

        if out_path:
            out_p = Path(out_path)
            if not out_p.is_absolute():
                out_p = (BASE_DIR / out_p).resolve()
            out_p.parent.mkdir(parents=True, exist_ok=True)
            tpl.save(str(out_p))
            return {"ok": True, "message": f"Saved to: {out_p}"}

        buf = io.BytesIO()
        tpl.save(buf)
        buf.seek(0)

        # Build custom file name (uppercase beneficiary name in filename as well)
        beneficiary_name_raw = (beneficiary.get("beneficiary_name") or "Unknown").strip().replace("/", "-")
        beneficiary_name = upper(beneficiary_name_raw)
        currency_code = currency.strip().upper()
        amount_value = (extra.get("amount_figures") or "0").strip().replace(",", "")
        created_date = (extra.get("date") or str(datetime.date.today())).strip()

        # Ensure safe characters in file name
        safe_name = f"FILE NAME – {beneficiary_name} – {currency_code} {amount_value} – {created_date}".replace(":", "-").replace("/", "-")

        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name=f"{safe_name}.docx",
        )
    except Exception as e:
        print("TEMPLATE ERROR:\n", traceback.format_exc())
        return {"ok": False, "error": f"{type(e).__name__}: {e}"}, 500

# ──────────────────────────────────────────────────────────────────────────────
# Debug helpers
# ──────────────────────────────────────────────────────────────────────────────
@app.get("/debug/template")
def debug_template():
    try:
        _assert_env()
        p = _template_path()
        wanted = list_template_vars(p)
        return jsonify({"ok": True, "template": str(p), "expects": wanted})
    except Exception as e:
        return jsonify({"ok": False, "error": f"{type(e).__name__}: {e}"}), 400

@app.get("/debug/env")
def debug_env():
    try:
        p = _template_path()
        tpl_exists = p.exists()
        tpl_str = str(p)
    except Exception:
        tpl_exists = False
        tpl_str = TT_TEMPLATE_RAW
    return {
        "token_prefix": (NOTION_TOKEN[:6]+"...") if NOTION_TOKEN else None,
        "rem_db_len": len(REMITTER_DB),
        "ben_db_len": len(BENEFICIARY_DB),
        "template_env": TT_TEMPLATE_RAW,
        "template_resolved": tpl_str,
        "template_exists": tpl_exists
    }

@app.get("/healthz")
def healthz():
    return {"ok": True}, 200

# ──────────────────────────────────────────────────────────────────────────────
# Entrypoint
# ──────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=PORT, debug=True)
