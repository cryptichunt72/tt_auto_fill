import os, io, datetime, traceback
from pathlib import Path
from typing import Dict, Any, List
from flask import Flask, request, render_template, jsonify, send_file, abort
from dotenv import load_dotenv
import requests
from docxtpl import DocxTemplate
from jinja2 import Environment
from num2words import num2words

# ──────────────────────────────────────────────────────────────────────────────
# Load env next to this file (works locally and on Render)
# ──────────────────────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
load_dotenv(dotenv_path=BASE_DIR / ".env")

app = Flask(__name__)

# ─── Config from env ──────────────────────────────────────────────────────────
NOTION_TOKEN   = os.getenv("NOTION_TOKEN", "")
REMITTER_DB    = os.getenv("REMITTER_DATABASE_ID", "")
BENEFICIARY_DB = os.getenv("BENEFICIARY_DATABASE_ID", "")
TT_TEMPLATE_RAW = os.getenv("TT_TEMPLATE", "")
PORT           = int(os.getenv("PORT", "5055"))

NOTION_HEADERS = {
    "Authorization": f"Bearer {NOTION_TOKEN}",
    "Notion-Version": "2022-06-28",
    "Content-Type": "application/json",
}

# ──────────────────────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────────────────────
def _template_path() -> Path:
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
    _ = _template_path()

# ──────────────────────────────────────────────────────────────────────────────
# Notion utilities
# ──────────────────────────────────────────────────────────────────────────────
def notion_get_database(dbid: str) -> Dict[str, Any]:
    r = requests.get(f"https://api.notion.com/v1/databases/{dbid}", headers=NOTION_HEADERS, timeout=30)
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

def get_rich(p):  arr = p.get("rich_text", []); return "".join([x.get("plain_text","") for x in arr]) if arr else ""
def get_title(p): arr = p.get("title", []);    return "".join([x.get("plain_text","") for x in arr]) if arr else ""
def get_phone(p): return p.get("phone_number") or ""
def get_select(p):obj = p.get("select") or {}; return obj.get("name","")

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

# ──────────────────────────────────────────────────────────────────────────────
# Template helpers
# ──────────────────────────────────────────────────────────────────────────────
def list_template_vars(docx_path: Path) -> List[str]:
    tpl = DocxTemplate(str(docx_path))
    try:
        env = Environment()
        return sorted(list(tpl.get_undeclared_template_variables(env)))
    except Exception:
        return []

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
        major_unit, minor_unit = "Rupees", "Paise"
    elif cur in ("USD","CAD","AUD","NZD","SGD","HKD"):
        major_unit, minor_unit = "Dollars","Cents"
    elif cur == "EUR":
        major_unit, minor_unit = "Euros","Cents"
    elif cur == "GBP":
        major_unit, minor_unit = "Pounds","Pence"
    elif cur in ("AED","SAR","QAR","OMR","BHD","KWD"):
        major_unit, minor_unit = "Dirhams" if cur=="AED" else "Rials" if cur in ("SAR","QAR","OMR") else "Dinars", "Fils"
    else:
        major_unit, minor_unit = "", ""
    if minor:
        return f"{major_words.title()} {major_unit} And {minor_words.title()} {minor_unit} Only"
    return f"{major_words.title()} {major_unit} Only"

# ──────────────────────────────────────────────────────────────────────────────
# API: generate DOCX
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
    amount_figures_text = amount_to_words(amount_figures, currency)

    ctx = {
        "beneficiary_account_number": beneficiary.get("beneficiary_account_number",""),
        "beneficiary_name": beneficiary.get("beneficiary_name",""),
        "beneficiary_address": beneficiary.get("beneficiary_address",""),
        "beneficiary_country": beneficiary.get("beneficiary_country",""),
        "beneficiary_bank_name": beneficiary.get("beneficiary_bank_name",""),
        "beneficiary_bank_address": beneficiary.get("beneficiary_bank_address",""),
        "beneficiary_bank_country": beneficiary.get("beneficiary_bank_country",""),
        "beneficiary_bank_swift": beneficiary.get("beneficiary_bank_swift",""),
        "intermediary_bank_name": beneficiary.get("intermediary_bank_name",""),
        "intermediary_bank_address": beneficiary.get("intermediary_bank_address",""),
        "intermediary_bank_swift": beneficiary.get("intermediary_bank_swift",""),
        "remitter_name": remitter.get("name",""),
        "remitter_account_no": remitter.get("account_no",""),
        "remitter_address": remitter.get("address",""),
        "remitter_phone": remitter.get("phone",""),
        "remitter_id_type": remitter.get("id_type",""),
        "remitter_id_value": remitter.get("id_value",""),
        "date": extra.get("date") or str(datetime.date.today()),
        "currency": currency,
        "amount_figures": amount_figures,
        "amount_figures_text": amount_figures_text,
        "charges": extra.get("charges","SHA"),
        "notes": extra.get("notes",""),
    }

    try:
        tpl_path = _template_path()
        tpl = DocxTemplate(str(tpl_path))

        # FIXED: case-insensitive check
        wanted = list_template_vars(tpl_path)
        render_ctx = {**ctx, **{k.upper(): v for k, v in ctx.items()}}
        keyset = {k.upper() for k in render_ctx}
        missing = [m for m in wanted if m.upper() not in keyset]
        if missing:
            return {"ok": False, "error": "Missing variables for template", "missing_variables": missing}, 400

        tpl.render(render_ctx)
        buf = io.BytesIO()
        tpl.save(buf)
        buf.seek(0)

        fname = f"Remittance_{beneficiary.get('beneficiary_name','Unknown')}_{currency}_{amount_figures}_{datetime.date.today()}.docx"
        return send_file(buf, as_attachment=True, download_name=fname,
                         mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    except Exception as e:
        print("TEMPLATE ERROR:\n", traceback.format_exc())
        return {"ok": False, "error": f"{type(e).__name__}: {e}"}, 500

@app.get("/healthz")
def healthz():
    return {"ok": True}, 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=PORT, debug=True)
