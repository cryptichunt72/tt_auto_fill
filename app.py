import os, json, io, datetime
from typing import Dict, Any, List, Tuple
from flask import Flask, request, render_template, jsonify, send_file, abort
from dotenv import load_dotenv
import requests

from docxtpl import DocxTemplate
from docx.shared import Mm

load_dotenv()

app = Flask(__name__)

NOTION_TOKEN = os.getenv("NOTION_TOKEN", "")
REMITTER_DB  = os.getenv("REMITTER_DATABASE_ID", "")
BENEFICIARY_DB = os.getenv("BENEFICIARY_DATABASE_ID", "")
TT_TEMPLATE = os.getenv("TT_TEMPLATE", "TT_Form_template.docx")

NOTION_HEADERS = {
    "Authorization": f"Bearer {NOTION_TOKEN}",
    "Notion-Version": "2022-06-28",
    "Content-Type": "application/json"
}

def _assert_env():
    if not NOTION_TOKEN.startswith("secret_"):
        raise RuntimeError("Set NOTION_TOKEN for an internal integration (starts with 'secret_').")
    if len(REMITTER_DB) < 30 or len(BENEFICIARY_DB) < 30:
        raise RuntimeError("Set both REMITTER_DATABASE_ID and BENEFICIARY_DATABASE_ID as 32-char IDs (no dashes).")

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
    url = f"https://api.notion.com/v1/databases/{dbid}/query"
    results = []
    payload = {}
    while True:
        r = requests.post(url, headers=NOTION_HEADERS, json=payload, timeout=30)
        r.raise_for_status()
        data = r.json()
        results.extend(data.get("results", []))
        if not data.get("has_more"):
            break
        payload["start_cursor"] = data.get("next_cursor")
    return results

def get_rich(prop):
    arr = prop.get("rich_text", [])
    return "".join([x.get("plain_text","") for x in arr]) if arr else ""

def get_title(prop):
    arr = prop.get("title", [])
    return "".join([x.get("plain_text","") for x in arr]) if arr else ""

def get_phone(prop):
    return prop.get("phone_number") or ""

def get_select(prop):
    obj = prop.get("select") or {}
    return obj.get("name","")

def parse_remitter(page: Dict[str, Any], title_prop: str) -> Dict[str, Any]:
    props = page.get("properties", {})
    return {
        "id": page.get("id"),
        "name": get_title(props.get(title_prop, {})),
        "account_no": get_rich(props.get("Account No", {})),
        "address": get_rich(props.get("Address", {})),
        "phone": get_phone(props.get("Phone", {})),
        "id_type": get_select(props.get("ID Type", {})),
        "id_value": get_rich(props.get("ID Value", {})),
    }

def parse_beneficiary(page: Dict[str, Any], title_prop: str) -> Dict[str, Any]:
    props = page.get("properties", {})
    return {
        "id": page.get("id"),
        "beneficiary_name": get_title(props.get(title_prop, {})),
        "beneficiary_account_number": get_rich(props.get("Beneficiary Account Number", {})),
        "beneficiary_address": get_rich(props.get("Beneficiary Address", {})),
        "beneficiary_country": get_select(props.get("Beneficiary Country", {})),
        "beneficiary_bank_name": get_rich(props.get("Beneficiary Bank Name", {})),
        "beneficiary_bank_address": get_rich(props.get("Beneficiary Bank Address", {})),
        "beneficiary_bank_country": get_select(props.get("Beneficiary Bank Country", {})),
        "beneficiary_bank_swift": get_rich(props.get("Beneficiary Bank SWIFT", {})),
        "intermediary_bank_name": get_rich(props.get("Intermediary Bank Name", {})),
        "intermediary_bank_address": get_rich(props.get("Intermediary Bank Address", {})),
        "intermediary_bank_swift": get_rich(props.get("Intermediary Bank SWIFT", {})),
    }

def _title(v):  return {"title": [{"text": {"content": str(v)}}]}
def _rich(v):   return {"rich_text": [{"text": {"content": str(v)}}]}
def _phone(v):  return {"phone_number": str(v) if v is not None else ""}
def _select(v): return {"select": {"name": str(v)}} if v else {"select": None}

def build_remitter_properties(row: Dict[str, Any], title_prop: str) -> Dict[str, Any]:
    props = {}
    props[title_prop] = _title(row.get("name",""))
    props["Account No"] = _rich(row.get("account_no",""))
    props["Address"]    = _rich(row.get("address",""))
    props["Phone"]      = _phone(row.get("phone",""))
    props["ID Type"]    = _select(row.get("id_type",""))
    props["ID Value"]   = _rich(row.get("id_value",""))
    return props

def build_beneficiary_properties(row: Dict[str, Any], title_prop: str) -> Dict[str, Any]:
    props = {}
    props[title_prop] = _title(row.get("beneficiary_name",""))
    r = lambda k: _rich(row.get(k,""))
    s = lambda k: _select(row.get(k,""))
    props["Beneficiary Account Number"] = r("beneficiary_account_number")
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

def notion_create_page(dbid: str, properties: Dict[str, Any]) -> Dict[str, Any]:
    payload = {"parent": {"database_id": dbid}, "properties": properties}
    r = requests.post("https://api.notion.com/v1/pages", headers=NOTION_HEADERS, json=payload, timeout=30)
    if r.status_code >= 300:
        return {"error": r.text, "status": r.status_code}
    return r.json()

def notion_update_page(page_id: str, properties: Dict[str, Any]) -> Dict[str, Any]:
    r = requests.patch(f"https://api.notion.com/v1/pages/{page_id}", headers=NOTION_HEADERS, json={"properties": properties}, timeout=30)
    if r.status_code >= 300:
        return {"error": r.text, "status": r.status_code}
    return r.json()

@app.route("/")
def index():
    try:
        _assert_env()
    except Exception as e:
        return render_template("index.html", env_error=str(e))
    return render_template("index.html", env_error=None)

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
        res = notion_update_page(page_id, props)
    else:
        target_db = REMITTER_DB if rtype=="remitter" else BENEFICIARY_DB
        res = notion_create_page(target_db, props)

    if "error" in res:
        return jsonify({"ok": False, "error": res["error"], "status": res.get("status", 500)}), 400
    return jsonify({"ok": True, "page": res})

@app.post("/generate")
def generate():
    # Expects JSON: { beneficiary:{...}, remitter:{...}, extra:{...} }
    data = request.get_json(force=True)
    beneficiary = data.get("beneficiary", {})
    remitter = data.get("remitter", {})
    extra = data.get("extra", {})

    context = {
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
        "currency": extra.get("currency","USD"),
        "amount_figures": extra.get("amount_figures",""),
        "charges": extra.get("charges","SHA"),
        "account_to_be_debited_number_currency": extra.get("account_to_be_debited_number_currency",""),
        "notes": extra.get("notes","")
    }

    try:
        if os.path.exists(TT_TEMPLATE):
            tpl = DocxTemplate(TT_TEMPLATE)
            tpl.render(context)
            buf = io.BytesIO()
            tpl.save(buf)
            buf.seek(0)
            return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                             as_attachment=True, download_name="TT_Filled.docx")
        else:
            raise FileNotFoundError(TT_TEMPLATE)
    except Exception:
        # Fallback: simple generated document
        from docx import Document
        from docx.shared import Pt
        doc = Document()
        doc.add_heading("TT Remittance Form (Auto-Filled)", level=1)

        def add_section(title, kv: List[Tuple[str,str]]):
            p = doc.add_paragraph()
            p.add_run(title).bold = True
            for k,v in kv:
                r = doc.add_paragraph().add_run(f"{k}: {v}")
                r.font.size = Pt(10)

        add_section("Beneficiary", [
            ("Account Number", context["beneficiary_account_number"]),
            ("Name", context["beneficiary_name"]),
            ("Address", context["beneficiary_address"]),
            ("Country", context["beneficiary_country"]),
        ])
        add_section("Beneficiary Bank", [
            ("Name", context["beneficiary_bank_name"]),
            ("Address", context["beneficiary_bank_address"]),
            ("Country", context["beneficiary_bank_country"]),
            ("SWIFT", context["beneficiary_bank_swift"]),
        ])
        add_section("Intermediary Bank", [
            ("Name", context["intermediary_bank_name"]),
            ("Address", context["intermediary_bank_address"]),
            ("SWIFT", context["intermediary_bank_swift"]),
        ])
        add_section("Remitter", [
            ("Name", context["remitter_name"]),
            ("Account No.", context["remitter_account_no"]),
            ("Address", context["remitter_address"]),
            ("Phone", context["remitter_phone"]),
            ("ID Type", context["remitter_id_type"]),
            ("ID Value", context["remitter_id_value"]),
        ])
        add_section("Payment Details", [
            ("Date", context["date"]),
            ("Currency", context["currency"]),
            ("Amount (figures)", context["amount_figures"]),
            ("Charges", context["charges"]),
            ("Account to be Debited (No & CCY)", context["account_to_be_debited_number_currency"]),
            ("Notes", context["notes"]),
        ])

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                         as_attachment=True, download_name="TT_Filled_Simple.docx")

@app.get("/healthz")
def healthz():
    return {"ok": True}

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
