import os, io, datetime, traceback
from pathlib import Path
from typing import Dict, Any, List
from flask import Flask, request, render_template, jsonify, send_file, abort
from dotenv import load_dotenv
import requests
from docxtpl import DocxTemplate
from jinja2 import Environment  # <-- NEW: for template var discovery
from num2words import num2words

# Load .env from this folder
load_dotenv(dotenv_path=Path(__file__).with_name(".env"))

app = Flask(__name__)

# ─── Config from .env ─────────────────────────────────────────────────────────
NOTION_TOKEN   = os.getenv("NOTION_TOKEN", "")
REMITTER_DB    = os.getenv("REMITTER_DATABASE_ID", "")
BENEFICIARY_DB = os.getenv("BENEFICIARY_DATABASE_ID", "")
TT_TEMPLATE    = os.getenv("TT_TEMPLATE", "")          # absolute path to your TT DOCX
PORT           = int(os.getenv("PORT", "5055"))

NOTION_HEADERS = {
    "Authorization": f"Bearer {NOTION_TOKEN}",
    "Notion-Version": "2022-06-28",
    "Content-Type": "application/json",
}

# ─── Guards ───────────────────────────────────────────────────────────────────
def _assert_env():
    if not NOTION_TOKEN:
        raise RuntimeError("NOTION_TOKEN missing in .env")
    if not (NOTION_TOKEN.startswith("secret_") or NOTION_TOKEN.startswith("ntn_")):
        raise RuntimeError("NOTION_TOKEN must start with 'secret_' or 'ntn_'")
    if len(REMITTER_DB) != 32 or len(BENEFICIARY_DB) != 32:
        raise RuntimeError("Database IDs must be 32 chars (no dashes)")
    if not TT_TEMPLATE:
        raise RuntimeError("TT_TEMPLATE missing in .env")
    if not os.path.isabs(TT_TEMPLATE):
        raise RuntimeError("TT_TEMPLATE must be an absolute path")
    if not os.path.exists(TT_TEMPLATE):
        raise RuntimeError(f"Template not found: {TT_TEMPLATE}")

# ─── Notion helpers ───────────────────────────────────────────────────────────
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

# ─── Template utilities (version-safe) ────────────────────────────────────────
def list_template_vars(docx_path: str) -> List[str]:
    """
    Return the set of undeclared variables found in the DOCX template, in a
    version-safe way across docxtpl releases.
    """
    tpl = DocxTemplate(docx_path)
    try:
        # Older docxtpl: requires a Jinja2 Environment argument
        env = Environment()
        return sorted(list(tpl.get_undeclared_template_variables(env)))
    except TypeError:
        # Newer docxtpl variants may not require env
        return sorted(list(tpl.get_undeclared_template_variables()))
    except Exception:
        # As a last resort, don't crash the app; report empty list
        return []

# ─── UI ───────────────────────────────────────────────────────────────────────
@app.get("/")
def ui_home():
    try:
        _assert_env()
    except Exception as e:
        return render_template("index.html", env_error=str(e))
    return render_template("index.html", env_error=None)

# ─── API: load options / read / upsert ────────────────────────────────────────
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
        res = requests.patch(f"https://api.notion.com/v1/pages/{page_id}",
                             headers=NOTION_HEADERS, json={"properties": props}, timeout=30)
    else:
        target_db = REMITTER_DB if rtype=="remitter" else BENEFICIARY_DB
        payload = {"parent": {"database_id": target_db}, "properties": props}
        res = requests.post("https://api.notion.com/v1/pages",
                            headers=NOTION_HEADERS, json=payload, timeout=30)

    if res.status_code >= 300:
        return jsonify({"ok": False, "error": res.text, "status": res.status_code}), 400
    return jsonify({"ok": True, "page": res.json()})

# ─── API: generate DOCX (download / overwrite / save to path) ────────────────
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

    amount_figures = extra.get("amount_figures", "")
    amount_figures_text = ""
    if amount_figures:
        try:
            # Convert to words safely
            amt = float(amount_figures.replace(",", "").strip())
            amount_figures_text = num2words(amt, to="cardinal", lang="en").replace("-", " ")
            # Capitalize nicely
            amount_figures_text = amount_figures_text.title() + " Only"
        except Exception:
            amount_figures_text = ""

    ctx = {
        # Beneficiary
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
        # Remitter
        "remitter_name": remitter.get("name",""),
        "remitter_account_no": remitter.get("account_no",""),
        "remitter_address": remitter.get("address",""),
        "remitter_phone": remitter.get("phone",""),
        "remitter_id_type": remitter.get("id_type",""),
        "remitter_id_value": remitter.get("id_value",""),
        # TT extras
        "date": extra.get("date") or str(datetime.date.today()),
        "currency": extra.get("currency","USD"),
        "amount_figures": amount_figures,
        "amount_figures_text": amount_figures_text,   # <-- NEW FIELD
        "charges": extra.get("charges","SHA"),
        "account_to_be_debited_number_currency": extra.get("account_to_be_debited_number_currency",""),
        "notes": extra.get("notes",""),
    }

    try:
        tpl = DocxTemplate(TT_TEMPLATE)

        # VERSION-SAFE missing-vars check
        missing = list_template_vars(TT_TEMPLATE)  # template placeholders
        # remove those keys present in ctx (if present, they're not "missing")
        missing = [m for m in missing if m not in ctx]
        if missing:
            return {
                "ok": False,
                "error": "Missing variables in context for this template.",
                "missing_variables": sorted(missing)
            }, 400

        tpl.render(ctx)

        overwrite = bool(extra.get("overwrite"))
        out_path  = extra.get("out_path")

        if overwrite:
            tpl.save(TT_TEMPLATE)
            return {"ok": True, "message": f"Overwrote template: {TT_TEMPLATE}"}

        if out_path:
            if not os.path.isabs(out_path):
                return {"ok": False, "error": "out_path must be an absolute path"}, 400
            tpl.save(out_path)
            return {"ok": True, "message": f"Saved to: {out_path}"}

        buf = io.BytesIO()
        tpl.save(buf)
        buf.seek(0)
        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name="TT_Filled.docx",
        )
    except Exception as e:
        print("TEMPLATE ERROR:\n", traceback.format_exc())
        return {"ok": False, "error": f"{type(e).__name__}: {e}"}, 500

# ─── Debug helpers ────────────────────────────────────────────────────────────
@app.get("/debug/template")
def debug_template():
    try:
        _assert_env()
        wanted = list_template_vars(TT_TEMPLATE)
        return jsonify({"ok": True, "template": TT_TEMPLATE, "expects": wanted})
    except Exception as e:
        return jsonify({"ok": False, "error": f"{type(e).__name__}: {e}"}), 400

# ─── UI template ──────────────────────────────────────────────────────────────
@app.get("/ui")
def ui_redirect():
    return render_template("index.html", env_error=None)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=PORT, debug=True)
