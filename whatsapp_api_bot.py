"""
WhatsApp Cloud API Bot — COD Label Generator
=============================================
This runs as a Flask webhook server. No browser needed.
Works with your phone's WhatsApp — send orders from anywhere.

SETUP (one-time):
1. Go to https://developers.facebook.com → Create App → Business → WhatsApp
2. In WhatsApp > API Setup:
   - Note your Phone Number ID, WhatsApp Business Account ID
   - Generate a permanent access token
3. Set up a webhook:
   - URL: https://your-server.com/webhook  (use ngrok for testing)
   - Verify token: pick any string, put it in VERIFY_TOKEN below
   - Subscribe to "messages" field
4. Fill in the config below and run:
   python whatsapp_api_bot.py

FREE TUNNEL (for testing without a server):
   ngrok http 5000
   Then paste the ngrok URL into Meta's webhook config.
"""

import re
import os
import json
import hashlib
import hmac
import time
import subprocess
import requests

from flask import Flask, request, jsonify

from docx import Document
from num2words import num2words
from copy import deepcopy

# Try importing docx2pdf (needs MS Word — only works on Windows with Word installed)
try:
    from docx2pdf import convert as docx2pdf_convert
except ImportError:
    docx2pdf_convert = None


# ============== CONFIG ==============
# Reads from environment variables first, falls back to hardcoded values.
# On Render, set these as Environment Variables in the dashboard.

WHATSAPP_TOKEN = os.environ.get("WHATSAPP_TOKEN", "")          # Your permanent access token
PHONE_NUMBER_ID = os.environ.get("PHONE_NUMBER_ID", "")        # Your WhatsApp Phone Number ID
VERIFY_TOKEN = os.environ.get("VERIFY_TOKEN", "cod_bot_verify")  # Webhook verification token
APP_SECRET = os.environ.get("APP_SECRET", "")                  # Optional: App secret for signature verification
ALLOWED_NUMBERS = []         # e.g. ["919342901848"] — leave empty to allow all
SEND_PROGRESS_REPLY = False  # False = don't send "Added X label(s)" after every order message

# ============== FILE CONFIG ==============

TEMPLATE = "cod_template.docx"
OUTPUT = "generated_labels.docx"
PROCESSED_FILE = "processed_orders_api.json"
BATCH_COUNTER_FILE = "batch_counter.txt"
STATE_FILE = "bot_state.json"

# ============== APP ==============

app = Flask(__name__)

DELIM = r"[,:;.\s]"


# ---------------- STATE ----------------

def load_state():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r") as f:
            return json.load(f)
    return {"collecting": False, "batch_count": 0}


def save_state(state):
    with open(STATE_FILE, "w") as f:
        json.dump(state, f)


def load_processed():
    if os.path.exists(PROCESSED_FILE):
        with open(PROCESSED_FILE, "r", encoding="utf-8") as f:
            return set(json.load(f))
    return set()


def save_processed(processed):
    with open(PROCESSED_FILE, "w", encoding="utf-8") as f:
        json.dump(list(processed), f)


def order_hash(data):
    key = f"{data['name']}|{data['phone']}|{data['pincode']}|{data['price']}|{data['item']}"
    return hashlib.sha256(key.encode()).hexdigest()[:16]


def get_next_batch_number():
    num = 1
    if os.path.exists(BATCH_COUNTER_FILE):
        with open(BATCH_COUNTER_FILE, "r") as f:
            try:
                num = int(f.read().strip()) + 1
            except ValueError:
                num = 1
    with open(BATCH_COUNTER_FILE, "w") as f:
        f.write(str(num))
    return num


# ---------------- WHATSAPP API ----------------

def send_message(to, text):
    """Send a text message via WhatsApp Cloud API."""
    url = f"https://graph.facebook.com/v21.0/{PHONE_NUMBER_ID}/messages"
    headers = {
        "Authorization": f"Bearer {WHATSAPP_TOKEN}",
        "Content-Type": "application/json",
    }
    payload = {
        "messaging_product": "whatsapp",
        "to": to,
        "type": "text",
        "text": {"body": text},
    }
    resp = requests.post(url, headers=headers, json=payload, timeout=30)
    if resp.status_code != 200:
        print(f"  Send message failed: {resp.text}")
    return resp.status_code == 200


def send_document(to, file_path, caption=""):
    """Upload and send a document via WhatsApp Cloud API.
    Returns (success, error_message)."""
    # Step 1: Upload media
    url_upload = f"https://graph.facebook.com/v21.0/{PHONE_NUMBER_ID}/media"
    headers = {"Authorization": f"Bearer {WHATSAPP_TOKEN}"}

    filename = os.path.basename(file_path)
    mime = "application/pdf" if file_path.endswith(".pdf") else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

    with open(file_path, "rb") as f:
        resp = requests.post(
            url_upload,
            headers=headers,
            files={"file": (filename, f, mime)},
            data={"messaging_product": "whatsapp"},
            timeout=60,
        )

    if resp.status_code != 200:
        err = f"Media upload failed: {resp.status_code} {resp.text}"
        print(f"  {err}")
        return False, err

    media_id = resp.json().get("id")
    if not media_id:
        err = f"Media upload response missing id: {resp.text}"
        print(f"  {err}")
        return False, err

    # Step 2: Send document message
    url_send = f"https://graph.facebook.com/v21.0/{PHONE_NUMBER_ID}/messages"
    payload = {
        "messaging_product": "whatsapp",
        "to": to,
        "type": "document",
        "document": {
            "id": media_id,
            "filename": filename,
        },
    }
    if caption:
        payload["document"]["caption"] = caption

    resp = requests.post(
        url_send,
        headers={"Authorization": f"Bearer {WHATSAPP_TOKEN}", "Content-Type": "application/json"},
        json=payload,
        timeout=30,
    )
    if resp.status_code != 200:
        err = f"Send document failed: {resp.status_code} {resp.text}"
        print(f"  {err}")
        return False, err
    return True, ""


# ---------------- PARSING (same as whatsapp_label_bot.py) ----------------

def convert_words(price):
    words = num2words(price)
    words = words.replace(",", "")
    words = words.replace("-", " ")
    words = words.title()
    return words


def parse_order(text):
    text = text.strip()
    name_m = re.search(rf"Name{DELIM}+(.+)", text, re.I)
    pincode_m = re.search(rf"Pincode{DELIM}+\s*(\d{{6}})", text, re.I)
    state_m = re.search(rf"State{DELIM}+(.+)", text, re.I)
    phone_m = re.search(rf"Phone\s*(?:number)?{DELIM}+\s*(\d{{10}})", text, re.I)

    missing = []
    if not name_m:    missing.append("Name")
    if not pincode_m: missing.append("Pincode")
    if not state_m:   missing.append("State")
    if not phone_m:   missing.append("Phone")
    if missing:
        raise ValueError(f"Missing fields: {', '.join(missing)}")

    name = name_m.group(1).strip()
    pincode = pincode_m.group(1).strip()
    state = state_m.group(1).strip()
    phone = phone_m.group(1).strip()

    addr_m = re.search(rf"Address{DELIM}+(.+?)(?=\s*(?:City|Pincode))", text, re.S | re.I)
    address = addr_m.group(1).strip() if addr_m else ""
    address = re.sub(r"\s*\n\s*", ", ", address)
    address = address.strip(", ")
    address = re.sub(r",\s*,", ",", address)

    lines = text.strip().split("\n")
    price = None
    price_line_idx = None
    for idx in range(len(lines) - 1, -1, -1):
        stripped = lines[idx].strip()
        if re.match(r"^\d{2,6}$", stripped):
            price = int(stripped)
            price_line_idx = idx
            break

    if price is not None:
        phone_line_idx = None
        for idx, ln in enumerate(lines):
            if re.search(r"Phone", ln, re.I):
                phone_line_idx = idx
                break
        start = (phone_line_idx + 1) if phone_line_idx is not None else (price_line_idx - 1)
        while start < price_line_idx and not lines[start].strip():
            start += 1
        item_text = " ".join(lines[start:price_line_idx]).strip()
    else:
        item_price_m = re.search(r"(\d+\s*[A-Za-z].*?)[.;:]?\s+(\d{2,6})\s*$", text)
        if item_price_m:
            item_text = item_price_m.group(1).strip()
            price = int(item_price_m.group(2))
        else:
            raise ValueError("Could not parse item/price")

    item = parse_item_text(item_text)
    return {
        "name": name, "address": address, "state": state,
        "pincode": pincode, "phone": phone, "price": price, "item": item,
    }


def parse_item_text(item_text):
    cxe_m = re.match(r"(\d+)\s*([Cc][Xx][Ee])[.;:,]?\s*(.*)", item_text)
    if not cxe_m:
        code_m = re.match(r"(\d+)\s*(.*)", item_text)
        if code_m:
            qty = code_m.group(1)
            code = code_m.group(2).strip().upper().rstrip(".,;:")
            return f"{qty} {code}" if code else f"{qty} ITEM"
        return item_text.upper()

    cxe_qty = cxe_m.group(1)
    cxe_code = cxe_m.group(2).upper()
    rest = cxe_m.group(3).strip()
    items = [f"{cxe_qty} {cxe_code}"]
    if rest:
        parts = re.findall(r"(\d+\s*[A-Za-z][A-Za-z ]*)", rest)
        for p in parts:
            m = re.match(r"(\d+)\s*(.*)", p.strip())
            if m:
                items.append(f"{m.group(1)} {m.group(2).strip().title()}")
    return ", ".join(items)


def split_orders(text):
    parts = re.split(rf"(?=Name{DELIM})", text, flags=re.I)
    return [p.strip() for p in parts if p.strip() and re.search(rf"Name{DELIM}", p, re.I)]


# ---------------- LABEL GENERATION (same logic) ----------------

from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def fill_table(table, data):
    price_words = convert_words(data["price"])
    placeholders = {
        "{{NAME}}": data["name"], "{{ADDRESS}}": data["address"],
        "{{STATE}}": data["state"], "{{PINCODE}}": data["pincode"],
        "{{PHONE}}": data["phone"], "{{PRICE}}": str(data["price"]),
        "{{PRICE_WORDS}}": price_words, "{{ITEM}}": data["item"],
    }
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                text = p.text
                replaced = text
                for key, value in placeholders.items():
                    replaced = replaced.replace(key, value)
                if replaced != text:
                    if p.runs:
                        p.runs[0].text = replaced
                        for i in range(1, len(p.runs)):
                            p.runs[i].text = ""
                    else:
                        p.text = replaced


def make_gap_element():
    gap = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "200")
    spacing.set(qn("w:after"), "200")
    spacing.set(qn("w:line"), "240")
    spacing.set(qn("w:lineRule"), "exact")
    pPr.append(spacing)
    rPr = OxmlElement("w:rPr")
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "12")
    rPr.append(sz)
    pPr.append(rPr)
    gap.append(pPr)
    return gap


def remove_empty_body_paragraphs(doc):
    body = doc.element.body
    for p in list(body.findall(qn("w:p"))):
        if p.getparent() != body:
            continue
        has_text = any(t.text for t in p.findall(f".//{qn('w:t')}") if t.text and t.text.strip())
        has_break = p.findall(f".//{qn('w:br')}")
        if not has_text and not has_break:
            body.remove(p)


def add_label(data):
    template_doc = Document(TEMPLATE)
    template_table = template_doc.tables[0]

    if os.path.exists(OUTPUT):
        doc = Document(OUTPUT)
    else:
        doc = Document(TEMPLATE)
        remove_empty_body_paragraphs(doc)
        fill_table(doc.tables[0], data)
        doc.save(OUTPUT)
        return

    tables_count = len(doc.tables)
    last_table_elem = doc.tables[-1]._element

    if tables_count % 2 == 0:
        page_break_p = OxmlElement("w:p")
        page_break_r = OxmlElement("w:r")
        page_break_br = OxmlElement("w:br")
        page_break_br.set(qn("w:type"), "page")
        page_break_r.append(page_break_br)
        page_break_p.append(page_break_r)
        last_table_elem.addnext(page_break_p)
        anchor = page_break_p
    else:
        anchor = last_table_elem

    gap = make_gap_element()
    anchor.addnext(gap)
    new_table = deepcopy(template_table._element)
    gap.addnext(new_table)
    fill_table(doc.tables[-1], data)
    doc.save(OUTPUT)


def stop_and_export():
    """Returns (pdf_path_or_None, docx_path)."""
    if not os.path.exists(OUTPUT):
        return None, None
    batch_num = get_next_batch_number()
    cod_docx = f"cod{batch_num}.docx"
    cod_pdf = f"cod{batch_num}.pdf"
    pdf_path = convert_to_pdf(OUTPUT)
    os.rename(OUTPUT, cod_docx)
    final_pdf = None
    if pdf_path and os.path.exists(pdf_path):
        os.rename(pdf_path, cod_pdf)
        final_pdf = os.path.abspath(cod_pdf)
    if os.path.exists(PROCESSED_FILE):
        os.remove(PROCESSED_FILE)
    return final_pdf, os.path.abspath(cod_docx)


def convert_to_pdf(docx_path):
    pdf_path = docx_path.replace(".docx", ".pdf")
    # Try docx2pdf (needs MS Word on Windows)
    if docx2pdf_convert:
        try:
            docx2pdf_convert(docx_path, pdf_path)
            if os.path.exists(pdf_path):
                return pdf_path
        except Exception as e:
            print(f"  docx2pdf failed (needs MS Word): {e}")
    # Try LibreOffice (works on Linux servers / Windows if installed)
    try:
        outdir = os.path.dirname(os.path.abspath(docx_path)) or "."
        result = subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", outdir, docx_path],
            capture_output=True, text=True, timeout=120
        )
        if result.returncode == 0 and os.path.exists(pdf_path):
            return pdf_path
        print(f"  LibreOffice conversion failed: {result.stderr}")
    except FileNotFoundError:
        print("  LibreOffice not found on this machine")
    except Exception as e:
        print(f"  LibreOffice error: {e}")
    print("  PDF conversion unavailable — will send DOCX instead")
    return None


# ---------------- WEBHOOK ----------------

@app.route("/webhook", methods=["GET"])
def verify():
    """Meta sends a GET request to verify the webhook."""
    mode = request.args.get("hub.mode")
    token = request.args.get("hub.verify_token")
    challenge = request.args.get("hub.challenge")

    if mode == "subscribe" and token == VERIFY_TOKEN:
        print("Webhook verified.")
        return challenge, 200
    return "Forbidden", 403


@app.route("/webhook", methods=["POST"])
def webhook():
    """Receive incoming WhatsApp messages."""
    # Verify signature if APP_SECRET is set
    if APP_SECRET:
        signature = request.headers.get("X-Hub-Signature-256", "")
        expected = "sha256=" + hmac.new(
            APP_SECRET.encode(), request.data, hashlib.sha256
        ).hexdigest()
        if not hmac.compare_digest(signature, expected):
            return "Invalid signature", 403

    body = request.get_json()

    if not body:
        return "OK", 200

    # Extract messages
    try:
        entry = body.get("entry", [{}])[0]
        changes = entry.get("changes", [{}])[0]
        value = changes.get("value", {})
        messages = value.get("messages", [])
    except (IndexError, KeyError):
        return "OK", 200

    for msg in messages:
        if msg.get("type") != "text":
            continue

        sender = msg["from"]  # e.g. "919342901848"
        text = msg["text"]["body"].strip()

        # Check allowed numbers
        if ALLOWED_NUMBERS and sender not in ALLOWED_NUMBERS:
            continue

        print(f"\n[MSG from {sender}]: {text[:100]}")
        handle_message(sender, text)

    return "OK", 200


def handle_message(sender, text):
    """Process a single incoming message."""
    state = load_state()
    processed = load_processed()
    lower = text.strip().lower()

    # --- START command ---
    if lower == "start":
        state["collecting"] = True
        state["batch_count"] = 0
        # Clear previous batch data
        save_processed(set())
        if os.path.exists(OUTPUT):
            os.remove(OUTPUT)
        save_state(state)
        send_message(sender, "Started collecting orders.\nPaste order details now.\nSend 'stop' when done to get the PDF.")
        print("  STARTED collecting")
        return

    # --- STOP command ---
    if lower == "stop":
        if not state.get("collecting"):
            send_message(sender, "Not currently collecting. Send 'start' first.")
            return

        count = state.get("batch_count", 0)
        state["collecting"] = False
        state["batch_count"] = 0
        save_state(state)

        if count == 0:
            send_message(sender, "No orders were collected. Nothing to export.")
            return

        send_message(sender, f"Stopped. {count} label(s) collected.\nGenerating PDF...")

        pdf_path, docx_path = stop_and_export()
        if pdf_path:
            batch_num = int(open(BATCH_COUNTER_FILE).read().strip())
            success, send_err = send_document(sender, pdf_path, caption=f"COD Labels — Batch {batch_num}")
            if success:
                send_message(sender, f"PDF sent! (cod{batch_num}.pdf)\nSend 'start' for the next batch.")
            else:
                send_message(sender, f"PDF saved locally but could not send.\nReason: {send_err[:700]}")
        elif docx_path:
            batch_num = int(open(BATCH_COUNTER_FILE).read().strip())
            success, send_err = send_document(sender, docx_path, caption=f"COD Labels — Batch {batch_num}")
            if success:
                send_message(sender, f"DOCX file sent! (cod{batch_num}.docx)\nInstall MS Word or LibreOffice for PDF.\nSend 'start' for the next batch.")
            else:
                send_message(sender, f"Could not send file.\nReason: {send_err[:700]}")
        else:
            send_message(sender, "No labels found to export.")

        print(f"  STOPPED — {count} labels exported")
        return

    # --- STATUS command ---
    if lower == "status":
        collecting = state.get("collecting", False)
        count = state.get("batch_count", 0)
        if collecting:
            send_message(sender, f"Collecting orders: {count} label(s) so far.\nSend 'stop' to export PDF.")
        else:
            send_message(sender, "Not collecting. Send 'start' to begin.")
        return

    # --- Order data ---
    if not state.get("collecting"):
        send_message(sender, "Send 'start' first to begin collecting orders.")
        return

    # Try to parse orders from the message
    if not re.search(rf"Name{DELIM}", text, re.I):
        return  # Not an order, ignore silently

    order_blocks = split_orders(text)
    added = 0

    for block in order_blocks:
        try:
            data = parse_order(block)
            h = order_hash(data)
            if h in processed:
                continue
            add_label(data)
            processed.add(h)
            state["batch_count"] = state.get("batch_count", 0) + 1
            added += 1
            print(f"  + Label for: {data['name']} (₹{data['price']})")
        except ValueError as e:
            send_message(sender, f"Could not parse order: {e}")

    save_processed(processed)
    save_state(state)

    if added and SEND_PROGRESS_REPLY:
        total = state["batch_count"]
        send_message(sender, f"Added {added} label(s). Total this batch: {total}\nSend more orders or 'stop' to export.")


# ---------------- MAIN ----------------

if __name__ == "__main__":
    if not WHATSAPP_TOKEN or not PHONE_NUMBER_ID:
        print("=" * 55)
        print("  WhatsApp Cloud API Bot — COD Label Generator")
        print("=" * 55)
        print()
        print("ERROR: Please fill in your API credentials in the")
        print("config section at the top of this file:")
        print("  WHATSAPP_TOKEN   = your access token")
        print("  PHONE_NUMBER_ID  = your phone number ID")
        print()
        print("See setup instructions at the top of this file.")
        sys.exit(1)

    import sys
    print("=" * 55)
    print("  WhatsApp Cloud API Bot — COD Label Generator")
    print("=" * 55)
    print()
    print("Commands (send via WhatsApp):")
    print("  start    — Begin collecting orders")
    print("  stop     — Export PDF & send it back")
    print("  status   — Check current batch count")
    print()
    print("Starting webhook server on port 5000...")
    print("If ngrok shows ERR_NGROK_4018, run once:")
    print("  .\\ngrok.exe config add-authtoken <YOUR_NGROK_TOKEN>")
    print("Then start tunnel:")
    print("  .\\ngrok.exe http 5000")
    print()

    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
