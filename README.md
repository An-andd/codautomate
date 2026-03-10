# COD Label Bot — WhatsApp Cloud API

A WhatsApp bot that generates COD (Cash on Delivery) shipping labels from order details sent via WhatsApp. No browser needed — works from your phone anywhere.

## How It Works

1. Send **start** on WhatsApp to begin collecting orders
2. Paste order details (Name, Address, Phone, Pincode, State, Item, Price)
3. Send **stop** to generate and receive the label file (PDF or DOCX)
4. Send **status** to check how many labels are in the current batch

The bot parses order messages, fills a DOCX template with label data (2 labels per page), converts to PDF if possible, and sends the file back to you on WhatsApp.

## Requirements

- Python 3.8+
- WhatsApp Business API account ([Meta Developers](https://developers.facebook.com))
- `cod_template.docx` — the label template with placeholders

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Configure WhatsApp API

1. Go to [Meta Developers](https://developers.facebook.com) → Create App → Business → WhatsApp
2. In **WhatsApp > API Setup**:
   - Note your **Phone Number ID**
   - Generate a **permanent access token**
3. Set environment variables:

```bash
set WHATSAPP_TOKEN=your_access_token
set PHONE_NUMBER_ID=your_phone_number_id
set VERIFY_TOKEN=cod_bot_verify
```

### 3. Run the bot

```bash
python whatsapp_api_bot.py
```

The server starts on port 5000 (or the `PORT` env var).

### 4. Set up webhook

For **local testing**, use ngrok:

```bash
ngrok http 5000
```

Then in Meta Developers → WhatsApp → Configuration:
- **Callback URL**: `https://your-url.ngrok-free.dev/webhook`
- **Verify token**: `cod_bot_verify`
- Subscribe to the **messages** field

## Deploy to Render (Free, 24/7)

1. Push this repo to GitHub
2. Go to [render.com](https://render.com) → **New** → **Web Service**
3. Connect your GitHub repo
4. Configure:
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn whatsapp_api_bot:app`
   - **Instance Type**: Free
5. Add environment variables (`WHATSAPP_TOKEN`, `PHONE_NUMBER_ID`, `VERIFY_TOKEN`)
6. Update Meta webhook URL to your Render URL: `https://your-app.onrender.com/webhook`

## Order Message Format

The bot expects messages like:

```
Name: John Doe
Address: 123 Main Street, Apt 4B
City: Mumbai
Pincode: 400001
State: Maharashtra
Phone: 9876543210
2 CXE
850
```

- Last standalone number = price
- Lines between Phone and price = item description
- Multiple orders in one message are supported

## Files

| File | Purpose |
|------|---------|
| `whatsapp_api_bot.py` | Main bot (WhatsApp Cloud API + Flask webhook) |
| `cod_template.docx` | Label template with placeholders |
| `requirements.txt` | Python dependencies |
| `generate_label.py` | Standalone label generator (for testing) |

## Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `WHATSAPP_TOKEN` | Yes | WhatsApp Cloud API access token |
| `PHONE_NUMBER_ID` | Yes | WhatsApp phone number ID |
| `VERIFY_TOKEN` | No | Webhook verification token (default: `cod_bot_verify`) |
| `APP_SECRET` | No | App secret for request signature verification |
| `PORT` | No | Server port (default: `5000`, set automatically on Render) |
