# Financial Valuation Tool — Render Deployment (Fullstack)

AI-powered DCF valuation for deep-tech medtech startups.  
**Architecture:** Express backend (Node.js) + React frontend — all in one Render Web Service.

---

## 📁 Structure

```
valuation-tool/
├── server/
│   └── index.js          # Express API proxy (extraction, report, chat)
├── client/
│   ├── src/
│   │   ├── index.jsx     # React entry point
│   │   └── App.jsx       # Full UI — Valuation Tool v6
│   ├── public/
│   │   └── index.html
│   ├── package.json
│   └── webpack.config.js
├── package.json          # Root — installs server deps + builds client
├── render.yaml
└── README.md
```

---

## 🚀 Deploy to Render

### Step 1 — Push to GitHub
```bash
git init
git add .
git commit -m "Initial deploy"
git remote add origin https://github.com/YOUR_USERNAME/valuation-tool.git
git push -u origin main
```

### Step 2 — Create Web Service on Render
1. Go to [render.com](https://render.com) → **New → Web Service**
2. Connect your GitHub repo
3. Render reads `render.yaml` automatically — settings pre-filled

### Step 3 — Add API Key (REQUIRED for AI features)
In Render dashboard → your service → **Environment**:
```
ANTHROPIC_API_KEY = sk-ant-xxxxxxxxxxxx
```
> Without the key, the tool still works for WACC + Excel generation.
> AI document extraction, PDF narrative, and chat advisor require the key.

### Step 4 — Deploy
Click **Deploy** — build takes ~3–4 minutes.

---

## 🔧 How it works

```
Browser                    Render Server              Anthropic API
  │                             │                          │
  │── POST /api/extract ───────►│── claude.messages ──────►│
  │   (multipart, files)        │   (server-side key)      │
  │◄── extracted JSON ─────────│◄── JSON response ─────────│
  │                             │
  │── POST /api/report ────────►│── claude.messages ──────►│
  │   (valuation context)       │                          │
  │◄── narrative text ─────────│◄── text response ─────────│
  │                             │
  │── POST /api/chat ──────────►│── claude.messages ──────►│
  │   (conversation)            │                          │
  │◄── reply text ─────────────│◄── text response ─────────│
```

The API key **never touches the browser** — all Anthropic calls go through the Express server.

---

## Supported file types for extraction
| Format | How it's processed |
|---|---|
| PDF | Sent as native document block to Claude |
| Images (PNG, JPG) | Sent as image block — Claude reads tables visually |
| Excel (.xlsx, .xls) | Parsed to CSV text server-side via SheetJS |
| Word (.docx) | Text extracted and sent |
| CSV / TXT | Sent as text |

---

## Local development
```bash
# Install all dependencies
npm install
cd client && npm install && cd ..

# Build client
npm run build

# Start server (serves built client + API)
npm start
# → http://localhost:10000
```
