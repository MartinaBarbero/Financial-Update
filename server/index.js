const express = require("express");
const cors = require("cors");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const Anthropic = require("@anthropic-ai/sdk");
const XLSX = require("xlsx");

const app = express();
const PORT = process.env.PORT || 10000;
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });

// ─── MIDDLEWARE ───────────────────────────────────────────────────────────────
app.use(cors());
app.use(express.json({ limit: "10mb" }));

// Serve built React app
const DIST = path.join(__dirname, "../client/dist");
if (fs.existsSync(DIST)) {
  app.use(express.static(DIST));
}

// ─── ANTHROPIC CLIENT ─────────────────────────────────────────────────────────
const getClient = () => {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) throw new Error("ANTHROPIC_API_KEY environment variable not set");
  return new Anthropic({ apiKey });
};

// ─── HELPER: build content blocks from uploaded files ─────────────────────────
const buildContentBlocks = (files) => {
  const blocks = [];
  for (const file of files || []) {
    const mime = file.mimetype || "application/octet-stream";
    const b64 = file.buffer.toString("base64");

    if (mime === "application/pdf") {
      blocks.push({ type: "document", source: { type: "base64", media_type: "application/pdf", data: b64 } });
    } else if (mime.startsWith("image/")) {
      blocks.push({ type: "image", source: { type: "base64", media_type: mime, data: b64 } });
    } else if (file.originalname.match(/\.xlsx?$/i)) {
      // Parse Excel → CSV text
      try {
        const wb = XLSX.read(file.buffer, { type: "buffer" });
        let out = "";
        wb.SheetNames.forEach(sh => {
          out += `\n--- Sheet: ${sh} ---\n`;
          out += XLSX.utils.sheet_to_csv(wb.Sheets[sh]);
        });
        blocks.push({ type: "text", text: `File: ${file.originalname}\n${out.slice(0, 20000)}` });
      } catch (e) {
        blocks.push({ type: "text", text: `File: ${file.originalname}\n[Could not parse Excel: ${e.message}]` });
      }
    } else if (file.originalname.match(/\.docx?$/i)) {
      // Send raw text content (best effort)
      const text = file.buffer.toString("utf8", 0, 15000).replace(/[^\x20-\x7E\n\r\t]/g, " ");
      blocks.push({ type: "text", text: `File: ${file.originalname}\n${text}` });
    } else {
      // CSV, TXT, etc.
      const text = file.buffer.toString("utf8", 0, 15000);
      blocks.push({ type: "text", text: `File: ${file.originalname}\n${text}` });
    }
  }
  return blocks;
};

// ─── ROUTE 1: DOCUMENT EXTRACTION ────────────────────────────────────────────
// POST /api/extract  — multipart with files[] + context (string)
app.post("/api/extract", upload.array("files", 10), async (req, res) => {
  try {
    const client = getClient();
    const context = req.body.context || "deep-tech medtech startup";
    const contentBlocks = buildContentBlocks(req.files);

    if (contentBlocks.length === 0) {
      return res.json({ error: "No files received" });
    }

    // Add the extraction prompt
    contentBlocks.push({
      type: "text",
      text: `You are a senior financial analyst. Extract ALL financial data from the documents above.
Company context: ${context}

CRITICAL: Extract every number you can find. Founders rarely include WACC inputs directly — focus on operational data: revenues, costs, headcount, burn rate, margins, market size, milestones, deal terms.

Return ONLY valid JSON — no markdown fences, no explanation. Use null for missing values. Monetary values as plain numbers. Year arrays = 11 values for 2025–2035. Map historical actuals to early years; project forward using stated growth rates.

{
  "currency": "CHF",
  "company_stage": null,
  "funding_raised_total": null,
  "last_round_size": null,
  "pre_money_valuation": null,
  "revenue_year1": null,
  "revenue_year2": null,
  "revenue_year3": null,
  "revenue_year5": null,
  "revenue_growth_rate": null,
  "gross_margin": null,
  "ebitda_margin": null,
  "net_margin": null,
  "burn_rate_monthly": null,
  "runway_months": null,
  "cash_on_hand": null,
  "headcount_current": null,
  "headcount_projected_3y": null,
  "avg_salary": null,
  "tam_size": null,
  "sam_size": null,
  "market_growth_rate": null,
  "deal1_signing_year": null,
  "deal1_upfront_fee": null,
  "deal1_royalty_rate": null,
  "deal1_partner_revenue": [null,null,null,null,null,null,null,null,null,null,null],
  "deal2_codev_fee": [null,null,null,null,null,null,null,null,null,null,null],
  "deal2_royalty_rate": null,
  "deal2_partner_revenue": [null,null,null,null,null,null,null,null,null,null,null],
  "deal3_revenue": [null,null,null,null,null,null,null,null,null,null,null],
  "ftes": [null,null,null,null,null,null,null,null,null,null,null],
  "avg_fte_cost": null,
  "lab_materials": [null,null,null,null,null,null,null,null,null,null,null],
  "ip_patent": [null,null,null,null,null,null,null,null,null,null,null],
  "regulatory_clinical": [null,null,null,null,null,null,null,null,null,null,null],
  "ga_facilities": [null,null,null,null,null,null,null,null,null,null,null],
  "cpi_escalator": null,
  "capex": [null,null,null,null,null,null,null,null,null,null,null],
  "da_rate": null,
  "nwc_pct_rev": null,
  "exit_ev_ebitda": null,
  "wacc_rf": null,
  "wacc_erp": null,
  "tax_rate": null,
  "cost_of_debt": null,
  "beta_unlevered": null,
  "debt_equity_ratio": null,
  "terminal_growth": null,
  "milestone_fda_year": null,
  "milestone_commercial_year": null,
  "milestone_partnership_year": null,
  "rd_spend_annual": null,
  "regulatory_approval_cost": null,
  "clinical_trial_cost": null,
  "ip_portfolio_size": null,
  "patents_filed": null,
  "notes": "List every specific number found and its source section. Be exhaustive."
}`
    });

    const response = await client.messages.create({
      model: "claude-sonnet-4-5-20251001",
      max_tokens: 3000,
      messages: [{ role: "user", content: contentBlocks }]
    });

    const raw = (response.content?.[0]?.text || "")
      .replace(/```[a-z]*\n?/g, "").replace(/```/g, "").trim();

    try {
      const parsed = JSON.parse(raw);
      res.json(parsed);
    } catch {
      res.json({ notes: "Could not parse extraction response", raw });
    }
  } catch (err) {
    console.error("/api/extract error:", err.message);
    res.status(500).json({ error: err.message });
  }
});

// ─── ROUTE 2: PDF REPORT NARRATIVE ───────────────────────────────────────────
// POST /api/report  — JSON body with { prompt }
app.post("/api/report", async (req, res) => {
  try {
    const client = getClient();
    const { prompt } = req.body;
    if (!prompt) return res.status(400).json({ error: "Missing prompt" });

    const response = await client.messages.create({
      model: "claude-sonnet-4-5-20251001",
      max_tokens: 4000,
      messages: [{ role: "user", content: prompt }]
    });

    res.json({ text: response.content?.[0]?.text || "" });
  } catch (err) {
    console.error("/api/report error:", err.message);
    res.status(500).json({ error: err.message });
  }
});

// ─── ROUTE 3: ADVISOR CHAT ────────────────────────────────────────────────────
// POST /api/chat  — JSON body with { system, messages }
app.post("/api/chat", async (req, res) => {
  try {
    const client = getClient();
    const { system, messages } = req.body;
    if (!messages) return res.status(400).json({ error: "Missing messages" });

    const response = await client.messages.create({
      model: "claude-sonnet-4-5-20251001",
      max_tokens: 600,
      system: system || "",
      messages
    });

    res.json({ text: response.content?.[0]?.text || "" });
  } catch (err) {
    console.error("/api/chat error:", err.message);
    res.status(500).json({ error: err.message });
  }
});

// ─── HEALTH CHECK ─────────────────────────────────────────────────────────────
app.get("/api/health", (req, res) => {
  res.json({
    status: "ok",
    apiKeySet: !!process.env.ANTHROPIC_API_KEY,
    timestamp: new Date().toISOString()
  });
});

// ─── SPA FALLBACK ─────────────────────────────────────────────────────────────
app.get("*", (req, res) => {
  const index = path.join(DIST, "index.html");
  if (fs.existsSync(index)) {
    res.sendFile(index);
  } else {
    res.send(`
      <h2>Valuation Tool — Build required</h2>
      <p>Run <code>cd client && npm install && npm run build</code></p>
      <p>API health: <a href="/api/health">/api/health</a></p>
    `);
  }
});

app.listen(PORT, () => {
  console.log(`✅ Valuation Tool server running on port ${PORT}`);
  console.log(`   API key: ${process.env.ANTHROPIC_API_KEY ? "✅ set" : "❌ NOT SET — AI features disabled"}`);
});
