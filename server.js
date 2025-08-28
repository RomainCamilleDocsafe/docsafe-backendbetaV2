/* DocSafe Backend V2 - BETA
 * Node.js/Express (Render-ready)
 *
 * Endpoints:
 *  - GET  /health
 *  - POST /clean     -> Nettoie fichier (PDF/DOCX)
 *  - POST /clean-v2  -> Nettoie + génère rapport LanguageTool (ZIP)
 */

const express = require("express");
const cors = require("cors");
const multer = require("multer");
const { PDFDocument } = require("pdf-lib");
const JSZip = require("jszip");
const fetch = require("node-fetch");

// pdf-parse optionnel (si installé dans package.json)
let pdfParse = null;
try {
  pdfParse = require("pdf-parse");
} catch (e) {
  console.warn("ℹ️ pdf-parse non installé, extraction texte PDF désactivée.");
}

const app = express();
app.use(express.json());

// ---- CORS ----
const allowed = (process.env.ALLOWED_ORIGINS || "*").split(",");
app.use(
  cors({
    origin: (origin, cb) => {
      if (!origin || allowed.includes("*") || allowed.includes(origin)) return cb(null, true);
      return cb(new Error("Not allowed by CORS"));
    },
  })
);

// ---- Upload config ----
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024, files: 1 }, // 20 Mo max
});

const EXT = (name = "") => (name.split(".").pop() || "").toLowerCase();
const BASENAME = (name = "file") => name.replace(/\.[^/.]+$/, "");
const SAFE = (s = "") => s.replace(/[^\w.\-() ]+/g, "_").replace(/\s+/g, "_").slice(0, 180);

// ---- Nettoyage PDF (supprimer métadonnées) ----
async function cleanPDF(buffer) {
  const src = await PDFDocument.load(buffer);
  const dst = await PDFDocument.create();
  const pages = await dst.copyPages(src, src.getPageIndices());
  pages.forEach((p) => dst.addPage(p));
  dst.setTitle("");
  dst.setAuthor("");
  dst.setSubject("");
  dst.setKeywords([]);
  dst.setProducer("");
  dst.setCreator("");
  dst.setCreationDate(new Date());
  dst.setModificationDate(new Date());
  return await dst.save();
}

// ---- Nettoyage DOCX ----
async function cleanDOCX(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  ["docProps/core.xml", "docProps/app.xml", "docProps/custom.xml"].forEach((f) => {
    if (zip.file(f)) zip.remove(f);
  });
  if (zip.folder("customXml")) zip.remove("customXml");

  // Nettoyage texte léger
  const docFile = zip.file("word/document.xml");
  if (docFile) {
    let xml = await docFile.async("string");
    xml = xml.replace(/\s+/g, " ");
    xml = xml.replace(/ ,/g, ",").replace(/ \./g, ".").replace(/ !/g, "!").replace(/ \?/g, "?");
    zip.file("word/document.xml", xml);
  }

  return await zip.generateAsync({ type: "nodebuffer" });
}

// ---- Extraction texte DOCX ----
async function extractTextFromDOCX(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const doc = zip.file("word/document.xml");
  if (!doc) return "";
  let xml = await doc.async("string");
  xml = xml.replace(/<\/w:p>/g, "\n");
  const texts = [];
  const regex = /<w:t[^>]*>(.*?)<\/w:t>/g;
  let m;
  while ((m = regex.exec(xml)) !== null) {
    let t = m[1]
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&amp;/g, "&")
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'");
    texts.push(t);
  }
  return texts.join(" ");
}

// ---- Extraction texte PDF ----
async function extractTextFromPDF(buffer) {
  if (!pdfParse) return "";
  try {
    const data = await pdfParse(buffer);
    return (data.text || "").trim();
  } catch (e) {
    return "";
  }
}

// ---- LanguageTool ----
const LT_API_URL = process.env.LT_API_URL || "https://api.languagetool.org/v2/check";
const LT_API_KEY = process.env.LT_API_KEY || "";

async function languageToolCheck(text, lang = "auto") {
  if (!text || !text.trim()) return { matches: [] };
  const resp = await fetch(LT_API_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
      ...(LT_API_KEY ? { Authorization: `Bearer ${LT_API_KEY}` } : {}),
    },
    body: new URLSearchParams({
      text,
      language: lang,
    }),
  });
  return await resp.json();
}

// ---- Routes ----
app.get("/health", (req, res) => res.json({ ok: true, message: "Backend is running ✅" }));

// V1 - Nettoyage simple
app.post("/clean", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });
    const ext = EXT(req.file.originalname);
    let cleaned;
    if (ext === "pdf") cleaned = await cleanPDF(req.file.buffer);
    else if (ext === "docx") cleaned = await cleanDOCX(req.file.buffer);
    else return res.status(400).json({ error: "Unsupported format (PDF/DOCX)" });

    const fname = SAFE(BASENAME(req.file.originalname));
    res.setHeader("Content-Disposition", `attachment; filename=\"${fname}_cleaned.${ext}\"`);
    res.setHeader("Content-Type", "application/octet-stream");
    res.send(cleaned);
  } catch (err) {
    console.error("❌ /clean error:", err);
    res.status(500).json({ error: "Processing failed", details: err.message });
  }
});

// V2 - Nettoyage + Rapport
app.post("/clean-v2", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });
    const ext = EXT(req.file.originalname);
    let cleaned;
    let text = "";
    if (ext === "pdf") {
      cleaned = await cleanPDF(req.file.buffer);
      text = await extractTextFromPDF(req.file.buffer);
    } else if (ext === "docx") {
      cleaned = await cleanDOCX(req.file.buffer);
      text = await extractTextFromDOCX(req.file.buffer);
    } else {
      return res.status(400).json({ error: "Unsupported format (PDF/DOCX)" });
    }

    const lang = req.body.lt_language || "auto";
    const ltResult = await languageToolCheck(text, lang);

    const zip = new JSZip();
    const fname = SAFE(BASENAME(req.file.originalname));
    zip.file(`${fname}_cleaned.${ext}`, cleaned);
    zip.file("report.json", JSON.stringify(ltResult, null, 2));
    const html = `<html><body><h2>Rapport LanguageTool</h2><pre>${JSON.stringify(
      ltResult,
      null,
      2
    )}</pre></body></html>`;
    zip.file("report.html", html);

    const outZip = await zip.generateAsync({ type: "nodebuffer" });
    res.setHeader("Content-Disposition", `attachment; filename=\"${fname}_bundle.zip\"`);
    res.setHeader("Content-Type", "application/zip");
    res.send(outZip);
  } catch (err) {
    console.error("❌ /clean-v2 error:", err);
    res.status(500).json({ error: "Processing failed", details: err.message });
  }
});

// ---- Start ----
const PORT = process.env.PORT || 3001;
app.listen(PORT, () => console.log(`✅ DocSafe backend running on port ${PORT}`));
