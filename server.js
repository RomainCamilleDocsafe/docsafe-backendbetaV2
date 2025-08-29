/**
 * DocSafe Backend — Beta V2
 * Node.js / Express (compatible Render)
 *
 * Endpoints:
 *  - POST /clean        -> retourne le fichier nettoyé (PDF/DOCX) avec suffixe *_cleaned
 *  - POST /clean-v2     -> nettoyé + rapport LanguageTool en ZIP (cleaned + report.json + report.html)
 *  - GET  /health       -> status OK
 *
 * Variables d'env (Render -> Settings -> Environment):
 *  - PORT                               (Render la définit)
 *  - ALLOWED_ORIGINS                    ex: https://ton-site.vercel.app,https://autre.com
 *  - LT_API_URL   (optionnel)           ex: https://api.languagetool.org/v2/check  (défaut)
 *  - LT_API_KEY   (optionnel)           si vous utilisez une offre payante / proxy
 */

import express from "express";
import cors from "cors";
import multer from "multer";
import fs from "fs";
import fsp from "fs/promises";
import os from "os";
import path from "path";
import axios from "axios";
import JSZip from "jszip";
import { PDFDocument } from "pdf-lib";
import pdfParse from "pdf-parse";

const app = express();

// ---------- CORS ----------
const allowed = (process.env.ALLOWED_ORIGINS || "")
  .split(",")
  .map(s => s.trim())
  .filter(Boolean);

app.use(
  cors({
    origin: function (origin, cb) {
      if (!origin) return cb(null, true); // outils CLI / tests
      if (allowed.length === 0 || allowed.includes(origin)) return cb(null, true);
      return cb(new Error("Origin not allowed by CORS"));
    },
    credentials: true
  })
);

app.use(express.json({ limit: "25mb" }));

// ---------- Upload ----------
const upload = multer({
  storage: multer.diskStorage({
    destination: (req, file, cb) => cb(null, os.tmpdir()),
    filename: (req, file, cb) => cb(null, Date.now() + "-" + file.originalname.replace(/\s+/g, "_"))
  }),
  limits: { fileSize: 20 * 1024 * 1024 }, // 20 Mo
});

// ---------- Utilitaires ----------
const LT_URL = process.env.LT_API_URL || "https://api.languagetool.org/v2/check";
const LT_KEY = process.env.LT_API_KEY || null;

// Nettoyage linguistique basique (espaces/ponctuation simple)
function cleanTextBasic(str) {
  if (!str) return str;
  return String(str)
    .replace(/\u200B/g, "")               // zero-width space
    .replace(/[ \t]+/g, " ")              // espaces multiples -> un espace
    .replace(/ *\n */g, "\n")             // espaces autour des \n
    .replace(/ ?([,;:!?]) ?/g, "$1 ")     // espace après , ; : ! ?
    .replace(/ \./g, ".")                 // pas d'espace avant .
    .replace(/[ ]{2,}/g, " ")             // multiples -> un
    .trim();
}

// Nettoyage DOCX: suppr. métadonnées + nettoyage texte dans <w:t>
async function processDOCXBasic(inputBuffer) {
  const zip = await JSZip.loadAsync(inputBuffer);

  // 1) Nettoyage des métadonnées
  const corePath = "docProps/core.xml";
  if (zip.file(corePath)) {
    let coreXml = await zip.file(corePath).async("string");
    coreXml = coreXml
      .replace(/<dc:creator>.*?<\/dc:creator>/s, "<dc:creator></dc:creator>")
      .replace(/<cp:lastModifiedBy>.*?<\/cp:lastModifiedBy>/s, "<cp:lastModifiedBy></cp:lastModifiedBy>")
      .replace(/<dc:title>.*?<\/dc:title>/s, "<dc:title></dc:title>")
      .replace(/<dc:subject>.*?<\/dc:subject>/s, "<dc:subject></dc:subject>")
      .replace(/<cp:keywords>.*?<\/cp:keywords>/s, "<cp:keywords></cp:keywords>");
    zip.file(corePath, coreXml);
  }
  if (zip.file("docProps/custom.xml")) {
    zip.remove("docProps/custom.xml"); // souvent des métadonnées perso
  }

  // 2) Nettoyage simple du texte dans word/document.xml
  const docPath = "word/document.xml";
  if (zip.file(docPath)) {
    let xml = await zip.file(docPath).async("string");
    xml = xml.replace(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g, (m, inner) => {
      const cleaned = cleanTextBasic(inner);
      return m.replace(inner, cleaned);
    });
    zip.file(docPath, xml);
  }

  const outBuffer = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  return outBuffer;
}

// Extraction de texte DOCX (pour LT)
async function extractTextFromDOCX(inputBuffer) {
  const zip = await JSZip.loadAsync(inputBuffer);
  const docPath = "word/document.xml";
  if (!zip.file(docPath)) return "";
  const xml = await zip.file(docPath).async("string");
  let text = "";
  xml.replace(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g, (_m, inner) => {
    text += inner + " ";
    return _m;
  });
  return cleanTextBasic(text);
}

// Nettoyage métadonnées PDF (pdf-lib)
async function processPDFBasic(inputBuffer) {
  const pdfDoc = await PDFDocument.load(inputBuffer, { updateMetadata: true });
  pdfDoc.setTitle("");
  pdfDoc.setAuthor("");
  pdfDoc.setSubject("");
  pdfDoc.setKeywords([]);
  pdfDoc.setProducer("");
  pdfDoc.setCreator("");
  const epoch = new Date(0);
  pdfDoc.setCreationDate(epoch);
  pdfDoc.setModificationDate(epoch);
  const out = await pdfDoc.save();
  return Buffer.from(out);
}

// Extraction de texte PDF (pour LT)
async function extractTextFromPDF(inputBuffer) {
  try {
    const data = await pdfParse(inputBuffer);
    return cleanTextBasic(data.text || "");
  } catch {
    return ""; // si extraction échoue on renvoie vide
  }
}

// Appel LanguageTool avec découpage en chunks (~20k chars)
async function runLanguageTool(fullText, lang = "auto") {
  const limit = 20000;
  const chunks = [];
  for (let i = 0; i < fullText.length; i += limit) {
    chunks.push(fullText.slice(i, i + limit));
  }
  let allMatches = [];
  for (const chunk of chunks) {
    if (!chunk.trim()) continue;
    const payload = new URLSearchParams();
    payload.set("text", chunk);
    payload.set("language", lang || "auto");
    if (LT_KEY) payload.set("apiKey", LT_KEY);
    // vous pouvez activer d'autres options LT ici (enabledRules, etc.)

    const resp = await axios.post(LT_URL, payload.toString(), {
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      timeout: 30_000,
    });
    if (resp?.data?.matches?.length) {
      // repositionnement offset non strict (pas nécessaire pour notre rapport)
      allMatches = allMatches.concat(resp.data.matches);
    }
  }
  return allMatches;
}

function buildReportJSON({ fileName, language, matches, textLength }) {
  const summary = {
    fileName,
    language,
    textLength,
    totalIssues: matches.length,
    byRule: {},
  };
  for (const m of matches) {
    const key = m.rule?.id || "GENERIC";
    summary.byRule[key] = (summary.byRule[key] || 0) + 1;
  }
  return { summary, matches };
}

function buildReportHTML(reportJSON) {
  const { summary, matches } = reportJSON;
  const rows = matches
    .map((m, idx) => {
      const repl = (m.replacements || []).slice(0, 3).map(r => r.value).join(", ");
      const context = m.context?.text || "";
      return `<tr>
        <td style="padding:8px;border-bottom:1px solid #e5e7eb;">${idx + 1}</td>
        <td style="padding:8px;border-bottom:1px solid #e5e7eb;">${m.rule?.id || ""}</td>
        <td style="padding:8px;border-bottom:1px solid #e5e7eb;">${(m.message || "").replace(/</g,"&lt;")}</td>
        <td style="padding:8px;border-bottom:1px solid #e5e7eb;">${(repl || "-").replace(/</g,"&lt;")}</td>
        <td style="padding:8px;border-bottom:1px solid #e5e7eb;">${(context || "").replace(/</g,"&lt;")}</td>
      </tr>`;
    })
    .join("");

  return `<!doctype html>
<html lang="fr">
<head>
<meta charset="utf-8"/>
<title>DocSafe — Rapport LanguageTool</title>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<style>
body{font-family:ui-sans-serif,system-ui,-apple-system,Segoe UI,Roboto,Ubuntu,Cantarell,Noto Sans,sans-serif;background:#0f172a;color:#e5e7eb;margin:0;padding:24px;}
.card{background:#0b1220;border:1px solid #1f2937;border-radius:16px;max-width:1000px;margin:0 auto;box-shadow:0 10px 25px rgba(0,0,0,.35);}
.card h1{font-size:22px;margin:0;padding:16px 20px;border-bottom:1px solid #1f2937;}
.section{padding:16px 20px;}
.kv{display:flex;flex-wrap:wrap;gap:12px;font-size:14px;}
.kv div{background:#111827;border:1px solid #1f2937;padding:10px 12px;border-radius:10px;}
.table{width:100%;border-collapse:collapse;margin-top:16px;font-size:14px;background:#0b1220;}
th{background:#111827;text-align:left;padding:10px;border-bottom:1px solid #1f2937;}
td{vertical-align:top;}
</style>
</head>
<body>
  <div class="card">
    <h1>Rapport LanguageTool</h1>
    <div class="section">
      <div class="kv">
        <div><b>Fichier:</b> ${summary.fileName}</div>
        <div><b>Langue:</b> ${summary.language}</div>
        <div><b>Longueur texte:</b> ${summary.textLength}</div>
        <div><b>Total issues:</b> ${summary.totalIssues}</div>
      </div>
      <table class="table">
        <thead>
          <tr>
            <th>#</th><th>Règle</th><th>Message</th><th>Suggestions</th><th>Contexte</th>
          </tr>
        </thead>
        <tbody>
          ${rows || `<tr><td colspan="5" style="padding:10px;">Aucune suggestion.</td></tr>`}
        </tbody>
      </table>
    </div>
  </div>
</body>
</html>`;
}

// ---------- Routes ----------
app.get("/health", (req, res) => res.json({ ok: true }));

app.post("/clean", upload.single("file"), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: "Missing file" });

  const filePath = req.file.path;
  const originalName = req.file.originalname;
  const lower = originalName.toLowerCase();

  try {
    let outBuffer, outName, mime;

    if (lower.endsWith(".pdf")) {
      const inputBuf = await fsp.readFile(filePath);
      outBuffer = await processPDFBasic(inputBuf);
      outName = originalName.replace(/\.pdf$/i, "") + "_cleaned.pdf";
      mime = "application/pdf";
    } else if (lower.endsWith(".docx")) {
      const inputBuf = await fsp.readFile(filePath);
      outBuffer = await processDOCXBasic(inputBuf);
      outName = originalName.replace(/\.docx$/i, "") + "_cleaned.docx";
      mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    } else {
      return res.status(400).json({ error: "Only PDF or DOCX are supported" });
    }

    res.setHeader("Content-Type", mime);
    res.setHeader("Content-Disposition", `attachment; filename="${encodeURIComponent(outName)}"`);
    return res.send(outBuffer);
  } catch (e) {
    console.error("CLEAN error:", e);
    return res.status(500).json({ error: "Processing failed" });
  } finally {
    // nettoyage du tmp
    fs.existsSync(filePath) && fs.unlink(filePath, () => {});
  }
});

app.post("/clean-v2", upload.single("file"), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: "Missing file" });
  const lang = (req.body?.lt_language || "auto").toString();

  const filePath = req.file.path;
  const originalName = req.file.originalname;
  const lower = originalName.toLowerCase();

  try {
    let cleanedBuffer, cleanedName, textForLT = "";

    if (lower.endsWith(".pdf")) {
      const inputBuf = await fsp.readFile(filePath);
      cleanedBuffer = await processPDFBasic(inputBuf);
      cleanedName = originalName.replace(/\.pdf$/i, "") + "_cleaned.pdf";
      // on extrait le texte du PDF nettoyé
      textForLT = await extractTextFromPDF(cleanedBuffer);
    } else if (lower.endsWith(".docx")) {
      const inputBuf = await fsp.readFile(filePath);
      cleanedBuffer = await processDOCXBasic(inputBuf);
      cleanedName = originalName.replace(/\.docx$/i, "") + "_cleaned.docx";
      textForLT = await extractTextFromDOCX(cleanedBuffer);
    } else {
      return res.status(400).json({ error: "Only PDF or DOCX are supported" });
    }

    // Appel LanguageTool
    const matches = await runLanguageTool(textForLT || "", lang);
    const reportJSON = buildReportJSON({
      fileName: cleanedName,
      language: lang,
      textLength: (textForLT || "").length,
      matches
    });
    const reportHTML = buildReportHTML(reportJSON);

    // ZIP (cleaned + report.json + report.html)
    const zip = new JSZip();
    zip.file(cleanedName, cleanedBuffer);
    zip.file("report.json", JSON.stringify(reportJSON, null, 2));
    zip.file("report.html", reportHTML);

    const zipBuf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
    const zipName = originalName.replace(/\.(pdf|docx)$/i, "") + "_docsafe_report.zip";

    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", `attachment; filename="${encodeURIComponent(zipName)}"`);
    return res.send(zipBuf);
  } catch (e) {
    console.error("CLEAN-V2 error:", e?.response?.data || e);
    return res.status(500).json({ error: "Processing failed (V2)" });
  } finally {
    fs.existsSync(filePath) && fs.unlink(filePath, () => {});
  }
});

// ---------- Start ----------
const PORT = process.env.PORT || 8080;
app.listen(PORT, () => {
  console.log("DocSafe backend running on port", PORT);
});
