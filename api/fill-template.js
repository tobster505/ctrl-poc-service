/**
 * CTRL PoC Export Service · fill-template (Clean V1)
 *
 * Purpose:
 * - Accept Botpress V17 PDF payload
 * - Fill the correct PDF template
 * - Return debug JSON with ?debug=1
 * - Otherwise return the completed PDF directly
 * - No Lovable references
 * - No storage
 * - No external callback
 */

export const config = { runtime: "nodejs" };

import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";

/* --------------------------------------------------
 * Small helpers
 * -------------------------------------------------- */
const S = (v, fb = "") => (v == null ? String(fb) : String(v));
const N = (v, fb = 0) => (Number.isFinite(+v) ? +v : fb);
const clean = (v) => S(v).trim();
const okObj = (o) => o && typeof o === "object" && !Array.isArray(o);
const okArr = (a) => Array.isArray(a);

function normEmail(v) {
  return clean(v).toLowerCase();
}

function looksEmail(v) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normEmail(v));
}

function safeJson(obj) {
  try {
    return JSON.parse(JSON.stringify(obj));
  } catch {
    return { _error: "Could not serialise object" };
  }
}

function clampStrForFilename(s) {
  return S(s)
    .trim()
    .replace(/[^\w\s-]/g, "")
    .replace(/\s+/g, "_")
    .replace(/_+/g, "_")
    .slice(0, 80) || "CTRL_Profile";
}

function pick(obj, paths, fb = "") {
  for (const p of paths) {
    const parts = p.split(".");
    let cur = obj;
    let ok = true;
    for (const part of parts) {
      if (!cur || typeof cur !== "object" || !(part in cur)) {
        ok = false;
        break;
      }
      cur = cur[part];
    }
    if (ok && cur != null && String(cur).trim() !== "") return cur;
  }
  return fb;
}

function joinParas(arr) {
  return (arr || [])
    .map((x) => clean(x))
    .filter(Boolean)
    .join("\n\n")
    .trim();
}

/* --------------------------------------------------
 * Template path helpers
 * -------------------------------------------------- */
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

function resolveTemplatePath(templateKey) {
  // Update these filenames to match your real template files in Vercel
  const MAP = {
    C: path.join(__dirname, "../templates/CTRL_C.pdf"),
    T: path.join(__dirname, "../templates/CTRL_T.pdf"),
    R: path.join(__dirname, "../templates/CTRL_R.pdf"),
    L: path.join(__dirname, "../templates/CTRL_L.pdf")
  };

  return MAP[clean(templateKey).toUpperCase()] || MAP.R;
}

/* --------------------------------------------------
 * Text wrapping / drawing
 * -------------------------------------------------- */
function wrapText(text, font, fontSize, maxWidth) {
  const raw = clean(text);
  if (!raw) return [];

  const paragraphs = raw.split(/\n\s*\n/).map((p) => p.trim()).filter(Boolean);
  const lines = [];

  for (const para of paragraphs) {
    const words = para.split(/\s+/).filter(Boolean);
    let line = "";

    for (const word of words) {
      const testLine = line ? `${line} ${word}` : word;
      const width = font.widthOfTextAtSize(testLine, fontSize);

      if (width <= maxWidth) {
        line = testLine;
      } else {
        if (line) lines.push(line);
        line = word;
      }
    }

    if (line) lines.push(line);
    lines.push(""); // paragraph gap
  }

  if (lines[lines.length - 1] === "") lines.pop();
  return lines;
}

function drawTextBlock(page, text, opts) {
  const {
    x,
    y,
    maxWidth,
    maxHeight,
    font,
    fontSize = 11,
    lineHeight = 14,
    color = rgb(0, 0, 0)
  } = opts;

  const lines = wrapText(text, font, fontSize, maxWidth);
  let cursorY = y;
  let drawn = 0;
  const bottomLimit = y - maxHeight;

  for (const line of lines) {
    if (cursorY - lineHeight < bottomLimit) break;

    if (line === "") {
      cursorY -= lineHeight;
      continue;
    }

    page.drawText(line, {
      x,
      y: cursorY,
      size: fontSize,
      font,
      color
    });

    cursorY -= lineHeight;
    drawn++;
  }

  return {
    linesAttempted: lines.length,
    linesDrawn: drawn,
    finalY: cursorY,
    truncated: drawn < lines.filter(Boolean).length
  };
}

/* --------------------------------------------------
 * Payload normalisation
 * -------------------------------------------------- */
function normalisePayload(payload) {
  const assessmentUUID = clean(
    pick(payload, [
      "meta.assessmentUUID",
      "identity.assessmentUUID",
      "ctrl.assessmentUUID"
    ])
  );

  const fullName = clean(
    pick(payload, [
      "identity.fullName",
      "identity.preferredName"
    ])
  );

  const preferredName = clean(
    pick(payload, [
      "identity.preferredName",
      "identity.fullName"
    ])
  );

  const email = clean(pick(payload, ["identity.email"]));
  const dateLabel = clean(pick(payload, ["identity.dateLabel", "dateLbl"]));
  const dominantKey = clean(
    pick(payload, ["ctrl.dominantKey", "ctrl.templateKey"])
  ).toUpperCase();
  const secondKey = clean(pick(payload, ["ctrl.secondKey"])).toUpperCase();
  const templateKey = clean(
    pick(payload, ["ctrl.templateKey", "ctrl.dominantKey"])
  ).toUpperCase();

  const text = okObj(payload.text) ? payload.text : {};

  return {
    assessmentUUID,
    fullName,
    preferredName,
    email,
    dateLabel,
    dominantKey,
    secondKey,
    templateKey: templateKey || dominantKey || "R",
    text: {
      snapshot: clean(text.snapshot),
      chart_overview: clean(text.chart_overview),
      awareness_movement: clean(text.awareness_movement),
      themes: clean(text.themes),
      interactions_with_others: clean(text.interactions_with_others),
      actions_bullets: clean(text.actions_bullets),
      act_1: clean(text.act_1),
      act_2: clean(text.act_2),
      act_3: clean(text.act_3),
      act_4: clean(text.act_4),
      act_5: clean(text.act_5),
      act_6: clean(text.act_6),
      mirror_summary: clean(text.mirror_summary)
    },
    raw: payload
  };
}

function validatePayload(p) {
  const errors = [];

  if (!p.fullName) errors.push("Missing identity.fullName");
  if (!p.email) errors.push("Missing identity.email");
  if (p.email && !looksEmail(p.email)) errors.push("Invalid identity.email");
  if (!p.templateKey) errors.push("Missing ctrl.templateKey / ctrl.dominantKey");
  if (!p.text.snapshot) errors.push("Missing text.snapshot");
  if (!p.text.chart_overview) errors.push("Missing text.chart_overview");

  return errors;
}

/* --------------------------------------------------
 * PDF generation
 * -------------------------------------------------- */
async function buildPdfBuffer(payload) {
  const P = normalisePayload(payload);
  const templatePath = resolveTemplatePath(P.templateKey);

  const templateBytes = await fs.readFile(templatePath);
  const pdfDoc = await PDFDocument.load(templateBytes);

  const fontRegular = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

  const pages = pdfDoc.getPages();
  const page1 = pages[0];

  // -----------------------------------------
  // IMPORTANT:
  // Replace these coordinates with your real ones.
  // This is the clean structure for testing.
  // -----------------------------------------
  const debugBlocks = {};

  debugBlocks.fullName = drawTextBlock(page1, P.fullName, {
    x: 72,
    y: 730,
    maxWidth: 220,
    maxHeight: 30,
    font: fontBold,
    fontSize: 16,
    lineHeight: 18
  });

  debugBlocks.dateLabel = drawTextBlock(page1, P.dateLabel, {
    x: 72,
    y: 708,
    maxWidth: 180,
    maxHeight: 20,
    font: fontRegular,
    fontSize: 10,
    lineHeight: 12
  });

  debugBlocks.snapshot = drawTextBlock(page1, P.text.snapshot, {
    x: 72,
    y: 640,
    maxWidth: 220,
    maxHeight: 140,
    font: fontRegular,
    fontSize: 10.5,
    lineHeight: 13
  });

  debugBlocks.chart_overview = drawTextBlock(page1, P.text.chart_overview, {
    x: 320,
    y: 640,
    maxWidth: 220,
    maxHeight: 140,
    font: fontRegular,
    fontSize: 10.5,
    lineHeight: 13
  });

  if (pages[1]) {
    debugBlocks.awareness_movement = drawTextBlock(pages[1], P.text.awareness_movement, {
      x: 72,
      y: 730,
      maxWidth: 220,
      maxHeight: 180,
      font: fontRegular,
      fontSize: 10.5,
      lineHeight: 13
    });

    debugBlocks.themes = drawTextBlock(pages[1], P.text.themes, {
      x: 320,
      y: 730,
      maxWidth: 220,
      maxHeight: 180,
      font: fontRegular,
      fontSize: 10.5,
      lineHeight: 13
    });

    const actionsText = joinParas([
      P.text.actions_bullets,
      P.text.act_1,
      P.text.act_2,
      P.text.act_3
    ]);

    debugBlocks.actions = drawTextBlock(pages[1], actionsText, {
      x: 72,
      y: 500,
      maxWidth: 468,
      maxHeight: 160,
      font: fontRegular,
      fontSize: 10.5,
      lineHeight: 13
    });
  }

  const pdfBytes = await pdfDoc.save();

  return {
    pdfBytes,
    meta: {
      templateKey: P.templateKey,
      templatePath,
      pages: pages.length,
      fullName: P.fullName,
      preferredName: P.preferredName,
      email: P.email,
      assessmentUUID: P.assessmentUUID,
      textLens: {
        snapshot: P.text.snapshot.length,
        chart_overview: P.text.chart_overview.length,
        awareness_movement: P.text.awareness_movement.length,
        themes: P.text.themes.length,
        interactions_with_others: P.text.interactions_with_others.length
      },
      draw: debugBlocks
    }
  };
}

/* --------------------------------------------------
 * Read input payload
 * -------------------------------------------------- */
async function getPayload(req) {
  if (req.method === "POST") {
    if (okObj(req.body)) return req.body;

    if (typeof req.body === "string" && req.body.trim()) {
      return JSON.parse(req.body);
    }

    return {};
  }

  if (req.method === "GET" && req.query?.data) {
    const raw = Buffer.from(String(req.query.data), "base64").toString("utf8");
    return JSON.parse(raw);
  }

  return {};
}

/* --------------------------------------------------
 * Main handler
 * -------------------------------------------------- */
export default async function handler(req, res) {
  const debugMode = String(req.query?.debug || "") === "1";
  const downloadMode = String(req.query?.download || "") === "1";

  if (!["GET", "POST"].includes(req.method)) {
    return res.status(405).json({
      ok: false,
      error: "Method not allowed",
      allowed: ["GET", "POST"]
    });
  }

  try {
    const payload = await getPayload(req);
    const P = normalisePayload(payload);
    const errors = validatePayload(P);

    if (errors.length) {
      return res.status(400).json({
        ok: false,
        stage: "validate",
        errors,
        received: {
          assessmentUUID: P.assessmentUUID,
          fullName: P.fullName,
          preferredName: P.preferredName,
          email: P.email,
          dateLabel: P.dateLabel,
          dominantKey: P.dominantKey,
          secondKey: P.secondKey,
          templateKey: P.templateKey
        }
      });
    }

    const { pdfBytes, meta } = await buildPdfBuffer(payload);

    const filenameBase = clampStrForFilename(
      `${P.fullName || "CTRL_Profile"}_${P.templateKey}_${P.assessmentUUID || "preview"}`
    );
    const filename = `${filenameBase}.pdf`;

    if (debugMode) {
      return res.status(200).json({
        ok: true,
        mode: "debug",
        filename,
        bytes: pdfBytes.length,
        received: {
          assessmentUUID: P.assessmentUUID,
          fullName: P.fullName,
          preferredName: P.preferredName,
          email: P.email,
          dateLabel: P.dateLabel,
          dominantKey: P.dominantKey,
          secondKey: P.secondKey,
          templateKey: P.templateKey
        },
        pdf: safeJson(meta)
      });
    }

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader(
      "Content-Disposition",
      `${downloadMode ? "attachment" : "inline"}; filename="${filename}"`
    );
    res.setHeader("Content-Length", String(pdfBytes.length));
    res.setHeader("Cache-Control", "no-store");

    return res.status(200).send(Buffer.from(pdfBytes));
  } catch (err) {
    return res.status(500).json({
      ok: false,
      stage: "handler",
      error: clean(err?.message || err),
      stack: process.env.NODE_ENV === "development" ? clean(err?.stack || "") : ""
    });
  }
}
