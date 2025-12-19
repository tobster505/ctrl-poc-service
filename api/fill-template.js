/**
 * CTRL PoC Export Service · fill-template
 * VERSION: V7.1 (V7 + missing visible render pieces restored)
 *
 * Adds:
 * - Cover page: writes Full Name + Date of completion
 * - Header: writes “The CTRL Model PoC Profile for <Full Name>” on pages 2–10
 * - Page 8: renders WorkWith text for Concealed/Triggered/Regulated/Lead
 */
export const config = { runtime: "nodejs" };

import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";
import { PDFDocument, StandardFonts } from "pdf-lib";

/* ───────── base64url decode + safe parse ───────── */
function b64urlToUtf8(b64url) {
  const b64 = String(b64url || "")
    .replace(/-/g, "+")
    .replace(/_/g, "/")
    .padEnd(Math.ceil(String(b64url || "").length / 4) * 4, "=");
  return Buffer.from(b64, "base64").toString("utf8");
}
function safeJsonParse(str) {
  try { return JSON.parse(str); } catch { return null; }
}
function normStr(v) {
  if (v === null || v === undefined) return "";
  return String(v);
}
function pickFirst(obj, keys, fallback = "") {
  for (const k of keys) {
    const v = obj?.[k];
    if (v !== null && v !== undefined && String(v).trim().length) return v;
  }
  return fallback;
}
function getPath(obj, dottedPath) {
  if (!obj || !dottedPath) return undefined;
  const parts = String(dottedPath).split(".");
  let cur = obj;
  for (const p of parts) {
    if (!cur || typeof cur !== "object") return undefined;
    cur = cur[p];
  }
  return cur;
}
function pickFirstPath(obj, paths, fallback = "") {
  for (const p of paths) {
    const v = getPath(obj, p);
    if (v !== null && v !== undefined && String(v).trim().length) return v;
  }
  return fallback;
}

/* ───────── state normalisation ───────── */
const VALID_TEMPLATE_KEYS = new Set([
  "CT","CR","CL",
  "TC","TR","TL",
  "RC","RT","RL",
  "LC","LT","LR",
]);
function toStateLetter(v) {
  const s = normStr(v).trim();
  if (!s) return "";
  const up = s.toUpperCase();
  if (up === "C" || up === "T" || up === "R" || up === "L") return up;

  const map = { CONCEALED:"C", TRIGGERED:"T", REGULATED:"R", LEAD:"L" };
  if (map[up]) return map[up];

  const cleaned = up.replace(/[^A-Z]/g, "");
  if (map[cleaned]) return map[cleaned];

  if (cleaned.includes("CONCEALED")) return "C";
  if (cleaned.includes("TRIGGERED")) return "T";
  if (cleaned.includes("REGULATED")) return "R";
  if (cleaned.includes("LEAD")) return "L";
  return "";
}
function normaliseTemplateKey(v) {
  const s = normStr(v).trim().toUpperCase();
  if (!s) return "";
  const cleaned = s.replace(/[^A-Z]/g, "");
  if (cleaned.length >= 2) {
    const tk = cleaned.slice(0, 2);
    if (VALID_TEMPLATE_KEYS.has(tk)) return tk;
  }
  return "";
}
function buildTemplateKey(domLetter, secondLetter) {
  const d = toStateLetter(domLetter);
  const s = toStateLetter(secondLetter);
  const tk = `${d}${s}`;
  return VALID_TEMPLATE_KEYS.has(tk) ? tk : "";
}
function deriveDomSecondFromTemplateKey(templateKey) {
  const tk = normaliseTemplateKey(templateKey);
  if (!tk) return { domKey: "", secondKey: "" };
  return { domKey: tk[0], secondKey: tk[1] };
}

/* ───────── layout ───────── */
const LAYOUT = {
  // Cover page overlays (page 1)
  p1: {
    name: { x: 370, y: 510, w: 520, size: 18, maxLines: 1, align: "center" },
    date: { x: 370, y: 300, w: 520, size: 16, maxLines: 1, align: "center" },
  },
  // Header line on pages 2–10
  header: {
    x: 60, y: 815, w: 950, size: 14, maxLines: 1, align: "left"
  },
  // Existing V4 blocks
  p3: {
    exec: { x: 55, y: 285, w: 950, size: 18, maxLines: 20, align: "left" },
    execTLDR: { x: 55, y: 215, w: 950, size: 18, maxLines: 10, align: "left" },
    tipAct: { x: 55, y: 730, w: 950, size: 18, maxLines: 10, align: "left" },
  },
  p4: {
    main: { x: 55, y: 250, w: 950, size: 18, maxLines: 28, align: "left" },
    tldr: { x: 55, y: 190, w: 950, size: 18, maxLines: 10, align: "left" },
    act: { x: 55, y: 735, w: 950, size: 18, maxLines: 10, align: "left" },
  },
  p5: {
    main: { x: 55, y: 250, w: 950, size: 18, maxLines: 22, align: "left" },
    tldr: { x: 55, y: 190, w: 950, size: 18, maxLines: 10, align: "left" },
  },
  p6: {
    main: { x: 55, y: 250, w: 950, size: 18, maxLines: 22, align: "left" },
    tldr: { x: 55, y: 190, w: 950, size: 18, maxLines: 10, align: "left" },
    act: { x: 55, y: 735, w: 950, size: 18, maxLines: 10, align: "left" },
  },
  p7: {
    top: { x: 55, y: 250, w: 950, size: 18, maxLines: 25, align: "left" },
    topTLDR: { x: 55, y: 190, w: 950, size: 18, maxLines: 10, align: "left" },
    tip: { x: 55, y: 740, w: 950, size: 18, maxLines: 9, align: "left" },
  },
  // Page 8 work-with boxes (your template shows 4 columns; we overlay into 4 rectangles)
  p8: {
    concealed: { x: 60,  y: 525, w: 225, size: 14, maxLines: 16, align: "left" },
    triggered: { x: 305, y: 525, w: 225, size: 14, maxLines: 16, align: "left" },
    regulated: { x: 550, y: 525, w: 225, size: 14, maxLines: 16, align: "left" },
    lead:      { x: 795, y: 525, w: 225, size: 14, maxLines: 16, align: "left" },
  },
  p9: {
    anchor: { x: 55, y: 240, w: 950, size: 18, maxLines: 14, align: "left" },
  },
};

/* ───────── text drawing ───────── */
function splitToLines(text, maxCharsPerLine) {
  const t = normStr(text).replace(/\r/g, "");
  if (!t.trim()) return [];
  const words = t.split(/\s+/);
  const lines = [];
  let line = "";
  for (const w of words) {
    const cand = line ? `${line} ${w}` : w;
    if (cand.length <= maxCharsPerLine) line = cand;
    else {
      if (line) lines.push(line);
      line = w;
    }
  }
  if (line) lines.push(line);
  return lines;
}
function drawTextBox(page, font, text, box) {
  const { x, y, w, size, maxLines, align } = box;
  const t = normStr(text);
  if (!t.trim()) return;

  const maxChars = Math.max(18, Math.floor(w / (size * 0.55)));
  const lines = splitToLines(t, maxChars).slice(0, maxLines);

  const lineHeight = size * 1.25;
  let cursorY = y;

  for (const line of lines) {
    let tx = x;
    if (align === "center") {
      const tw = font.widthOfTextAtSize(line, size);
      tx = x + (w - tw) / 2;
    } else if (align === "right") {
      const tw = font.widthOfTextAtSize(line, size);
      tx = x + (w - tw);
    }
    page.drawText(line, { x: tx, y: cursorY, size, font });
    cursorY -= lineHeight;
  }
}

/* ───────── payload normaliser (V7) ───────── */
function normaliseInput(payload) {
  const identity =
    payload?.identity ||
    payload?.ctrl?.summary?.identity ||
    payload?.ctrl?.identity ||
    payload?.summary?.identity ||
    {};

  const fullName = pickFirst(identity, ["fullName", "FullName", "name", "Name"], "");
  const email = pickFirst(identity, ["email", "Email"], "");
  const dateLabel = pickFirst(identity, ["dateLabel", "dateLbl", "date", "Date"], payload?.dateLbl || "");

  const ctrl = payload?.ctrl || payload?.CTRL || {};
  const summary = ctrl?.summary || payload?.ctrlSummary || payload?.summary || {};

  let dominantRaw = pickFirstPath(payload, [
    "ctrl.summary.dominantKey","ctrl.summary.domKey","ctrl.summary.dominantState","ctrl.summary.domState",
    "ctrl.dominantKey","ctrl.domKey","summary.dominantKey","summary.domKey",
    "domSecond.domKey","dominantKey","domKey","dominantState","domState",
  ], "");

  let secondRaw = pickFirstPath(payload, [
    "ctrl.summary.secondKey","ctrl.summary.secondState","ctrl.secondKey","summary.secondKey",
    "domSecond.secondKey","secondKey","secondState",
  ], "");

  let templateRaw = pickFirstPath(payload, [
    "ctrl.summary.templateKey","ctrl.templateKey","summary.templateKey","domSecond.templateKey",
    "templateKey","tplKey",
  ], "");

  let dominantKey = toStateLetter(dominantRaw);
  let secondKey = toStateLetter(secondRaw);
  let templateKey = normaliseTemplateKey(templateRaw);

  if (templateKey && (!dominantKey || !secondKey)) {
    const d = deriveDomSecondFromTemplateKey(templateKey);
    dominantKey = dominantKey || d.domKey;
    secondKey = secondKey || d.secondKey;
  }
  if (!templateKey && dominantKey && secondKey) {
    templateKey = buildTemplateKey(dominantKey, secondKey);
  }
  if (!templateKey) templateKey = "CT";

  const bands = ctrl?.bands || payload?.bands || payload?.ctrl12 || payload?.ctrl?.bands || {};
  const text = payload?.text || payload?.gen || payload?.copy || {};
  const workWith = payload?.workWith || payload?.workwith || payload?.work_with || {};
  const questions = payload?.questions || payload?.ctrl?.questions || payload?.ctrl?.summary?.questions || payload?.summary?.questions || [];

  return {
    raw: payload,
    identity: { fullName, email, dateLabel },
    ctrl: { summary: { dominantKey, secondKey, templateKey }, bands },
    text,
    workWith,
    questions,
  };
}

function pickTemplateInfo(P) {
  let tk = normaliseTemplateKey(P?.ctrl?.summary?.templateKey);
  if (!tk) tk = "CT";
  return { templateKey: tk, filename: `CTRL_PoC_Assessment_Profile_template_${tk}.pdf` };
}

/* ───────── handler ───────── */
export default async function handler(req, res) {
  try {
    const __filename = fileURLToPath(import.meta.url);
    const __dirname = path.dirname(__filename);

    const url = new URL(req.url, "https://example.local");
    const dataParam = url.searchParams.get("data");
    if (!dataParam) return res.status(400).json({ ok: false, error: "Missing ?data=" });

    const payload = safeJsonParse(b64urlToUtf8(dataParam));
    if (!payload) return res.status(400).json({ ok: false, error: "Invalid JSON in ?data=" });

    const P = normaliseInput(payload);
    const info = pickTemplateInfo(P);

    const pdfPath = path.join(__dirname, "..", "public", info.filename);
    const pdfBytes = await fs.readFile(pdfPath);

    const pdfDoc = await PDFDocument.load(pdfBytes);
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const pages = pdfDoc.getPages();

    // ───────── MASTER DEBUG ─────────
    const debug = String(url.searchParams.get("debug") || "").trim();
    const wantDebug = debug === "1" || debug === "2";

    const trunc = (v, n = 140) => {
      const s = (v === null || v === undefined) ? "" : String(v);
      return s.length <= n ? s : s.slice(0, n) + "…";
    };
    const strInfo = (v) => {
      if (v === null || v === undefined) return { has: false, len: 0, preview: "" };
      const s = String(v);
      return { has: s.length > 0, len: s.length, preview: trunc(s) };
    };

    const expectedTextKeys = [
      "execSummary_tldr","execSummary","execSummary_tipact",
      "state_tldr","domState","bottomState","state_tipact",
      "frequency_tldr","frequency",
      "sequence_tldr","sequence","sequence_tipact",
      "theme_tldr","theme","theme_tipact",
      "act_anchor",
    ];
    const expectedWorkWithKeys = ["concealed","triggered","regulated","lead"];

    const buildMasterProbe = () => {
      const text = P.text || {};
      const workWith = P.workWith || {};
      const bands = (P.ctrl && P.ctrl.bands) || {};

      const bandsPresent12 =
        ["C_low","C_mid","C_high","T_low","T_mid","T_high","R_low","R_mid","R_high","L_low","L_mid","L_high"]
          .filter((k) => typeof bands?.[k] !== "undefined").length;

      const missing = { identity: [], ctrl: [], text: [], workWith: [] };

      if (!P.identity?.fullName) missing.identity.push("identity.fullName");
      if (!P.identity?.email) missing.identity.push("identity.email");
      if (!P.identity?.dateLabel) missing.identity.push("identity.dateLabel");

      if (!P.ctrl?.summary?.dominantKey) missing.ctrl.push("ctrl.summary.dominantKey");
      if (!P.ctrl?.summary?.secondKey) missing.ctrl.push("ctrl.summary.secondKey");
      if (!P.ctrl?.summary?.templateKey) missing.ctrl.push("ctrl.summary.templateKey");
      if (bandsPresent12 !== 12) missing.ctrl.push("ctrl.bands (12/12)");

      expectedTextKeys.forEach((k) => {
        if (!String(text[k] || "").trim()) missing.text.push(`text.${k}`);
      });
      expectedWorkWithKeys.forEach((k) => {
        if (!String(workWith[k] || "").trim()) missing.workWith.push(`workWith.${k}`);
      });

      const summary = {
        ok: true,
        where: "fill-template:v7.1:master_probe:summary",
        domSecond: {
          domKey: P.ctrl?.summary?.dominantKey || null,
          secondKey: P.ctrl?.summary?.secondKey || null,
          templateKey: P.ctrl?.summary?.templateKey || null,
        },
        identity: {
          fullName: strInfo(P.identity?.fullName),
          email: strInfo(P.identity?.email),
          dateLabel: strInfo(P.identity?.dateLabel),
        },
        counts: {
          questions: Array.isArray(P.questions) ? P.questions.length : 0,
          bandsKeys: Object.keys(bands || {}).length,
          bandsPresent12,
          textKeys: Object.keys(text || {}).length,
          workWithKeys: Object.keys(workWith || {}).length,
        },
        missing,
        previews: {
          execSummary_tldr: trunc(text.execSummary_tldr),
          execSummary: trunc(text.execSummary),
          act_anchor: trunc(text.act_anchor),
          workWith_triggered: trunc(workWith.triggered),
        },
      };

      const full = {
        ok: true,
        where: "fill-template:v7.1:master_probe:full",
        identity: P.identity || null,
        ctrlSummary: P.ctrl?.summary || null,
        bands: bands || null,
        questions: P.questions || null,
        text: text || null,
        workWith: workWith || null,
      };

      return { summary, full };
    };

    const MASTER = buildMasterProbe();
    try { console.log("[fill-template] MASTER_PROBE", JSON.stringify(MASTER.summary)); } catch {}

    if (wantDebug) return res.status(200).json(debug === "2" ? MASTER.full : MASTER.summary);

    // ───────── NEW: Cover page name/date overlays ─────────
    const p1 = pages[0];
    if (p1 && LAYOUT.p1) {
      drawTextBox(p1, font, P.identity.fullName, LAYOUT.p1.name);
      drawTextBox(p1, font, P.identity.dateLabel, LAYOUT.p1.date);
    }

    // ───────── NEW: Header on pages 2–10 ─────────
    const headerText = `The CTRL Model PoC Profile for ${P.identity.fullName || ""}`.trim();
    for (let i = 1; i < pages.length; i++) {
      const pg = pages[i];
      if (pg && LAYOUT.header) drawTextBox(pg, font, headerText, LAYOUT.header);
    }

    // Page 3
    const p3 = pages[2];
    if (p3 && LAYOUT.p3) {
      drawTextBox(p3, font, P.text.execSummary_tldr, LAYOUT.p3.execTLDR);
      drawTextBox(p3, font, P.text.execSummary, LAYOUT.p3.exec);
      drawTextBox(p3, font, P.text.execSummary_tipact, LAYOUT.p3.tipAct);
    }

    // Page 4
    const p4 = pages[3];
    if (p4 && LAYOUT.p4) {
      drawTextBox(p4, font, P.text.state_tldr, LAYOUT.p4.tldr);
      const domAndBottom = `${normStr(P.text.domState)}${P.text.bottomState ? "\n\n" + normStr(P.text.bottomState) : ""}`;
      drawTextBox(p4, font, domAndBottom, LAYOUT.p4.main);
      drawTextBox(p4, font, P.text.state_tipact, LAYOUT.p4.act);
    }

    // Page 5
    const p5 = pages[4];
    if (p5 && LAYOUT.p5) {
      drawTextBox(p5, font, P.text.frequency_tldr, LAYOUT.p5.tldr);
      drawTextBox(p5, font, P.text.frequency, LAYOUT.p5.main);
    }

    // Page 6
    const p6 = pages[5];
    if (p6 && LAYOUT.p6) {
      drawTextBox(p6, font, P.text.sequence_tldr, LAYOUT.p6.tldr);
      drawTextBox(p6, font, P.text.sequence, LAYOUT.p6.main);
      drawTextBox(p6, font, P.text.sequence_tipact, LAYOUT.p6.act);
    }

    // Page 7
    const p7 = pages[6];
    if (p7 && LAYOUT.p7) {
      drawTextBox(p7, font, P.text.theme_tldr, LAYOUT.p7.topTLDR);
      drawTextBox(p7, font, P.text.theme, LAYOUT.p7.top);
      drawTextBox(p7, font, P.text.theme_tipact, LAYOUT.p7.tip);
    }

    // ───────── NEW: Page 8 WorkWith boxes ─────────
    const p8 = pages[7];
    if (p8 && LAYOUT.p8) {
      drawTextBox(p8, font, P.workWith.concealed, LAYOUT.p8.concealed);
      drawTextBox(p8, font, P.workWith.triggered, LAYOUT.p8.triggered);
      drawTextBox(p8, font, P.workWith.regulated, LAYOUT.p8.regulated);
      drawTextBox(p8, font, P.workWith.lead, LAYOUT.p8.lead);
    }

    // Page 9
    const p9 = pages[8];
    if (p9 && LAYOUT.p9) {
      drawTextBox(p9, font, P.text.act_anchor, LAYOUT.p9.anchor);
    }

    const outBytes = await pdfDoc.save();
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Cache-Control", "no-store");
    res.status(200).send(Buffer.from(outBytes));
  } catch (err) {
    console.error("[fill-template] CRASH", err);
    res.status(500).json({ ok: false, error: err?.message || String(err), stack: err?.stack || null });
  }
}
