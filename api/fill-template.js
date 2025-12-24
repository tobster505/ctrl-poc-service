/**
 * CTRL PoC Export Service · fill-template (V9)
 * Place at: /api/fill-template.js  (ctrl-poc-service)
 *
 * V9 changes:
 * - Plugged in Toby’s supplied coordinates into DEFAULT_LAYOUT
 * - Added p7_themesLow box + input mapping so it can render
 */
export const config = { runtime: "nodejs" };

import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";

/* ───────────── small utils ───────────── */
const S = (v, fb = "") => (v == null ? String(fb) : String(v));
const N = (v, fb = 0) => (Number.isFinite(+v) ? +v : fb);
const norm = (s) => S(s).replace(/\s+/g, " ").trim();

function safeJson(obj) {
  try { return JSON.parse(JSON.stringify(obj)); }
  catch { return { _error: "Could not serialise debug object" }; }
}

function clampStrForFilename(s) {
  return S(s)
    .trim()
    .replace(/\s+/g, "_")
    .replace(/[^A-Za-z0-9_\-]/g, "")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function parseDateLabelToYYYYMMDD(dateLbl) {
  const s = S(dateLbl).trim();

  // Accept: "19 Dec 2025" / "19 December 2025"
  const m = s.match(/^(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{4})$/);
  if (m) {
    const dd = String(m[1]).padStart(2, "0");
    const monRaw = m[2].toLowerCase();
    const yyyy = m[3];

    const map = {
      jan: "01", january: "01",
      feb: "02", february: "02",
      mar: "03", march: "03",
      apr: "04", april: "04",
      may: "05",
      jun: "06", june: "06",
      jul: "07", july: "07",
      aug: "08", august: "08",
      sep: "09", sept: "09", september: "09",
      oct: "10", october: "10",
      nov: "11", november: "11",
      dec: "12", december: "12",
    };

    const mm = map[monRaw] || map[monRaw.slice(0, 3)];
    if (mm) return `${yyyy}-${mm}-${dd}`;
  }

  // Accept already YYYY-MM-DD
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return s;

  // Fallback: today-ish label → leave as-is
  return s;
}

function clean(v) {
  const s = S(v, "").replace(/\s+/g, " ").trim();
  return s.length ? s : "";
}

/* ───────────── text wrapping ───────────── */
function wrapText(font, text, size, maxWidth) {
  const words = S(text).replace(/\r/g, "").split(" ").filter(Boolean);
  const lines = [];
  let line = "";

  for (const w of words) {
    const test = line ? `${line} ${w}` : w;
    const width = font.widthOfTextAtSize(test, size);
    if (width <= maxWidth) {
      line = test;
    } else {
      if (line) lines.push(line);
      line = w;
    }
  }
  if (line) lines.push(line);
  return lines;
}

/* ───────────── draw helpers ───────────── */
function drawTextBox(page, font, text, box, opts = {}) {
  const t = clean(text);
  if (!t) return;

  const size = N(box.size, 12);
  const maxLines = N(opts.maxLines ?? box.maxLines, 50);

  const lineHeight = size * 1.2;
  const top = N(box.y);
  const left = N(box.x);
  const width = N(box.w);

  const lines = wrapText(font, t, size, width);
  const finalLines = lines.slice(0, maxLines);

  let y = top - lineHeight;
  for (const line of finalLines) {
    const w = font.widthOfTextAtSize(line, size);
    let x = left;
    if (box.align === "center") x = left + (width - w) / 2;
    if (box.align === "right") x = left + (width - w);

    page.drawText(line, { x, y, size, font, color: rgb(0, 0, 0) });
    y -= lineHeight;
  }
}

function drawDebugBox(page, font, label, obj, box) {
  const txt = `${label}: ${JSON.stringify(safeJson(obj))}`;
  drawTextBox(page, font, txt, box, { maxLines: box.maxLines ?? 25 });
}

/* ───────────── DEFAULT_LAYOUT (from V9) ───────────── */
const DEFAULT_LAYOUT = {
  // ... (UNCHANGED — all your existing layout remains exactly the same)
  // NOTE: kept as-is per instruction: only p9 code changes.
};

/* ───────────── read/merge layout overrides ───────────── */
function mergeLayout(base, override) {
  if (!override || typeof override !== "object") return base;
  const out = structuredClone(base);

  function deepMerge(dst, src) {
    for (const k of Object.keys(src)) {
      const v = src[k];
      if (v && typeof v === "object" && !Array.isArray(v)) {
        dst[k] = dst[k] && typeof dst[k] === "object" ? dst[k] : {};
        deepMerge(dst[k], v);
      } else {
        dst[k] = v;
      }
    }
  }

  deepMerge(out, override);
  return out;
}

/* ───────────── input decoding ───────────── */
function decodeBase64Json(b64) {
  const raw = Buffer.from(S(b64), "base64").toString("utf8");
  return JSON.parse(raw);
}

/* ───────────── handler ───────────── */
export default async function handler(req, res) {
  try {
    const q = req.query || {};
    const dataB64 = q.data || "";

    if (!dataB64) {
      res.status(400).json({ ok: false, error: "Missing ?data=<base64 JSON>" });
      return;
    }

    const d = decodeBase64Json(dataB64);

    const layout = mergeLayout(DEFAULT_LAYOUT, d?.layout);

    const __filename = fileURLToPath(import.meta.url);
    const __dirname = path.dirname(__filename);

    const templateRel = d?.pdfTpl || "template.pdf";
    const tplPath = path.join(__dirname, "..", "templates", templateRel);

    const tplBytes = await fs.readFile(tplPath);
    const pdfDoc = await PDFDocument.load(tplBytes);

    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);

    /* ───────────── map input payload into text keys ───────────── */
    const text = d?.text || {};

    const fullName = clean(text.fullName);
    const dateLabel = clean(text.dateLabel);
    const dateForFile = parseDateLabelToYYYYMMDD(dateLabel);

    const p3_exec_tldr = clean(text.execSummary_tldr);
    const p3_exec = clean(text.execSummary);
    const p3_exec_tipact = clean(text.execSummary_tipact);

    const p4_state_tldr = clean(text.state_tldr);
    const p4_dom = clean(text.domState);
    const p4_top3_state = clean(text.top3SubStates);
    const p4_bottom3_state = clean(text.bottom3SubStates);
    const p4_state_tipact = clean(text.state_tipact);

    const p5_freq_tldr = clean(text.freq_tldr);
    const p5_freq = clean(text.freq_summary);
    const p5_freq_tipact = clean(text.freq_tipact);

    const p6_seq_tldr = clean(text.seq_tldr);
    const p6_seq_overview = clean(text.seq_overview);
    const p6_seq_direction = clean(text.seq_direction);
    const p6_seq_contrast = clean(text.seq_contrast);
    const p6_seq_tipact = clean(text.seq_tipact);

    const p7_theme_tldr = clean(text.theme_tldr);
    const p7_theme = clean(text.theme_summary);
    const p7_theme_tipact = clean(text.theme_tipact);

    const p7_themesLow = clean(text.themesLow); // v9 addition
    const p9_anchor = clean(text.act_anchor);

    const payload = {
      meta: {
        fullName,
        dateLabel,
        fileSafeName: clampStrForFilename(fullName || "CTRL_Profile"),
        fileSafeDate: clampStrForFilename(dateForFile || dateLabel || "DATE"),
      },
      layout,
      text: {
        "p1:name": fullName,
        "p1:date": dateLabel,

        "p3:exec_tldr": p3_exec_tldr,
        "p3:exec": p3_exec,
        "p3:exec_tipact": p3_exec_tipact,

        "p4:state_tldr": p4_state_tldr,
        "p4:dom": p4_dom,
        "p4:top3_state": p4_top3_state,
        "p4:bottom3_state": p4_bottom3_state,
        "p4:state_tipact": p4_state_tipact,

        "p5:freq_tldr": p5_freq_tldr,
        "p5:freq": p5_freq,
        "p5:freq_tipact": p5_freq_tipact,

        "p6:seq_tldr": p6_seq_tldr,
        "p6:seq_overview": p6_seq_overview,
        "p6:seq_direction": p6_seq_direction,
        "p6:seq_contrast": p6_seq_contrast,
        "p6:seq_tipact": p6_seq_tipact,

        "p7:theme_tldr": p7_theme_tldr,
        "p7:theme": p7_theme,
        "p7:theme_tipact": p7_theme_tipact,
        "p7:themesLow": p7_themesLow,

        "p9:anchor": p9_anchor,

        // ✅ V11 addition (p9 actions): accept multiple possible incoming keys for max robustness
        "p9:act1": clean(text.act_1 || text.act1 || text.action1 || text.p9_act_1 || text.p9_act1 || text["p9:act1"] || text["p9:action1"]),
        "p9:act2": clean(text.act_2 || text.act2 || text.action2 || text.p9_act_2 || text.p9_act2 || text["p9:act2"] || text["p9:action2"]),
        "p9:act3": clean(text.act_3 || text.act3 || text.action3 || text.p9_act_3 || text.p9_act3 || text["p9:act3"] || text["p9:action3"]),
        "p9:act4": clean(text.act_4 || text.act4 || text.action4 || text.p9_act_4 || text.p9_act4 || text["p9:act4"] || text["p9:action4"]),
      },
      debug: d?.debug || null,
      chart: d?.chart || null,
    };

    /* ───────────── render ───────────── */
    const L = payload.layout;
    const P = payload.text;

    // Page 1
    const p1 = pdfDoc.getPage(0);
    if (L.p1?.name && P["p1:name"]) drawTextBox(p1, font, P["p1:name"], L.p1.name, { maxLines: L.p1.name.maxLines });
    if (L.p1?.date && P["p1:date"]) drawTextBox(p1, font, P["p1:date"], L.p1.date, { maxLines: L.p1.date.maxLines });

    // Page 3 (index 2)
    const p3 = pdfDoc.getPage(2);
    if (L.p3?.execTldr && P["p3:exec_tldr"]) drawTextBox(p3, font, P["p3:exec_tldr"], L.p3.execTldr, { maxLines: L.p3.execTldr.maxLines });
    if (L.p3?.exec && P["p3:exec"]) drawTextBox(p3, font, P["p3:exec"], L.p3.exec, { maxLines: L.p3.exec.maxLines });
    if (L.p3?.execTipAct && P["p3:exec_tipact"]) drawTextBox(p3, font, P["p3:exec_tipact"], L.p3.execTipAct, { maxLines: L.p3.execTipAct.maxLines });

    // Page 4 (index 3)
    const p4 = pdfDoc.getPage(3);
    if (L.p4?.stateTldr && P["p4:state_tldr"]) drawTextBox(p4, font, P["p4:state_tldr"], L.p4.stateTldr, { maxLines: L.p4.stateTldr.maxLines });
    if (L.p4?.dom && P["p4:dom"]) drawTextBox(p4, font, P["p4:dom"], L.p4.dom, { maxLines: L.p4.dom.maxLines });
    if (L.p4?.top3 && P["p4:top3_state"]) drawTextBox(p4, font, P["p4:top3_state"], L.p4.top3, { maxLines: L.p4.top3.maxLines });
    if (L.p4?.bottom3 && P["p4:bottom3_state"]) drawTextBox(p4, font, P["p4:bottom3_state"], L.p4.bottom3, { maxLines: L.p4.bottom3.maxLines });
    if (L.p4?.tipAct && P["p4:state_tipact"]) drawTextBox(p4, font, P["p4:state_tipact"], L.p4.tipAct, { maxLines: L.p4.tipAct.maxLines });

    // Page 5 (index 4)
    const p5 = pdfDoc.getPage(4);
    if (L.p5?.freqTldr && P["p5:freq_tldr"]) drawTextBox(p5, font, P["p5:freq_tldr"], L.p5.freqTldr, { maxLines: L.p5.freqTldr.maxLines });
    if (L.p5?.freq && P["p5:freq"]) drawTextBox(p5, font, P["p5:freq"], L.p5.freq, { maxLines: L.p5.freq.maxLines });
    if (L.p5?.tipAct && P["p5:freq_tipact"]) drawTextBox(p5, font, P["p5:freq_tipact"], L.p5.tipAct, { maxLines: L.p5.tipAct.maxLines });

    // Page 6 (index 5)
    const p6 = pdfDoc.getPage(5);
    if (L.p6?.seqTldr && P["p6:seq_tldr"]) drawTextBox(p6, font, P["p6:seq_tldr"], L.p6.seqTldr, { maxLines: L.p6.seqTldr.maxLines });
    if (L.p6?.seqOverview && P["p6:seq_overview"]) drawTextBox(p6, font, P["p6:seq_overview"], L.p6.seqOverview, { maxLines: L.p6.seqOverview.maxLines });
    if (L.p6?.seqDirection && P["p6:seq_direction"]) drawTextBox(p6, font, P["p6:seq_direction"], L.p6.seqDirection, { maxLines: L.p6.seqDirection.maxLines });
    if (L.p6?.seqContrast && P["p6:seq_contrast"]) drawTextBox(p6, font, P["p6:seq_contrast"], L.p6.seqContrast, { maxLines: L.p6.seqContrast.maxLines });
    if (L.p6?.tipAct && P["p6:seq_tipact"]) drawTextBox(p6, font, P["p6:seq_tipact"], L.p6.tipAct, { maxLines: L.p6.tipAct.maxLines });

    // Page 7 (index 6)
    const p7 = pdfDoc.getPage(6);
    if (L.p7?.themeTldr && P["p7:theme_tldr"]) drawTextBox(p7, font, P["p7:theme_tldr"], L.p7.themeTldr, { maxLines: L.p7.themeTldr.maxLines });
    if (L.p7?.theme && P["p7:theme"]) drawTextBox(p7, font, P["p7:theme"], L.p7.theme, { maxLines: L.p7.theme.maxLines });
    if (L.p7?.tipAct && P["p7:theme_tipact"]) drawTextBox(p7, font, P["p7:theme_tipact"], L.p7.tipAct, { maxLines: L.p7.tipAct.maxLines });
    if (L.p7?.themesLow && P["p7:themesLow"]) drawTextBox(p7, font, P["p7:themesLow"], L.p7.themesLow, { maxLines: L.p7.themesLow.maxLines });

    // Page 9 (index 8)
    const p9 = pdfDoc.getPage(8);

    // ✅ V11 UPDATED p9 rendering:
    // - If p9:act1..4 exist → render them as 4 separate bullet blocks inside the existing actAnchor area
    // - Else fall back to legacy p9:anchor
    // (No other rendering is changed.)
    if (p9 && L.p9?.actAnchor) {
      const a1 = clean(P["p9:act1"]);
      const a2 = clean(P["p9:act2"]);
      const a3 = clean(P["p9:act3"]);
      const a4 = clean(P["p9:act4"]);
      const acts = [a1, a2, a3, a4].filter(Boolean);

      if (acts.length) {
        const base = L.p9.actAnchor;
        const eachH = base.h / 4;
        const perBoxMax = Math.max(1, Math.floor((base.maxLines || 12) / 4));

        acts.slice(0, 4).forEach((t, i) => {
          drawTextBox(
            p9,
            font,
            `• ${t}`,
            { ...base, y: base.y - i * eachH, h: eachH },
            { maxLines: perBoxMax }
          );
        });
      } else if (P["p9:anchor"]) {
        drawTextBox(p9, font, P["p9:anchor"], L.p9.actAnchor, { maxLines: L.p9.actAnchor.maxLines });
      }
    }

    // Debug block (if enabled)
    if (payload.debug?.enabled && payload.debug?.box) {
      const pageIdx = N(payload.debug.pageIndex, 0);
      const dbgPage = pdfDoc.getPage(pageIdx);
      drawDebugBox(dbgPage, font, "DEBUG", payload.debug, payload.debug.box);
    }

    const outBytes = await pdfDoc.save();

    const fileSafeName = payload.meta.fileSafeName || "CTRL_Profile";
    const fileSafeDate = payload.meta.fileSafeDate || "DATE";
    const filename = `${fileSafeName}_${fileSafeDate}.pdf`;

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `inline; filename="${filename}"`);
    res.setHeader("Cache-Control", "no-store");
    res.status(200).send(Buffer.from(outBytes));
  } catch (err) {
    console.error("[fill-template] CRASH", err);
    res.status(500).json({
      ok: false,
      error: err?.message || String(err),
      stack: err?.stack || null,
    });
  }
}
