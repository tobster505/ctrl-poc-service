/**
 * CTRL PoC Export Service · fill-template (V14.0 · aligned to V15 build payload)
 *
 * Changes in this update:
 * - Fully aligned to the V15 Build PDF payload contract
 * - Treats direct paragraph fields as primary:
 *   - text.snapshot_p1..p4
 *   - text.chart_p1..p5
 *   - text.move_p1..p5
 *   - text.themes_p1..p2
 *   - text.interact_c / interact_t / interact_r / interact_l
 *   - text.act_1 / act_2 / act_3
 * - Uses rebuilt compatibility fields when present:
 *   - text.snapshot
 *   - text.chart_overview
 *   - text.awareness_movement
 *   - text.themes
 *   - text.interactions_with_others
 *   - text.actions
 *   - text.actions_bullets
 * - Validates parent-vs-slot consistency in debug probe
 * - Validates actions array / act slots / bullets consistency in debug probe
 * - Supports V15 identity and ctrl structure:
 *   - identity.fullName / preferredName / email / dateLabel
 *   - ctrl.dominantKey / secondKey / dominantSubState / templateKey / bands
 * - POST remains the primary transport
 * - Accepts both raw JSON payloads and wrapped bodies like { data: payload }
 * - GET ?data=... fallback retained only for backwards compatibility
 * - Keeps the 16-page user template structure:
 *   1  Cover
 *   2  Table of Contents
 *   3  How to Read This Profile
 *   4–5  Snapshot
 *   6–8  Distribution
 *   9–11 Movement
 *   12 Theme Overview
 *   13–14 Interaction with Others
 *   15 Actions
 *   16 Legal
 * - Correct template selection based on dominant + second key:
 *   e.g. RL -> CTRL_PoC_Assessment_Profile_template_RL.pdf
 * - True PDF fallback support:
 *   CTRL_PoC_Assessment_Profile_fallback.pdf
 * - Retains debug mode (?debug=1)
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
const okObj = (o) => o && typeof o === "object" && !Array.isArray(o);
const okArr = (a) => Array.isArray(a);

function safeJson(obj) {
  try { return JSON.parse(JSON.stringify(obj)); }
  catch { return { _error: "Could not serialise debug object" }; }
}

function clean(v) {
  return S(v).trim();
}

function strEq(a, b) {
  return clean(a).replace(/\r/g, "") === clean(b).replace(/\r/g, "");
}

function joinParas(arr) {
  return (arr || [])
    .map((x) => clean(x))
    .filter(Boolean)
    .join("\n\n")
    .trim();
}

function paraCount(s) {
  const t = clean(s);
  if (!t) return 0;
  return t
    .split(/\n\s*\n/)
    .map((x) => clean(x))
    .filter(Boolean)
    .length;
}

function bulletLineCount(s) {
  return clean(s)
    ? String(s)
        .split("\n")
        .map((x) => x.trim())
        .filter((x) => x.startsWith("• ")).length
    : 0;
}

function uniqueTrim(arr) {
  const out = [];
  const seen = new Set();

  for (const x of (arr || [])) {
    const v = clean(x);
    if (!v) continue;
    const key = v.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(v);
  }
  return out;
}

function toBullets(arr) {
  if (!okArr(arr)) return "";
  return arr
    .map((x) => clean(x))
    .filter(Boolean)
    .map((x) => `• ${x}`)
    .join("\n");
}

/* ───────── filename helpers ───────── */
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

  const m = s.match(/^(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]{3,9})\s+(\d{4})$/);
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

  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;

  return clampStrForFilename(s || "date");
}

function makeOutputFilename(fullName, dateLbl) {
  const parts = S(fullName).trim().split(/\s+/).filter(Boolean);
  const first = parts[0] || "First";
  const last = parts.length > 1 ? parts[parts.length - 1] : "Surname";
  const datePart = parseDateLabelToYYYYMMDD(dateLbl);
  const fn = clampStrForFilename(first);
  const ln = clampStrForFilename(last);
  return `PoC_Profile_${fn}_${ln}_${datePart}.pdf`;
}

/* ───────── text wrapping + drawing ───────── */
function wrapText(font, text, size, w) {
  const raw = S(text);
  const paragraphs = raw.split("\n");
  const lines = [];

  for (const para of paragraphs) {
    const words = para.split(/\s+/).filter(Boolean);
    let line = "";

    if (!words.length) {
      lines.push("");
      continue;
    }

    for (const word of words) {
      const test = line ? `${line} ${word}` : word;
      const width = font.widthOfTextAtSize(test, size);

      if (width <= w) line = test;
      else {
        if (line) lines.push(line);
        line = word;
      }
    }

    if (line) lines.push(line);
  }

  while (lines.length && lines[lines.length - 1] === "") lines.pop();
  return lines;
}

function drawTextBox(page, font, text, box, opts = {}) {
  if (!page || !font || !box) return;

  const t0 = S(text);
  if (!t0.trim()) return;

  const pageH = page.getHeight();

  const size = N(opts.size ?? box.size ?? 12);
  const lineGap = N(opts.lineGap ?? box.lineGap ?? 2);
  const maxLines = N(opts.maxLines ?? box.maxLines ?? 999);
  const alignRaw = String(opts.align ?? box.align ?? "left").toLowerCase();
  const align = (alignRaw === "centre") ? "center" : alignRaw;
  const pad = N(opts.pad ?? box.pad ?? 0);

  let x = N(box.x);
  let w = Math.max(0, N(box.w));
  let h = Math.max(0, N(box.h));

  const autoExpand = (opts.autoExpand ?? box.autoExpand ?? true) !== false;
  if (autoExpand && Number.isFinite(maxLines) && maxLines > 0) {
    const lineHeight = size + lineGap;
    const hNeeded = (pad * 2) + size + (Math.max(0, maxLines - 1) * lineHeight);
    h = Math.max(h, hNeeded);
  }

  const y = pageH - N(box.y) - h;
  const innerW = Math.max(0, w - pad * 2);
  const lines = wrapText(font, t0.replace(/\r/g, ""), size, innerW).slice(0, maxLines);

  const lineHeight = size + lineGap;
  let cursorY = y + h - pad - size;

  for (let i = 0; i < lines.length; i++) {
    const ln = lines[i];
    let dx = x + pad;

    if (align !== "left") {
      const lw = font.widthOfTextAtSize(ln, size);
      if (align === "center") dx = x + (w - lw) / 2;
      if (align === "right") dx = x + w - pad - lw;
    }

    page.drawText(ln, {
      x: dx,
      y: cursorY,
      size,
      font,
      color: rgb(0, 0, 0)
    });

    cursorY -= lineHeight;
  }
}

/* ───────── paragraph helpers ───────── */
function splitParasFixed(raw, expected) {
  const parts = S(raw)
    .replace(/\r/g, "")
    .trim()
    .split(/\n\s*\n/)
    .map((p) => p.trim())
    .filter(Boolean);

  while (parts.length < expected) parts.push("");
  return parts.slice(0, expected);
}

/* ───────── template loader ───────── */
async function loadTemplateBytesLocal(fname) {
  if (!fname.endsWith(".pdf")) {
    throw new Error(`Invalid template filename: ${fname}`);
  }

  const __file = fileURLToPath(import.meta.url);
  const __dir = path.dirname(__file);

  const candidates = [
    path.join(process.cwd(), "public", fname),
    path.join(__dir, "..", "public", fname),
    path.join(__dir, "public", fname),
    path.join(__dir, fname),
  ];

  let lastErr;
  for (const pth of candidates) {
    try {
      return { bytes: await fs.readFile(pth), path: pth };
    } catch (err) {
      lastErr = err;
    }
  }

  throw new Error(
    `Template not found: ${fname}. Tried: ${candidates.join(" | ")}. Last: ${lastErr?.message || "no detail"}`
  );
}

/* ───────── payload parsing ───────── */
async function readPayload(req) {
  if (req.method === "POST") {
    const chunks = [];
    for await (const ch of req) chunks.push(ch);
    const raw = Buffer.concat(chunks).toString("utf8") || "{}";

    let parsed = {};
    try {
      parsed = JSON.parse(raw);
    } catch {
      return {};
    }

    if (okObj(parsed?.data)) return parsed.data;
    if (okObj(parsed?.payload)) return parsed.payload;
    if (okObj(parsed?.body)) return parsed.body;
    if (okObj(parsed)) return parsed;
    return {};
  }

  const url = new URL(req.url, "http://localhost");
  const dataB64 = url.searchParams.get("data") || "";
  if (!dataB64) return {};

  try {
    const raw = Buffer.from(dataB64, "base64").toString("utf8");
    const parsed = JSON.parse(raw);
    if (okObj(parsed?.data)) return parsed.data;
    if (okObj(parsed?.payload)) return parsed.payload;
    if (okObj(parsed?.body)) return parsed.body;
    return okObj(parsed) ? parsed : {};
  } catch {
    return {};
  }
}

/* ───────── dom/second detection ───────── */
function resolveStateKey(any) {
  const s = S(any).trim().toUpperCase();
  const c = s.charAt(0);
  if (["C", "T", "R", "L"].includes(c)) return c;

  const low = S(any).toLowerCase();
  if (low.includes("concealed")) return "C";
  if (low.includes("triggered")) return "T";
  if (low.includes("regulated")) return "R";
  if (low.includes("lead")) return "L";
  return null;
}

function computeDomAndSecondKeys(P) {
  const raw = P.raw || {};
  const ctrl = okObj(raw.ctrl) ? raw.ctrl : {};
  const summary = okObj(ctrl.summary) ? ctrl.summary : {};
  const results = okObj(summary.results) ? summary.results : {};

  const domKey =
    resolveStateKey(ctrl.dominantKey) ||
    resolveStateKey(results.dominant) ||
    resolveStateKey(summary.dominant) ||
    resolveStateKey(raw.dominantKey) ||
    resolveStateKey(raw.domKey) ||
    resolveStateKey(raw.dominant) ||
    "R";

  const secondKey =
    resolveStateKey(ctrl.secondKey) ||
    resolveStateKey(results.secondState) ||
    resolveStateKey(summary.secondState) ||
    resolveStateKey(raw.secondKey) ||
    resolveStateKey(raw.secondState) ||
    "T";

  return { domKey, secondKey, templateKey: `${domKey}${secondKey}` };
}

/* ───────── chart embed ───────── */
function makeSpiderChartUrl12(bandsRaw) {
  const keys = [
    "C_low","C_mid","C_high",
    "T_low","T_mid","T_high",
    "R_low","R_mid","R_high",
    "L_low","L_mid","L_high",
  ];

  const displayLabels = [
    "", "Concealed", "",
    "", "Triggered", "",
    "", "Regulated", "",
    "", "Lead", ""
  ];

  const vals = keys.map((k) => Number(bandsRaw?.[k] || 0));
  const maxVal = Math.max(...vals, 1);
  const data = vals.map((v) => (maxVal > 0 ? v / maxVal : 0));

  const CTRL_COLOURS = {
    C: { low: "rgba(230, 228, 225, 0.55)", mid: "rgba(184, 180, 174, 0.55)", high: "rgba(110, 106, 100, 0.55)" },
    T: { low: "rgba(244, 225, 198, 0.55)", mid: "rgba(211, 155,  74, 0.55)", high: "rgba(154,  94,  26, 0.55)" },
    R: { low: "rgba(226, 236, 230, 0.55)", mid: "rgba(143, 183, 161, 0.55)", high: "rgba( 79, 127, 105, 0.55)" },
    L: { low: "rgba(230, 220, 227, 0.55)", mid: "rgba(164, 135, 159, 0.55)", high: "rgba( 94,  63,  90, 0.55)" },
  };

  const colours = keys.map((k) => {
    const state = k[0];
    const tier = k.split("_")[1];
    return CTRL_COLOURS[state]?.[tier] || "rgba(0,0,0,0.10)";
  });

  const startAngle = -Math.PI / 4;

  const cfg = {
    type: "polarArea",
    data: {
      labels: displayLabels,
      datasets: [{
        data,
        backgroundColor: colours,
        borderWidth: 3,
        borderColor: "rgba(0, 0, 0, 0.20)",
      }],
    },
    options: {
      plugins: { legend: { display: false } },
      startAngle,
      scales: {
        r: {
          startAngle,
          min: 0,
          max: 1,
          ticks: { display: false },
          grid: { display: true },
          angleLines: { display: false },
          pointLabels: {
            display: true,
            padding: 14,
            font: { size: 26, weight: "bold" },
            centerPointLabels: true,
          },
        },
      },
    },
  };

  const enc = encodeURIComponent(JSON.stringify(cfg));
  return `https://quickchart.io/chart?c=${enc}&format=png&width=900&height=900&backgroundColor=transparent&version=4`;
}

async function embedRemoteImage(pdfDoc, url) {
  if (!url) return null;

  const res = await fetch(url);
  if (!res.ok) throw new Error(`Failed to fetch chart: ${res.status} ${res.statusText}`);

  const buf = new Uint8Array(await res.arrayBuffer());
  const sig = String.fromCharCode(buf[0], buf[1], buf[2], buf[3] || 0);

  if (sig.startsWith("\x89PNG")) return await pdfDoc.embedPng(buf);
  if (sig.startsWith("\xff\xd8")) return await pdfDoc.embedJpg(buf);

  try { return await pdfDoc.embedPng(buf); }
  catch { return await pdfDoc.embedJpg(buf); }
}

async function embedRadarFromBandsOrUrl(pdfDoc, page, box, bandsRaw, chartUrl) {
  if (!pdfDoc || !page || !box) return;

  let url = S(chartUrl).trim();
  if (!url) {
    const hasAny = bandsRaw && typeof bandsRaw === "object" &&
      Object.values(bandsRaw).some((v) => Number(v) > 0);
    if (!hasAny) return;
    url = makeSpiderChartUrl12(bandsRaw);
  }

  const img = await embedRemoteImage(pdfDoc, url);
  if (!img) return;

  const H = page.getHeight();
  page.drawImage(img, {
    x: box.x,
    y: H - box.y - box.h,
    width: box.w,
    height: box.h
  });
}

/* ───────── DEFAULT LAYOUT (16-page mapping) ───────── */
const DEFAULT_LAYOUT = {
  pages: {
    p1: {
      name: { x: 60, y: 458, w: 500, h: 60, size: 30, align: "center", maxLines: 1 },
      date: { x: 230, y: 613, w: 500, h: 40, size: 25, align: "left", maxLines: 1 },
    },

    p2:  { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p3:  { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p4:  { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p5:  { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p6:  { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p7:  { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p8:  { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p9:  { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p10: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p11: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p12: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p13: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p14: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p15: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },

    twoParaPage: {
      top:    { x: 25, y: 180, w: 550, h: 240, size: 16, align: "left", maxLines: 13 },
      bottom: { x: 25, y: 470, w: 550, h: 280, size: 16, align: "left", maxLines: 15 },
    },

    oneParaPage: {
      body:   { x: 25, y: 180, w: 550, h: 520, size: 16, align: "left", maxLines: 28 },
    },

    p6Distribution: {
      p1:    { x: 25,  y: 160, w: 200, h: 240, size: 16, align: "left", maxLines: 12 },
      p2:    { x: 25,  y: 500, w: 550, h: 240, size: 16, align: "left", maxLines: 13 },
      chart: { x: 250, y: 160, w: 320, h: 320 }
    },

    p15Actions: {
      act1: { x: 50,  y: 380, w: 440, h: 95, size: 17, align: "left", maxLines: 5 },
      act2: { x: 100, y: 530, w: 440, h: 95, size: 17, align: "left", maxLines: 5 },
      act3: { x: 50,  y: 670, w: 440, h: 95, size: 17, align: "left", maxLines: 5 },
    },
  },
};

/* ───────── URL layout overrides ───────── */
function applyLayoutOverridesFromUrl(layoutPages, url) {
  const allowed = new Set(["x", "y", "w", "h", "size", "maxLines", "align"]);
  const applied = [];
  const ignored = [];

  for (const [k, v] of url.searchParams.entries()) {
    if (!k.startsWith("L_")) continue;

    const bits = k.split("_");
    if (bits.length < 4) { ignored.push({ k, v, why: "bad_key_shape" }); continue; }

    const pageKey = bits[1];
    const boxKey = bits[2];
    const prop = bits.slice(3).join("_");

    if (!layoutPages?.[pageKey]) { ignored.push({ k, v, why: "unknown_page", pageKey }); continue; }
    if (!layoutPages?.[pageKey]?.[boxKey]) { ignored.push({ k, v, why: "unknown_box", pageKey, boxKey }); continue; }
    if (!allowed.has(prop)) { ignored.push({ k, v, why: "unsupported_prop", prop }); continue; }

    if (prop === "align") {
      const a0 = String(v || "").toLowerCase();
      const a = (a0 === "centre") ? "center" : a0;
      if (!["left", "center", "right"].includes(a)) { ignored.push({ k, v, why: "bad_align", got: a0 }); continue; }
      layoutPages[pageKey][boxKey][prop] = a;
      applied.push({ k, v, pageKey, boxKey, prop });
      continue;
    }

    const num = Number(v);
    if (!Number.isFinite(num)) { ignored.push({ k, v, why: "not_a_number" }); continue; }

    layoutPages[pageKey][boxKey][prop] = (prop === "maxLines") ? Math.max(0, Math.floor(num)) : num;
    applied.push({ k, v, pageKey, boxKey, prop });
  }

  return { applied, ignored, layoutPages };
}

/* ───────── actions normaliser ───────── */
function ensureActions(text) {
  let arr = [];

  if (okArr(text?.actions)) {
    arr = text.actions.map((x) => clean(x)).filter(Boolean);
  }

  const a1 = clean(text?.act_1 || text?.Act1);
  const a2 = clean(text?.act_2 || text?.Act2);
  const a3 = clean(text?.act_3 || text?.Act3);

  if (a1 || a2 || a3) {
    arr = [a1, a2, a3].filter(Boolean);
  }

  const bulletsRaw = clean(text?.actions_bullets);
  const bulletsArr = bulletsRaw
    ? bulletsRaw
        .split("\n")
        .map((x) => clean(x.replace(/^•\s*/, "")))
        .filter(Boolean)
    : [];

  if (!arr.length && bulletsArr.length) {
    arr = bulletsArr;
  }

  arr = uniqueTrim(arr);

  const fallbackActions = [
    "Notice what happens in the first few seconds after feedback lands.",
    "Try naming the feeling internally before you respond.",
    "Choose one moment this week to pause before replying."
  ];

  for (let i = 0; arr.length < 3 && i < fallbackActions.length; i++) {
    const cand = clean(fallbackActions[i]);
    if (!cand) continue;
    if (!arr.some((a) => a.toLowerCase() === cand.toLowerCase())) arr.push(cand);
  }

  while (arr.length < 3) {
    arr.push("Pick one small moment to pause and notice your first reaction.");
  }

  arr = uniqueTrim(arr).slice(0, 3);

  return {
    arr,
    act_1: clean(a1 || arr[0] || ""),
    act_2: clean(a2 || arr[1] || ""),
    act_3: clean(a3 || arr[2] || ""),
    bullets: bulletsRaw || toBullets(arr)
  };
}

/* ───────── input normaliser (aligned to V15 payload contract) ───────── */
function normaliseInput(d = {}) {
  const identity = okObj(d.identity) ? d.identity : {};
  const text = okObj(d.text) ? d.text : {};
  const ctrl = okObj(d.ctrl) ? d.ctrl : {};
  const summary = okObj(ctrl.summary) ? ctrl.summary : {};
  const chart = okObj(d.chart) ? d.chart : {};

  const fullName = S(
    identity.fullName ||
    d.fullName ||
    d.FullName ||
    summary?.identity?.user?.fullName ||
    summary?.identity?.fullName ||
    ""
  ).trim();

  const preferredName = S(
    identity.preferredName ||
    d.preferredName ||
    d.PreferredName ||
    summary?.identity?.user?.preferredName ||
    ""
  ).trim();

  const email = S(
    identity.email ||
    d.email ||
    d.Email ||
    summary?.identity?.user?.email ||
    summary?.identity?.email ||
    ""
  ).trim();

  const dateLabel = S(
    identity.dateLabel ||
    d.dateLbl ||
    d.dateLabel ||
    d.date ||
    d.Date ||
    summary?.dateLbl ||
    ""
  ).trim();

  const bandsRaw =
    (okObj(ctrl.bands) && Object.keys(ctrl.bands).length ? ctrl.bands : null) ||
    (okObj(summary.ctrl12) && Object.keys(summary.ctrl12).length ? summary.ctrl12 : null) ||
    (okObj(d.bands) && Object.keys(d.bands).length ? d.bands : null) ||
    {};

  const snapshot = S(text.snapshot || "");
  const chartOverview = S(text.chart_overview || "");
  const movement = S(text.awareness_movement || "");
  const themes = S(text.themes || "");
  const interactions = S(text.interactions_with_others || "");

  const snapshotParts = splitParasFixed(snapshot, 4);
  const chartParts = splitParasFixed(chartOverview, 5);
  const moveParts = splitParasFixed(movement, 5);
  const themeParts = splitParasFixed(themes, 2);
  const interactParts = splitParasFixed(interactions, 4);

  const ensuredActions = ensureActions(text);

  const out = {
    raw: d,

    identity: {
      fullName,
      preferredName,
      email,
      dateLabel
    },

    ctrl: {
      dominantKey: clean(ctrl.dominantKey || d.domKey || d.dominantKey || ""),
      secondKey: clean(ctrl.secondKey || d.secondKey || ""),
      dominantSubState: clean(ctrl.dominantSubState || d.domSub || ""),
      templateKey: clean(ctrl.templateKey || d.templateKey || "")
    },

    bands: bandsRaw,

    // parent strings
    snapshot,
    chart_overview: chartOverview,
    awareness_movement: movement,
    themes,
    interactions_with_others: interactions,

    // direct slots
    snapshot_p1: S(text.snapshot_p1 || snapshotParts[0]),
    snapshot_p2: S(text.snapshot_p2 || snapshotParts[1]),
    snapshot_p3: S(text.snapshot_p3 || snapshotParts[2]),
    snapshot_p4: S(text.snapshot_p4 || snapshotParts[3]),

    chart_p1: S(text.chart_p1 || chartParts[0]),
    chart_p2: S(text.chart_p2 || chartParts[1]),
    chart_p3: S(text.chart_p3 || chartParts[2]),
    chart_p4: S(text.chart_p4 || chartParts[3]),
    chart_p5: S(text.chart_p5 || chartParts[4]),

    move_p1: S(text.move_p1 || moveParts[0]),
    move_p2: S(text.move_p2 || moveParts[1]),
    move_p3: S(text.move_p3 || moveParts[2]),
    move_p4: S(text.move_p4 || moveParts[3]),
    move_p5: S(text.move_p5 || moveParts[4]),

    themes_p1: S(text.themes_p1 || themeParts[0]),
    themes_p2: S(text.themes_p2 || themeParts[1]),

    interact_c: S(text.interact_c || interactParts[0]),
    interact_t: S(text.interact_t || interactParts[1]),
    interact_r: S(text.interact_r || interactParts[2]),
    interact_l: S(text.interact_l || interactParts[3]),

    actions: ensuredActions.arr,
    actions_bullets: ensuredActions.bullets,

    act_1: ensuredActions.act_1,
    act_2: ensuredActions.act_2,
    act_3: ensuredActions.act_3,

    Act1: ensuredActions.act_1,
    Act2: ensuredActions.act_2,
    Act3: ensuredActions.act_3,

    chartUrl: S(chart.spiderUrl || chart.url || d.spiderChartUrl || d.spider_chart_url || "").trim()
  };

  return out;
}

/* ───────── debug probe ───────── */
function buildProbe(P, domSecond, templateInfo, ov, L) {
  const expectedSnapshot = joinParas([P.snapshot_p1, P.snapshot_p2, P.snapshot_p3, P.snapshot_p4]);
  const expectedChart = joinParas([P.chart_p1, P.chart_p2, P.chart_p3, P.chart_p4, P.chart_p5]);
  const expectedMovement = joinParas([P.move_p1, P.move_p2, P.move_p3, P.move_p4, P.move_p5]);
  const expectedThemes = joinParas([P.themes_p1, P.themes_p2]);
  const expectedInteractions = joinParas([P.interact_c, P.interact_t, P.interact_r, P.interact_l]);
  const expectedBullets = [P.act_1, P.act_2, P.act_3]
    .map((x) => clean(x))
    .filter(Boolean)
    .map((x) => `• ${x}`)
    .join("\n");

  return {
    ok: true,
    where: "fill-template:V14.0:debug",
    templateSelection: safeJson(templateInfo),
    domSecond: safeJson(domSecond),

    identity: {
      fullName: P.identity.fullName,
      preferredName: P.identity.preferredName,
      email: P.identity.email,
      dateLabel: P.identity.dateLabel
    },

    parentLengths: {
      snapshot: S(P.snapshot).length,
      chart_overview: S(P.chart_overview).length,
      awareness_movement: S(P.awareness_movement).length,
      themes: S(P.themes).length,
      interactions_with_others: S(P.interactions_with_others).length,
      actions_bullets: S(P.actions_bullets).length
    },

    parentParagraphCounts: {
      snapshot: paraCount(P.snapshot),
      chart_overview: paraCount(P.chart_overview),
      awareness_movement: paraCount(P.awareness_movement),
      themes: paraCount(P.themes),
      interactions_with_others: paraCount(P.interactions_with_others)
    },

    textLengths: {
      snapshot_p1: S(P.snapshot_p1).length,
      snapshot_p2: S(P.snapshot_p2).length,
      snapshot_p3: S(P.snapshot_p3).length,
      snapshot_p4: S(P.snapshot_p4).length,

      chart_p1: S(P.chart_p1).length,
      chart_p2: S(P.chart_p2).length,
      chart_p3: S(P.chart_p3).length,
      chart_p4: S(P.chart_p4).length,
      chart_p5: S(P.chart_p5).length,

      move_p1: S(P.move_p1).length,
      move_p2: S(P.move_p2).length,
      move_p3: S(P.move_p3).length,
      move_p4: S(P.move_p4).length,
      move_p5: S(P.move_p5).length,

      themes_p1: S(P.themes_p1).length,
      themes_p2: S(P.themes_p2).length,

      interact_c: S(P.interact_c).length,
      interact_t: S(P.interact_t).length,
      interact_r: S(P.interact_r).length,
      interact_l: S(P.interact_l).length,

      act_1: S(P.act_1).length,
      act_2: S(P.act_2).length,
      act_3: S(P.act_3).length
    },

    parentMatchesSlots: {
      snapshot: !clean(P.snapshot) ? null : strEq(P.snapshot, expectedSnapshot),
      chart_overview: !clean(P.chart_overview) ? null : strEq(P.chart_overview, expectedChart),
      awareness_movement: !clean(P.awareness_movement) ? null : strEq(P.awareness_movement, expectedMovement),
      themes: !clean(P.themes) ? null : strEq(P.themes, expectedThemes),
      interactions_with_others: !clean(P.interactions_with_others) ? null : strEq(P.interactions_with_others, expectedInteractions)
    },

    actions: {
      arrayLength: okArr(P.actions) ? P.actions.length : 0,
      bulletLines: bulletLineCount(P.actions_bullets),
      act1_matches_actions0: P.actions?.[0] ? strEq(P.act_1, P.actions[0]) : null,
      act2_matches_actions1: P.actions?.[1] ? strEq(P.act_2, P.actions[1]) : null,
      act3_matches_actions2: P.actions?.[2] ? strEq(P.act_3, P.actions[2]) : null,
      bullets_match_acts: !clean(P.actions_bullets) ? null : strEq(P.actions_bullets, expectedBullets)
    },

    layoutOverrides: {
      appliedCount: ov?.applied?.length || 0,
      ignoredCount: ov?.ignored?.length || 0,
      applied: ov?.applied || [],
      ignored: ov?.ignored || [],
      resolvedExamples: {
        p15Actions_act1: safeJson(L?.p15Actions?.act1 || {}),
        p15Actions_act2: safeJson(L?.p15Actions?.act2 || {}),
        p15Actions_act3: safeJson(L?.p15Actions?.act3 || {}),
      },
    },
  };
}

/* ───────── main handler ───────── */
export default async function handler(req, res) {
  try {
    const url = new URL(req.url, "http://localhost");
    const debug = url.searchParams.get("debug") === "1";
    const method = String(req.method || "GET").toUpperCase();

    if (!["GET", "POST"].includes(method)) {
      res.setHeader("Allow", "GET, POST");
      return res.status(405).json({
        ok: false,
        error: `Method ${method} not allowed`,
        allowed: ["GET", "POST"]
      });
    }

    if (method === "GET" && !debug) {
      const hasLegacyData = !!url.searchParams.get("data");
      if (!hasLegacyData) {
        return res.status(400).json({
          ok: false,
          error: "No payload supplied. Use POST with a JSON body, or GET only with legacy ?data=... support."
        });
      }
    }

    const payload = await readPayload(req);
    const P = normaliseInput(payload);
    const domSecond = computeDomAndSecondKeys({ raw: payload });

    const validCombos = new Set([
      "CT","CL","CR",
      "TC","TR","TL",
      "RC","RT","RL",
      "LC","LR","LT"
    ]);

    const comboIsValid = validCombos.has(domSecond.templateKey);
    const selectedTemplate = comboIsValid
      ? `CTRL_PoC_Assessment_Profile_template_${domSecond.templateKey}.pdf`
      : "CTRL_PoC_Assessment_Profile_fallback.pdf";

    const fallbackTemplate = "CTRL_PoC_Assessment_Profile_fallback.pdf";

    const templateInfo = {
      domKey: domSecond.domKey,
      secondKey: domSecond.secondKey,
      templateKey: domSecond.templateKey,
      comboIsValid,
      selectedTemplate,
      fallbackTemplate,
      usedTemplate: "",
      usedFallback: false,
      selectedTemplatePath: "",
      fallbackTemplatePath: ""
    };

    const L = safeJson(DEFAULT_LAYOUT.pages);
    const ov = applyLayoutOverridesFromUrl(L, url);

    if (debug) {
      return res.status(200).json(buildProbe(P, domSecond, templateInfo, ov, L));
    }

    let templateBytes;
    try {
      const selected = await loadTemplateBytesLocal(selectedTemplate);
      templateBytes = selected.bytes;
      templateInfo.usedTemplate = selectedTemplate;
      templateInfo.selectedTemplatePath = selected.path;
    } catch (errSelected) {
      const fallback = await loadTemplateBytesLocal(fallbackTemplate);
      templateBytes = fallback.bytes;
      templateInfo.usedTemplate = fallbackTemplate;
      templateInfo.usedFallback = true;
      templateInfo.fallbackTemplatePath = fallback.path;
    }

    const pdfDoc = await PDFDocument.load(templateBytes);
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const fontB = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
    const pages = pdfDoc.getPages();

    // Page 1 cover
    if (pages[0]) {
      drawTextBox(pages[0], fontB, P.identity.fullName, L.p1.name, { maxLines: 1 });
      drawTextBox(pages[0], font, P.identity.dateLabel, L.p1.date, { maxLines: 1 });
    }

    // Header name pages 2–15
    const headerName = norm(P.identity.fullName);
    if (headerName) {
      for (let i = 1; i < Math.min(pages.length, 15); i++) {
        const pk = `p${i + 1}`;
        const box = L?.[pk]?.hdrName;
        if (box) drawTextBox(pages[i], font, headerName, box, { maxLines: 1 });
      }
    }

    // Page mapping
    const p4  = pages[3]  || null;
    const p5  = pages[4]  || null;
    const p6  = pages[5]  || null;
    const p7  = pages[6]  || null;
    const p8  = pages[7]  || null;
    const p9  = pages[8]  || null;
    const p10 = pages[9]  || null;
    const p11 = pages[10] || null;
    const p12 = pages[11] || null;
    const p13 = pages[12] || null;
    const p14 = pages[13] || null;
    const p15 = pages[14] || null;

    // Snapshot
    if (p4) {
      drawTextBox(p4, font, P.snapshot_p1, L.twoParaPage.top);
      drawTextBox(p4, font, P.snapshot_p2, L.twoParaPage.bottom);
    }

    if (p5) {
      drawTextBox(p5, font, P.snapshot_p3, L.twoParaPage.top);
      drawTextBox(p5, font, P.snapshot_p4, L.twoParaPage.bottom);
    }

    // Distribution
    if (p6) {
      drawTextBox(p6, font, P.chart_p1, L.p6Distribution.p1);
      drawTextBox(p6, font, P.chart_p2, L.p6Distribution.p2);
      try {
        await embedRadarFromBandsOrUrl(pdfDoc, p6, L.p6Distribution.chart, P.bands || {}, P.chartUrl);
      } catch (e) {
        console.warn("[fill-template:V14.0] Chart skipped:", e?.message || String(e));
      }
    }

    if (p7) {
      drawTextBox(p7, font, P.chart_p3, L.twoParaPage.top);
      drawTextBox(p7, font, P.chart_p4, L.twoParaPage.bottom);
    }

    if (p8) {
      drawTextBox(p8, font, P.chart_p5, L.oneParaPage.body);
    }

    // Movement
    if (p9) {
      drawTextBox(p9, font, P.move_p1, L.twoParaPage.top);
      drawTextBox(p9, font, P.move_p2, L.twoParaPage.bottom);
    }

    if (p10) {
      drawTextBox(p10, font, P.move_p3, L.twoParaPage.top);
      drawTextBox(p10, font, P.move_p4, L.twoParaPage.bottom);
    }

    if (p11) {
      drawTextBox(p11, font, P.move_p5, L.oneParaPage.body);
    }

    // Themes
    if (p12) {
      drawTextBox(p12, font, P.themes_p1, L.twoParaPage.top);
      drawTextBox(p12, font, P.themes_p2, L.twoParaPage.bottom);
    }

    // Interactions
    if (p13) {
      drawTextBox(p13, font, P.interact_c, L.twoParaPage.top);
      drawTextBox(p13, font, P.interact_t, L.twoParaPage.bottom);
    }

    if (p14) {
      drawTextBox(p14, font, P.interact_r, L.twoParaPage.top);
      drawTextBox(p14, font, P.interact_l, L.twoParaPage.bottom);
    }

    // Actions
    if (p15) {
      drawTextBox(p15, font, P.act_1, L.p15Actions.act1);
      drawTextBox(p15, font, P.act_2, L.p15Actions.act2);
      drawTextBox(p15, font, P.act_3, L.p15Actions.act3);
    }

    const outBytes = await pdfDoc.save();
    const outName = makeOutputFilename(P.identity.fullName, P.identity.dateLabel);

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Cache-Control", "no-store");
    res.setHeader("Content-Disposition", `inline; filename="${outName}"`);
    res.status(200).send(Buffer.from(outBytes));
  } catch (err) {
    console.error("[fill-template:V14.0] CRASH", err);
    res.status(500).json({
      ok: false,
      error: err?.message || String(err),
      stack: err?.stack || null,
    });
  }
}
