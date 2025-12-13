/**
 * CTRL PoC Export Service · fill-template (Starter/PoC flow)
 * Place at: /api/fill-template.js  (ctrl-poc-service)
 */
export const config = { runtime: "nodejs" };

/* ───────────── imports ───────────── */
import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";

/* ───────────── utilities ───────────── */
const S = (v, fb = "") => (v == null ? String(fb) : String(v));
const N = (v, fb = 0) => (Number.isFinite(+v) ? +v : fb);
const clamp = (v, lo, hi) => (v < lo ? lo : v > hi ? hi : v);
const norm = (s) =>
  S(s)
    .replace(/\s+/g, " ")
    .trim();

function packSection(tldr, main, action) {
  const blocks = [];
  const T = norm(tldr);
  const M = norm(main);
  const A = norm(action);

  if (T) blocks.push(T); // TLDR first
  if (M) blocks.push(M); // then main
  if (A) blocks.push(A); // then action

  return blocks.filter(Boolean).join("\n\n\n");
}

/* brand colour */
const BRAND = { r: 0.72, g: 0.06, b: 0.44 };

/* ───────── TL→BL rect helper ───────── */
const rectTLtoBL = (page, box, inset = 0) => {
  const pageH = page.getHeight();
  const x = N(box.x) + inset;
  const w = Math.max(0, N(box.w) - inset * 2);
  const h = Math.max(0, N(box.h) - inset * 2);
  const y = pageH - N(box.y) - N(box.h) + inset;
  return { x, y, w, h };
};

/* L-shaped magenta “shadow” helper (kept for future use) */
function drawShadowL(page, absBox, strength = 1) {
  if (!page || !absBox) return;

  const rect = rectTLtoBL(page, absBox, 0);
  const x = rect.x;
  const y = rect.y;
  const w = rect.w;
  const h = rect.h;

  const sideWidth = 18 * strength;
  const baseHeight = 18 * strength;

  page.drawRectangle({
    x,
    y,
    width: w,
    height: baseHeight,
    color: rgb(BRAND.r, BRAND.g, BRAND.b),
  });

  page.drawRectangle({
    x: x + w - sideWidth,
    y,
    width: sideWidth,
    height: h + baseHeight,
    color: rgb(BRAND.r, BRAND.g, BRAND.b),
  });
}

/* ───────── dominant / second state helpers ───────── */
function resolveDomKey(dom, domChar, domDesc) {
  const d = S(dom || "").trim().charAt(0).toUpperCase();
  if (["C", "T", "R", "L"].includes(d)) return d;

  const s = S(domChar || domDesc || "").toLowerCase();
  if (/concealed/.test(s)) return "C";
  if (/triggered/.test(s)) return "T";
  if (/regulated/.test(s)) return "R";
  if (/lead/.test(s)) return "L";

  return "R"; // safe default
}

/**
 * New logic: derive dom + second from:
 * - ctrl.summary.dominant
 * - ctrl.summary.ctrlTotals or .mix
 * - any counts / stateFrequency stored
 */
function computeDomAndSecondKeys(P) {
  const raw = (P && P.raw) || {};
  const ctrl = raw.ctrl || {};
  const summary = ctrl.summary || {};

  // 1) direct
  const domKey = resolveDomKey(P["p3:dom"], raw.domchar, raw.domdesc);

  // 2) try structured totals
  const totals =
    summary.ctrlTotals ||
    summary.totals ||
    summary.mix ||
    ctrl.ctrlTotals ||
    ctrl.totals ||
    ctrl.mix ||
    raw.ctrlTotals ||
    raw.totals ||
    raw.mix ||
    null;

  const fallbackCounts =
    (summary.counts && typeof summary.counts === "object" && summary.counts) ||
    (ctrl.counts && typeof ctrl.counts === "object" && ctrl.counts) ||
    (raw.counts && typeof raw.counts === "object" && raw.counts) ||
    null;

  const score = { C: 0, T: 0, R: 0, L: 0 };

  function addObj(obj) {
    if (!obj || typeof obj !== "object") return;
    for (const k of ["C", "T", "R", "L"]) {
      const v =
        obj[k] ??
        obj[k.toLowerCase()] ??
        obj[{ C: "concealed", T: "triggered", R: "regulated", L: "lead" }[k]];
      if (v != null) score[k] += Number(v) || 0;
    }
  }

  addObj(totals);
  addObj(fallbackCounts);

  // 3) second-best (excluding dom)
  const ordered = ["C", "T", "R", "L"]
    .filter((k) => k !== domKey)
    .map((k) => [k, score[k]])
    .sort((a, b) => b[1] - a[1]);

  const secondKey = ordered[0] ? ordered[0][0] : null;

  return { domKey: domKey || null, secondKey };
}

/* ───────── template + asset loaders ───────── */
async function loadTemplateBytesLocal(fname) {
  if (!fname.endsWith(".pdf"))
    throw new Error(`Invalid template filename: ${fname}`);

  const __file = fileURLToPath(import.meta.url);
  const __dir = path.dirname(__file);

  const candidates = [
    path.join(__dir, "..", "..", "public", fname),
    path.join(__dir, "..", "public", fname),
    path.join(__dir, "public", fname),
    path.join(process.cwd(), "public", fname),
    path.join(__dir, fname),
  ];

  let lastErr;
  for (const pth of candidates) {
    try {
      return await fs.readFile(pth);
    } catch (err) {
      lastErr = err;
    }
  }
  throw new Error(
    `Template not found in any known path for /public: ${fname} (${
      lastErr?.message || "no detail"
    })`
  );
}

async function loadAssetBytes(fname) {
  const __file = fileURLToPath(import.meta.url);
  const __dir = path.dirname(__file);

  const candidates = [
    path.join(__dir, "..", "..", "public", fname),
    path.join(__dir, "..", "public", fname),
    path.join(__dir, "public", fname),
    path.join(process.cwd(), "public", fname),
    path.join(__dir, fname),
  ];

  let lastErr;
  for (const pth of candidates) {
    try {
      return await fs.readFile(pth);
    } catch (err) {
      lastErr = err;
    }
  }
  throw new Error(
    `Asset not found in any known path for /public: ${fname} (${
      lastErr?.message || "no detail"
    })`
  );
}

/* ───────── layout helpers ───────── */
function isPlainObject(obj) {
  return obj && typeof obj === "object" && !Array.isArray(obj);
}

function mergeLayout(base, override) {
  if (!isPlainObject(base)) return base;
  const out = { ...base };
  if (!isPlainObject(override)) return out;

  for (const [k, v] of Object.entries(override)) {
    if (isPlainObject(v) && isPlainObject(out[k])) {
      out[k] = mergeLayout(out[k], v);
    } else {
      out[k] = v;
    }
  }
  return out;
}

/* optional URL overrides, eg Execx, Execy, Execw, Execmaxlines */
function applyQueryLayoutOverrides(L, q) {
  if (!L || !q) return;

  const num = (k, fb) => (q[k] != null ? (Number(q[k]) || fb) : fb);

  if (L.p3 && L.p3.domDesc) {
    const box = L.p3.domDesc;
    L.p3.domDesc = {
      ...box,
      x: num("Execx", box.x),
      y: num("Execy", box.y),
      w: num("Execw", box.w),
      maxLines: num(
        "Execmaxlines",
        typeof box.maxLines === "number" ? box.maxLines : 12
      ),
    };
  }
}

/* ───────── simple helpers ───────── */
const pageOrNull = (pages, idx0) => pages[idx0] || null;

/* ───────── radar chart helpers ───────── */
function makeSpiderChartUrl12(bandsRaw) {
  const labels = [
    "C_low",
    "C_mid",
    "C_high",
    "T_low",
    "T_mid",
    "T_high",
    "R_low",
    "R_mid",
    "R_high",
    "L_low",
    "L_mid",
    "L_high",
  ];

  const vals = labels.map((k) => Number(bandsRaw?.[k] || 0));
  const maxVal = Math.max(...vals, 1);
  const scaled = vals.map((v) => (maxVal > 0 ? v / maxVal : 0));

  const cfg = {
    type: "radar",
    data: {
      labels,
      datasets: [
        {
          label: "",
          data: scaled,
          fill: true,
          borderWidth: 0,
          pointRadius: 0,
        },
      ],
    },
    options: {
      plugins: {
        legend: { display: false },
      },
      scales: {
        r: {
          min: 0,
          max: 1,
          ticks: { display: false },
          grid: { display: false },
          angleLines: { display: false },
          pointLabels: { display: false },
        },
      },
    },
  };

  const enc = encodeURIComponent(JSON.stringify(cfg));
  return `https://quickchart.io/chart?c=${enc}&format=png&width=800&height=800&backgroundColor=transparent`;
}

async function embedRemoteImage(pdfDoc, url) {
  if (!url) return null;
  try {
    const res = await fetch(url);
    if (!res.ok) return null;
    const buf = await res.arrayBuffer();
    const arr = new Uint8Array(buf);
    const sig = String.fromCharCode(arr[0], arr[1], arr[2], arr[3] || 0);
    if (sig.startsWith("\x89PNG")) {
      return await pdfDoc.embedPng(arr);
    }
    if (sig.startsWith("\xff\xd8")) {
      return await pdfDoc.embedJpg(arr);
    }
    try {
      return await pdfDoc.embedPng(arr);
    } catch {
      return await pdfDoc.embedJpg(arr);
    }
  } catch {
    return null;
  }
}

async function embedRadarFromBands(pdfDoc, page, box, bandsRaw) {
  if (!pdfDoc || !page || !box || !bandsRaw) return;

  const hasAny =
    bandsRaw && Object.values(bandsRaw).some((v) => Number(v) > 0);
  if (!hasAny) return;

  const url = makeSpiderChartUrl12(bandsRaw);
  if (!url) return;

  const img = await embedRemoteImage(pdfDoc, url);
  if (!img) return;

  const H = page.getHeight();
  const { x, y, w, h } = box;

  page.drawImage(img, {
    x,
    y: H - y - h,
    width: w,
    height: h,
  });
}

/* generic TL text drawing */
function drawTextBox(page, font, text, box, opts = {}) {
  if (!page || !font || !box) return;
  const raw = S(text || "");
  if (!raw) return;

  const {
    size = box.size || 12,
    align = box.align || "left",
    maxLines = opts.maxLines ?? box.maxLines ?? 99,
    lineGap = box.lineGap ?? 4,
    color = rgb(0, 0, 0),
  } = box;

  const txt = norm(raw);
  if (!txt) return;

  const pageH = page.getHeight();
  const x = N(box.x, 0);
  const yTop = N(box.y, 0);
  const w = N(box.w, 500);

  const words = txt.split(/\s+/);
  const lines = [];
  let current = "";

  const fontSize = size;
  const maxWidth = w;

  for (const word of words) {
    const testLine = current ? current + " " + word : word;
    const testWidth = font.widthOfTextAtSize(testLine, fontSize);
    if (testWidth <= maxWidth || !current) {
      current = testLine;
    } else {
      lines.push(current);
      current = word;
      if (lines.length >= maxLines) break;
    }
  }
  if (current && lines.length < maxLines) lines.push(current);

  const totalHeight = lines.length * fontSize + (lines.length - 1) * lineGap;
  let yStart = pageH - yTop - fontSize;
  if (box.valign === "middle" || box.valign === "center") {
    yStart = pageH - yTop - totalHeight / 2;
  } else if (box.valign === "bottom") {
    yStart = pageH - yTop - totalHeight;
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const lw = font.widthOfTextAtSize(line, fontSize);
    let xx = x;
    if (align === "center") xx = x + (w - lw) / 2;
    if (align === "right") xx = x + (w - lw);

    const yy = yStart - i * (fontSize + lineGap);
    page.drawText(line, { x: xx, y: yy, size: fontSize, font, color });
  }
}

/* ───────── payload readers ───────── */
async function readPayload(req) {
  if (req.method === "POST") {
    const chunks = [];
    for await (const ch of req) chunks.push(ch);
    const buf = Buffer.concat(chunks);
    const raw = buf.toString("utf8") || "{}";
    try {
      return JSON.parse(raw);
    } catch {
      return {};
    }
  }
  // GET: read ?data=<base64json>
  const url = new URL(req.url, "http://localhost");
  const b64 = url.searchParams.get("data") || "";
  if (!b64) return {};
  try {
    const raw = Buffer.from(b64, "base64").toString("utf8");
    return JSON.parse(raw);
  } catch {
    return {};
  }
}

/* ───────── input normaliser ───────── */
function normaliseInput(src = {}) {
  const d = isPlainObject(src) ? src : {};
  const identity = isPlainObject(d.identity) ? d.identity : {};
  const person = isPlainObject(d.person) ? d.person : {};
  const ctrl = isPlainObject(d.ctrl) ? d.ctrl : {};
  const text = isPlainObject(d.text) ? d.text : {};
  const workWith = isPlainObject(d.workWith) ? d.workWith : {};
  const actionsObj = isPlainObject(d.actions) ? d.actions : {};
  const chart = isPlainObject(d.chart) ? d.chart : {};

  const fullName =
    S(
      d.fullName ||
        person.fullName ||
        identity.fullName ||
        identity.name ||
        d.name ||
        "",
      ""
    ).trim() || "";

  const preferredName =
    S(person.preferredName || identity.preferredName || "", "").trim() || "";

  const name = preferredName || fullName || "";

  const dateLbl =
    S(
      d.dateLbl ||
        d.dateLabel ||
        identity.dateLabel ||
        identity.dateLbl ||
        person.date ||
        "",
      ""
    ).trim() || "";

  const chartUrl = S(d.chartUrl || chart.url || "", "").trim() || "";

  const tldrLines =
    Array.isArray(text.tldr) && text.tldr.length
      ? text.tldr
      : Array.isArray(d.tldr) && d.tldr.length
      ? d.tldr
      : [];

  const act1 = S(actionsObj.action1 || actionsObj[0] || "", "").trim();
  const act2 = S(actionsObj.action2 || actionsObj[1] || "", "").trim();

  const out = {
    raw: d,
    name,
    fullName,
    preferredName,
    dateLbl,

    "p1:n": d["p1:n"] || name || "",
    "p1:d": d["p1:d"] || dateLbl || "",

    "p3:dom": d["p3:dom"] || d.dom || ctrl.dom || "",
    "p3:exec": d["p3:exec"] || text.execSummary || "",
    "p3:tldr1": d["p3:tldr1"] || tldrLines[0] || "",
    "p3:tldr2": d["p3:tldr2"] || tldrLines[1] || "",
    "p3:tldr3": d["p3:tldr3"] || tldrLines[2] || "",
    "p3:tldr4": d["p3:tldr4"] || tldrLines[3] || "",
    "p3:tldr5": d["p3:tldr5"] || tldrLines[4] || "",
    "p3:tip": d["p3:tip"] || text.tipAction || "",

    "p4:stateDeep": d["p4:stateDeep"] || text.stateSubInterpretation || "",
    "p4:tldr": d["p4:tldr"] || text.p4tldr || "",
    "p4:action": d["p4:action"] || text.p4action || "",

    "p5:freq": d["p5:freq"] || text.frequency || "",
    "p5:chart": d["p5:chart"] || chartUrl || "",
    "p5:tldr": d["p5:tldr"] || text.p5tldr || "",
    "p5:action": d["p5:action"] || text.p5action || "",

    "p6:seq": d["p6:seq"] || text.sequence || "",
    "p6:tldr": d["p6:tldr"] || text.p6tldr || "",
    "p6:action": d["p6:action"] || text.p6action || "",

    "p7:themesTop": d["p7:themesTop"] || text.themesTop || "",
    "p7:themesLow": d["p7:themesLow"] || text.themesLow || "",

    "p8:collabC": d["p8:collabC"] || workWith.concealed || "",
    "p8:collabT": d["p8:collabT"] || workWith.triggered || "",
    "p8:collabR": d["p8:collabR"] || workWith.regulated || "",
    "p8:collabL": d["p8:collabL"] || workWith.lead || "",

    "p9:action1": act1,
    "p9:action2": act2,
    "p9:closing": d["p9:closing"] || actionsObj.closingNote || "",
  };

  return out;
}

/* ───────── main handler ───────── */
export default async function handler(req, res) {
  try {
    const q = req.method === "POST" ? req.body || {} : req.query || {};
    const src = await readPayload(req);

    console.log("[fill-template] DEBUG_DATE_SRC", {
      dateLbl: src.dateLbl || null,
      dateLabel: src.dateLabel || null,
      identity_dateLabel:
        (src.identity &&
          (src.identity.dateLabel || src.identity.dateLbl)) ||
        null,
    });

    const P = normaliseInput(src);

    console.log("[fill-template] DEBUG_DATE_NORMALISED", {
      P_dateLbl: P.dateLbl || null,
      P_p1d: P["p1:d"] || null,
    });

    // dominant + second → combo (CT, CL, CR, TC, TR, TL, RC, RT, RL, LC, LR, LT)
    const { domKey, secondKey } = computeDomAndSecondKeys(P);

    let combo = "";
    if (domKey && secondKey && secondKey !== domKey) {
      combo = `${domKey}${secondKey}`;
    }

    const validCombos = new Set([
      "CT",
      "CL",
      "CR",
      "TC",
      "TR",
      "TL",
      "RC",
      "RT",
      "RL",
      "LC",
      "LR",
      "LT",
    ]);

    let tplBase;
    if (validCombos.has(combo)) {
      tplBase = `CTRL_PoC_Assessment_Profile_template_${combo}.pdf`;
    } else {
      // last-resort fallback: keep dom, pick a simple neighbour
      const fallbackOrder = ["C", "T", "R", "L"];
      const fallbackSecond =
        secondKey && secondKey !== domKey
          ? secondKey
          : fallbackOrder.find((k) => k !== domKey) || "T";
      const fbCombo =
        domKey && fallbackSecond ? `${domKey}${fallbackSecond}` : "CT";
      tplBase = `CTRL_PoC_Assessment_Profile_template_${fbCombo}.pdf`;
    }

    console.log("[fill-template] TEMPLATE_SELECTED", {
      domKey,
      secondKey,
      combo,
      tplBase,
    });

    const tpl = S(tplBase).replace(/[^A-Za-z0-9._-]/g, "");
    const pdfBytes = await loadTemplateBytesLocal(tpl);
    const pdfDoc = await PDFDocument.load(pdfBytes);
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);

    const pages = pdfDoc.getPages();
    const p1 = pageOrNull(pages, 0);
    const p2 = pageOrNull(pages, 1);
    const p3 = pageOrNull(pages, 2);
    const p4 = pageOrNull(pages, 3);
    const p5 = pageOrNull(pages, 4);
    const p6 = pageOrNull(pages, 5);
    const p7 = pageOrNull(pages, 6);
    const p8 = pageOrNull(pages, 7);
    const p9 = pageOrNull(pages, 8);
    const p10 = pageOrNull(pages, 9);
    const p11 = pageOrNull(pages, 10);
    const p12 = pageOrNull(pages, 11);

    /* base layout (no boxes/images on p3) */
    let L = {
      p1: {
        name: { x: 7, y: 473, w: 500, size: 30, align: "center" },
        date: { x: 210, y: 600, w: 500, size: 25, align: "left" },
      },
      p3: {
        // single big text block for Exec + TLDRs + Tip
        domDesc: {
          x: 25,
          y: 685,
          w: 550,
          size: 18,
          align: "left",
          maxLines: 20,
        },
      },
      p4: {
        spider: {
          x: 25,
          y: 347,
          w: 550,
          size: 18,
          align: "left",
          maxLines: 20,
        },
      },
      p5: {
        seqpat: {
          x: 25,
          y: 347,
          w: 550,
          size: 18,
          align: "left",
          maxLines: 20,
        },
        chart: { x: 48, y: 462, w: 500, h: 300 },
      },
      p6: {
        themeExpl: {
          x: 25,
          y: 347,
          w: 550,
          size: 18,
          align: "left",
          maxLines: 20,
        },
      },
      p7: {
        colBoxes: [
          { x: 25, y: 330, w: 260, h: 420 },
          { x: 320, y: 330, w: 260, h: 420 },
        ],
        bodySize: 13,
        maxLines: 22,
      },
      p8: {
        colBoxes: [
          { x: 25, y: 330, w: 260, h: 120 },
          { x: 320, y: 330, w: 260, h: 120 },
          { x: 25, y: 595, w: 260, h: 120 },
          { x: 320, y: 595, w: 260, h: 120 },
        ],
        bodySize: 13,
        maxLines: 15,
      },
      p9: {
        tips1: { x: 30, y: 280, w: 550, size: 18, align: "left", maxLines: 4 },
        tips2: { x: 30, y: 340, w: 550, size: 18, align: "left", maxLines: 4 },
        acts1: { x: 30, y: 120, w: 550, size: 18, align: "left", maxLines: 4 },
        acts2: { x: 30, y: 180, w: 550, size: 18, align: "left", maxLines: 4 },
      },
    };

    // layout override via payload
    if (src && src.layout && isPlainObject(src.layout)) {
      L = mergeLayout(L, src.layout);
    }

    // optional URL overrides
    applyQueryLayoutOverrides(L, q);

    /* p1: name + date */
    if (p1 && L.p1) {
      if (P["p1:n"]) {
        drawTextBox(p1, font, P["p1:n"], L.p1.name, { maxLines: 1 });
      }
      if (P["p1:d"]) {
        drawTextBox(p1, font, P["p1:d"], L.p1.date, { maxLines: 1 });
      }
    }

    /* p3: Exec + TLDRs + Tip in one block */
    if (p3 && L.p3 && L.p3.domDesc) {
      const exec = norm(P["p3:exec"]);
      const tldrs = [
        norm(P["p3:tldr1"]),
        norm(P["p3:tldr2"]),
        norm(P["p3:tldr3"]),
        norm(P["p3:tldr4"]),
        norm(P["p3:tldr5"]),
      ].filter(Boolean);
      const tip = norm(P["p3:tip"]);

      const blocks = [];
      if (tldrs.length) blocks.push(tldrs.join("\n\n")); // TLDR FIRST
      if (exec) blocks.push(exec);
      if (tip) blocks.push(tip);

      const body = blocks.filter(Boolean).join("\n\n\n");
      if (body) {
        drawTextBox(p3, font, body, L.p3.domDesc, {
          maxLines: L.p3.domDesc.maxLines,
        });
      }
    }

    /* p4: deep dive narrative */
    if (p4 && L.p4 && L.p4.spider) {
      const body = packSection(P["p4:tldr"], P["p4:stateDeep"], P["p4:action"]);
      if (body) {
        drawTextBox(p4, font, body, L.p4.spider, {
          maxLines: L.p4.spider.maxLines,
        });
      }
    }

    /* p5: frequency narrative + radar chart */
    if (p5 && L.p5) {
      if (L.p5.seqpat) {
        const body = packSection(P["p5:tldr"], P["p5:freq"], P["p5:action"]);
        if (body) {
          drawTextBox(p5, font, body, L.p5.seqpat, {
            maxLines: L.p5.seqpat.maxLines,
          });
        }
      }

      if (L.p5.chart) {
        const bands =
          P.bands ||
          (P.raw &&
            P.raw.ctrl &&
            (P.raw.ctrl.bands ||
              (P.raw.ctrl.summary && P.raw.ctrl.summary.bands))) ||
          (P.raw && P.raw.bands) ||
          {};

        const H = p5.getHeight();
        const { x, y, w, h } = L.p5.chart;

        // white backing to clear any template artefact behind the image
        p5.drawRectangle({
          x,
          y: H - y - h,
          width: w,
          height: h,
          color: rgb(1, 1, 1),
        });

        await embedRadarFromBands(pdfDoc, p5, L.p5.chart, bands);
      }
    }

    /* p6: sequence */
    if (p6 && L.p6 && L.p6.themeExpl) {
      const body = packSection(P["p6:tldr"], P["p6:seq"], P["p6:action"]);
      if (body) {
        drawTextBox(p6, font, body, L.p6.themeExpl, {
          maxLines: L.p6.themeExpl.maxLines,
        });
      }
    }

    /* p7: themes top + low */
    if (p7 && Array.isArray(L.p7?.colBoxes) && L.p7.colBoxes.length >= 2) {
      const a = L.p7.colBoxes[0];
      const b = L.p7.colBoxes[1];

      const txtA = norm(P["p7:themesTop"]);
      const txtB = norm(P["p7:themesLow"]);

      if (txtA) {
        drawTextBox(
          p7,
          font,
          txtA,
          { x: a.x, y: a.y, w: a.w, size: L.p7.bodySize || 13, align: "left" },
          { maxLines: L.p7.maxLines || 22 }
        );
      }

      if (txtB) {
        drawTextBox(
          p7,
          font,
          txtB,
          { x: b.x, y: b.y, w: b.w, size: L.p7.bodySize || 13, align: "left" },
          { maxLines: L.p7.maxLines || 22 }
        );
      }
    }

    /* p8: work-with text C/T/R/L */
    if (p8 && Array.isArray(L.p8?.colBoxes) && L.p8.colBoxes.length >= 4) {
      const mapIdx = { C: 0, T: 1, R: 2, L: 3 };
      const txtByState = {
        C: norm(P["p8:collabC"]),
        T: norm(P["p8:collabT"]),
        R: norm(P["p8:collabR"]),
        L: norm(P["p8:collabL"]),
      };

      ["C", "T", "R", "L"].forEach((key) => {
        const txt = txtByState[key];
        if (!txt) return;
        const idx = mapIdx[key];
        const box = L.p8.colBoxes[idx];

        drawTextBox(
          p8,
          font,
          txt,
          { x: box.x, y: box.y, w: box.w, size: L.p8.bodySize || 13, align: "left" },
          { maxLines: L.p8.maxLines || 15 }
        );
      });
    }

    /* p9: actions + closing note */
    if (p9 && L.p9) {
      const tidy = (s) =>
        norm(String(s || ""))
          .replace(/^(?:[-–—•·]\s*)/i, "")
          .replace(/^\s*(tips?|tip)\s*:?\s*/i, "")
          .replace(/^\s*(actions?|next\s*action)\s*:?\s*/i, "")
          .trim();
      const good = (s) =>
        s && s.length >= 3 && !/^tips?$|^actions?$/i.test(s);

      const actionsPacked = [tidy(P["p9:action1"]), tidy(P["p9:action2"])]
        .filter(good)
        .slice(0, 2);

      const closing = tidy(P["p9:closing"]);

      const drawBullet = (page, spec, text) => {
        if (!page || !spec || !text) return;
        const bullet = `• ${text}`;
        drawTextBox(page, font, bullet, spec, {
          maxLines: spec.maxLines || 4,
        });
      };

      const slots = {
        tips1: L.p9.tips1,
        tips2: L.p9.tips2,
        acts1: L.p9.acts1,
        acts2: L.p9.acts2,
      };

      if (actionsPacked[0]) drawBullet(p9, slots.acts1, actionsPacked[0]);
      if (actionsPacked[1]) drawBullet(p9, slots.acts2, actionsPacked[1]);
      if (closing) drawBullet(p9, slots.tips1, closing);
    }

    /* footer name p2–p12 */
    const footerLabel = norm(P.name);
    const putFooter = (page) => {
      if (!page || !footerLabel) return;
      drawTextBox(
        page,
        font,
        footerLabel,
        { x: 380, y: 51, w: 400, size: 13, align: "left" },
        { maxLines: 1 }
      );
    };
    [p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12].forEach(putFooter);

    const bytes = await pdfDoc.save();

    const safe = (value, fallback = "") =>
      String(value || fallback)
        .trim()
        .replace(/[^A-Za-z0-9]+/g, "_")
        .replace(/^_+|_+$/g, "");

    const namePart = safe(P.name || P.fullName || "Profile");
    const datePart = safe(P.dateLbl || P.dateLabel || P.date || "");

    const fileName = datePart
      ? `PoC_Profile_${namePart}_${datePart}.pdf`
      : `PoC_Profile_${namePart}.pdf`;

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `inline; filename="${fileName}"`);
    res.send(Buffer.from(bytes));
  } catch (err) {
    console.error("PDF handler error:", err);
    res.status(500).json({
      error: "Failed to generate PDF",
      detail: err?.message || String(err),
    });
  }
}
