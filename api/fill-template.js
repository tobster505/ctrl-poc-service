/**
 * CTRL PoC Export Service · fill-template (V17.0 · Preview-only direct PDF return)
 *
 * Purpose:
 * - Accept the existing V17 Build PDF payload contract from Botpress
 * - Generate the filled PDF
 * - Return the completed PDF directly for preview/testing
 * - Keep debug mode (?debug=1)
 * - Keep optional direct PDF download mode (?download=1)
 * - Keep POST as the primary transport
 * - Keep GET ?data=... only as backwards-compatible fallback
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

function normEmail(v) {
  return clean(v).toLowerCase();
}

function looksEmail(v) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normEmail(v));
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

function codeToLabel(c) {
  return ({
    C: "Concealed",
    T: "Triggered",
    R: "Regulated",
    L: "Lead"
  }[String(c || "").toUpperCase()] || "");
}

/* ───────── filename helpers ───────── */
function clampStrForFilename(s) {
  return S(s)
    .trim()
    .replace(/\s+/g, "_")
   
