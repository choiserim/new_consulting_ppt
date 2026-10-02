"use strict";
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const JSZip = require("jszip");
const PptxGenJS = require("pptxgenjs");
const B = require("../src/core.js");

const FIX = path.join(__dirname, "fixtures");
const FIXED_DATE = new Date("2026-10-02T09:00:00+09:00"); // 결과 비교를 위해 날짜 고정

function bufOf(file) {
  const b = fs.readFileSync(path.join(FIX, file));
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.length);
}

async function rowsOf(file) {
  if (/\.pdf$/i.test(file)) {
    const pdfjs = require("pdfjs-dist/legacy/build/pdf.js");
    return B.rowsFromPdf(pdfjs, bufOf(file));
  }
  return B.rowsFromWorkbook(XLSX, bufOf(file));
}

async function analyzeFile(file) {
  const rows = await rowsOf(file);
  const d = B.parseReport(rows);
  const A = B.analyze(d, FIXED_DATE);
  return { rows, d, A, queryDate: B.queryDateOf(rows) };
}

/** PPT를 만들고 슬라이드별 글자·차트 데이터를 뽑는다 */
async function deckOf(file, author = {}) {
  const r = await analyzeFile(file);
  const pres = B.buildDeck(PptxGenJS, r.d, r.A, Object.assign({ date: FIXED_DATE, queryDate: r.queryDate }, author));
  const buf = await pres.write({ outputType: "nodebuffer" });
  const zip = await JSZip.loadAsync(buf);
  const slideNames = Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => +a.match(/(\d+)\.xml/)[1] - +b.match(/(\d+)\.xml/)[1]);
  const slides = [];
  for (const n of slideNames) {
    const xml = await zip.file(n).async("string");
    const texts = [...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => decode(m[1]));
    const rels = await zip.file(n.replace("slides/", "slides/_rels/") + ".rels").async("string");
    const charts = [];
    for (const m of rels.matchAll(/Target="(?:\.\.|\/ppt)\/charts\/(chart\d+\.xml)"/g)) {
      const cx = await zip.file("ppt/charts/" + m[1]).async("string");
      for (const ser of cx.split("<c:ser>").slice(1)) {
        const name = (ser.match(/<c:tx>[\s\S]*?<c:v>([^<]*)<\/c:v>/) || [])[1];
        const cat = (ser.match(/<c:cat>([\s\S]*?)<\/c:cat>/) || [])[1] || "";
        const val = (ser.match(/<c:val>([\s\S]*?)<\/c:val>/) || [])[1] || "";
        charts.push({ name: decode(name || ""), cats: [...cat.matchAll(/<c:v>([^<]*)<\/c:v>/g)].map((x) => decode(x[1])), vals: [...val.matchAll(/<c:v>([^<]*)<\/c:v>/g)].map((x) => +x[1]) });
      }
    }
    slides.push({ text: texts.join("\n"), charts });
  }
  const layouts = [];
  for (const n of Object.keys(zip.files).filter((x) => /slideLayouts\/slideLayout\d+\.xml$/.test(x))) layouts.push(await zip.file(n).async("string"));
  return Object.assign(r, { slides, layoutsXml: layouts.join("\n"), buffer: buf }); // pptxgenjs 객체는 한 번만 저장할 수 있어 버퍼를 돌려준다
}

function decode(s) {
  return s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
}

module.exports = { B, FIXED_DATE, analyzeFile, deckOf };
