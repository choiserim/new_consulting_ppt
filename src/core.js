/* PlanA 기업 컨설팅 PPT 생성 엔진
 * 한국평가데이터(KoDATA) 기업종합보고서(xls/xlsx/pdf) → 파싱 → 7대 영역 진단 → PPTX
 * 브라우저: window.BizReport / Node: module.exports
 */
(function (root) {
  "use strict";

  // ───────────────────────── 공통 유틸 ─────────────────────────
  const MAX_YEARS = 3; // 재무 표시 기본 연수
  const NUM_RE = /^[-+]?(\d+(\.\d+)?|\.\d+)$/;
  function num(s) {
    if (s === null || s === undefined) return null;
    if (typeof s === "number") return isFinite(s) ? s : null;
    let t = String(s).replace(/[,\s]/g, "").replace(/(백만원|천원|원|건|명|%|회|배|위)$/, "");
    if (t === "-" || t === "") return null;
    let neg = false;
    if (/^[△▽]/.test(t)) { neg = true; t = t.slice(1); }
    else if (/^\(.*\)$/.test(t)) { neg = true; t = t.slice(1, -1); }
    if (!NUM_RE.test(t)) return NaN;
    const v = parseFloat(t);
    return neg ? -Math.abs(v) : v;
  }
  const isNumCell = (s) => { const n = num(s); return n === null ? String(s).trim() === "-" : !isNaN(n); };
  const normLabel = (s) => String(s).replace(/^-\s*/, "").replace(/\s*:\s*$/, "").trim();
  const fmt = (v, dp = 0) => (v === null || v === undefined || isNaN(v)) ? "-" :
    Number(v).toLocaleString("ko-KR", { minimumFractionDigits: dp, maximumFractionDigits: dp });
  const last = (a) => (a && a.length ? a[a.length - 1] : null);
  const prev = (a) => (a && a.length > 1 ? a[a.length - 2] : null);
  const pct = (a, b) => (a && b !== null && a !== null && a !== 0 ? ((b - a) / Math.abs(a)) * 100 : null);

  function cleanRows(rows) {
    const out = [];
    for (const r of rows) {
      const cells = r.map((c) => (c === null || c === undefined ? "" : String(c).replace(/\s*\n\s*/g, " ").trim()))
        .filter((c) => c !== "");
      if (!cells.length) continue;
      if (/^COPYRIGHT/i.test(cells[0]) || /^\d+\s*\/\s*\d+$/.test(cells[0])) continue;
      const kept = cells.filter((c) => !/^COPYRIGHT/i.test(c) && !/^\d+\s*\/\s*\d+$/.test(c));
      if (kept.length) out.push(kept);
    }
    return out;
  }

  // ───────────────────────── 파일 → 행 ─────────────────────────
  function rowsFromWorkbook(XLSX, arrayBuffer) {
    const wb = XLSX.read(arrayBuffer, { type: "array" });
    const rows = [];
    for (const name of wb.SheetNames) {
      const ws = wb.Sheets[name];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
      for (const r of data) rows.push(r.map((v) => (typeof v === "number" ? String(v) : v)));
    }
    return cleanRows(rows);
  }

  async function rowsFromPdf(pdfjsLib, arrayBuffer) {
    const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(arrayBuffer), isEvalSupported: false }).promise;
    const rows = [];
    for (let p = 1; p <= pdf.numPages; p++) {
      const page = await pdf.getPage(p);
      const tc = await page.getTextContent();
      const items = tc.items.filter((i) => i.str && i.str.trim()).map((i) => ({
        s: i.str, x: i.transform[4], y: i.transform[5], w: i.width || 0, h: Math.abs(i.transform[3]) || Math.abs(i.transform[0]) || 8,
      }));
      items.sort((a, b) => b.y - a.y || a.x - b.x);
      const lines = [];
      for (const it of items) {
        const ln = lines.length ? lines[lines.length - 1] : null;
        if (ln && Math.abs(ln.y - it.y) <= Math.max(2, it.h * 0.45)) ln.items.push(it);
        else lines.push({ y: it.y, items: [it] });
      }
      for (const ln of lines) {
        ln.items.sort((a, b) => a.x - b.x);
        const cells = [];
        let cur = null;
        for (const it of ln.items) {
          if (cur) {
            const gap = it.x - (cur.x + cur.w);
            if (gap < it.h * 1.2) {
              cur.s += (gap > it.h * 0.2 ? " " : "") + it.s;
              cur.w = it.x + it.w - cur.x;
              continue;
            }
            cells.push(cur.s);
          }
          cur = { s: it.s, x: it.x, w: it.w };
        }
        if (cur) cells.push(cur.s);
        rows.push(cells);
      }
    }
    return mergeWrapped(cleanRows(rows));
  }

  // PDF 표에서 칸 안 글자가 두 줄로 나뉜 경우(예: "총자산순이익률" / 숫자 줄 / "(ROA)") 한 행으로 합친다
  function mergeWrapped(rows) {
    const out = [];
    const labelOnly = (r) => r && r.length <= 2 && r.every((c) => !isNumCell(c)) && (r.length === 1 || r[0] === r[1]);
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i], nx = rows[i + 1], nx2 = rows[i + 2];
      if (labelOnly(r) && r[0].length <= 20 && nx && nx.length >= 2 && isNumCell(nx[0])) {
        let label = r[0];
        let skip = 1;
        if (labelOnly(nx2) && /^\(.*\)$/.test(nx2[0])) { label += nx2[0]; skip = 2; }
        out.push([label, ...nx]);
        i += skip;
        continue;
      }
      out.push(r);
    }
    return out;
  }

  // ───────────────────────── 보고서 파싱 ─────────────────────────
  const INFO_LABELS = ["기업명", "영문기업명", "사업자번호", "대표자명", "종업원수", "설립형태", "설립년월", "기업유형", "기업규모",
    "전화번호", "팩스번호", "홈페이지", "이메일", "결산월", "기업공개일자", "주소", "표준산업분류(10차)", "표준산업분류(11차)",
    "주요제품(상품)", "무역업허가번호", "소속그룹", "주채권기관", "당좌거래은행", "휴폐업정보", "법인등기정보", "법인(주민)번호",
    "자가소유 여부", "실제가동 여부", "담보제공여부", "권리침해여부", "입지조건", "산업단지명", "대지", "건물", "소유자",
    "소유주와의 관계", "임차보증금규모", "월세금액규모", "주생산품", "주사업장여부", "사업장명", "구분"];
  const LSET = new Set(INFO_LABELS);
  const STMT = { "재무상태표": "BS", "손익계산서": "IS", "제조원가명세서": "MC", "공사원가명세서": "CC", "이익잉여금처분계산서": "RE", "결손금처리계산서": "RE", "현금흐름표": "CF", "자본변동표": "EQ" };
  const SUMMARY = { "요약 재무상태표": "sBS", "요약 손익계산서": "sIS", "요약 현금흐름분석": "sCF" };
  const OTHER = { "요약 재무비율": "ratio", "재무비율": "ratio", "재무진단": "diag", "연혁": "hist", "사업목적": "purp", "종합의견": "opinion",
    "인적사항": "person", "경영진현황": "exec", "주요주주현황": "holder", "관계회사현황": "related", "사업장 현황": "site", "거래처현황": "trade",
    "구매처현황": "buy", "판매처현황": "sell", "매출구성": "mix", "업계순위": "rank", "동종업계내 매출액 분포": "dist",
    "동종업계내 경영규모 비교": "peer", "MY 재무 Data": "my", "기술력": "tech", "신용정보": "credit", "기업개요": "overview" };
  const DIAG5 = ["성장성", "수익성", "재무구조", "부채상환능력", "활동성"];
  const BIZNO = /^\d{3}-\d{2}-\d{5}$/;

  function parseReport(rows) {
    const d = { info: {}, years: [], sumYears: [], fs: {}, ratio: {}, ind: {}, grade: {}, opinion: {}, peer: {}, certs: {}, ip: {},
      holders: [], execs: [], buyers: [], sellers: [], related: [], history: [], purposes: [], credit: {}, bid: null, rank: null };

    // 1) 기본 정보(라벨 | 값)
    for (const r of rows) {
      for (let i = 0; i < r.length - 1; i++) {
        const k = normLabel(r[i]);
        if (LSET.has(k) && k !== "구분" && d.info[k] === undefined) {
          const v = r[i + 1];
          d.info[k] = LSET.has(normLabel(v)) ? "" : v;
        }
      }
    }
    if (!d.info["기업명"]) {
      const r = rows.find((x) => /기업명/.test(x[0]) && x[1]);
      if (r) d.info["기업명"] = r[1];
    }
    const company = d.info["기업명"] || "";
    for (const r of rows) { const m = r.join(" ").match(/결산일자\s*:\s*(\d{4})/); if (m) { d.closingYear = m[1]; break; } }

    // 2) 섹션 순회
    let sec = null, diagCur = null, peerKeys = null;
    for (let idx = 0; idx < rows.length; idx++) {
      const r = rows[idx];
      const h = r[0];
      // 섹션과 무관하게 찾는 항목(PDF에서 머리글과 같은 줄에 붙어 나와도 읽도록 먼저 확인)
      if (r.includes("벤처") && r.includes("이노비즈")) {
        const nx = rows[idx + 1] || [];
        r.forEach((k, i) => { if (nx[i] !== undefined) d.certs[k] = nx[i]; });
      }
      if (r[0] === "특허" && r.includes("상표권")) {
        for (let j = 1; j <= 3; j++) {
          const nx = rows[idx + j] || [];
          const vals = nx.filter((c) => /^\d+\s*건$|^-$/.test(c));
          if (vals.length >= r.length) { const v4 = vals.slice(-r.length); r.forEach((k, i) => (d.ip[k] = num(v4[i]) || 0)); break; }
        }
      }
      if (r.includes("입찰건수") && r.includes("낙찰건수")) {
        const nx = rows[idx + 1] || [];
        d.bid = Object.assign(d.bid || {}, { bids: num(nx[0]), wins: num(nx[1]) });
      }
      if (r[0] === "총 낙찰금액") { const nx = rows[idx + 1] || []; d.bid = Object.assign(d.bid || {}, { amount: num(nx[0]) }); }
      if (r.some((c) => /단기연체정보/.test(c)) || (r[0] === "휴폐업정보" && r.length >= 4)) {
        const nx = rows[idx + 1] || [];
        r.forEach((k, i) => { if (nx[i] !== undefined) d.credit[k] = nx[i]; });
      }
      { const ri = r.findIndex((c, i) => /^\d{1,6}위$/.test(c) && r[i + 1] === company); if (ri >= 0 && !d.rank) d.rank = parseInt(r[ri], 10); }

      // 섹션 시작
      if (STMT[h] && !r.some((c) => /백만원/.test(c))) { sec = STMT[h]; continue; }
      if (SUMMARY[h]) { sec = SUMMARY[h]; continue; }
      if (OTHER[h] && !(h === "재무진단" && r.length > 3)) { sec = OTHER[h]; diagCur = null; continue; }


      if (!sec || /^조회일시/.test(h)) continue;

      if (sec === "BS" || sec === "IS" || sec === "MC" || sec === "CC" || sec === "RE" || sec === "CF") {
        if (h === "계정명") {
          const ys = r.slice(1).filter((c) => /^\d{4}-\d{2}-\d{2}$/.test(c)).map((c) => c.slice(0, 4));
          if (ys.length > d.years.length) d.years = ys;
          continue;
        }
        if (h === "감사의견" || /조회된 자료/.test(h)) continue;
        let k = r.length; while (k > 0 && isNumCell(r[k - 1])) k--;
        if (k === 0 || k === r.length) continue;
        const label = r[k - 1];
        const vals = r.slice(k).map((c) => (String(c).trim() === "-" ? 0 : num(c)));
        d.fs[sec] = d.fs[sec] || {};
        if (!(label in d.fs[sec])) d.fs[sec][label] = vals;
        continue;
      }
      if (sec === "sBS" || sec === "sIS" || sec === "sCF") {
        if (h === "구분") { const ys = r.slice(1).filter((c) => /^\d{4}$/.test(c)); if (ys.length > d.sumYears.length) d.sumYears = ys; continue; }
        let k = r.length; while (k > 0 && isNumCell(r[k - 1])) k--;
        if (k === 0 || k === r.length) continue;
        d.fs[sec] = d.fs[sec] || {};
        d.fs[sec][r[k - 1]] = r.slice(k).map((c) => num(c));
        continue;
      }
      if (sec === "ratio") {
        let label = null, vals = [];
        const flush = () => { if (label && vals.length) d.ratio[label] = vals; };
        for (const c of r) {
          if (isNumCell(c)) { if (label) vals.push(num(c)); }
          else { flush(); label = c; vals = []; }
        }
        flush();
        continue;
      }
      if (sec === "diag") {
        if (DIAG5.includes(h) && r[1] && /^기준일자/.test(r[1])) { diagCur = h; continue; }
        if (DIAG5.includes(h) && r[1] && r[1].length > 20) { d.opinion[h] = r[1]; continue; }
        if (r.length === 1 && diagCur && /(함|없음|부족)$/.test(h)) { d.grade[diagCur] = h; diagCur = null; continue; }
        if (r.length >= 4 && isNumCell(r[1]) && /증가|감소|^-$/.test(r[2])) d.ind[h] = { avg: num(r[1]), cur: num(r[3]) };
        continue;
      }
      if (sec === "peer") {
        if (h === "구분") { peerKeys = r.slice(1); continue; }
        if (["조회기업", "상위25%", "평균", "하위25%"].includes(h)) {
          const keys = peerKeys || ["총자산", "자본총계", "납입자본금", "매출액", "영업이익", "당기순이익"];
          d.peer[h] = {}; keys.forEach((k, i) => (d.peer[h][k] = num(r[i + 1])));
        }
        continue;
      }
      if (sec === "holder") {
        if (r.includes("주주명") || r[0] === "보통주" || /기준일자/.test(r.join(" "))) continue;
        const nums = r.slice(1).map(num).filter((v) => v !== null && !isNaN(v));
        const pctv = [...nums].reverse().find((v) => v <= 100);
        const rel = r.find((c) => ["본인", "가족", "타인", "기타", "임원", "친인척"].includes(c)) || "";
        const shares = nums.find((v) => v > 100) || null;
        if (pctv && pctv > 0 && r[0] !== "기타") d.holders.push({ name: r[0], pct: pctv, rel, shares });
        continue;
      }
      if (sec === "exec") {
        if (r.includes("성명") || r[0] === "주식") continue;
        const posRe = /^(대표이사|사내이사|사외이사|감사|기타비상무이사|이사)$/;
        if (!/(대표|이사|감사|사원)/.test(h) || !r[1]) continue;
        const pos = r.slice(2).find((c) => posRe.test(c)) || h.replace("및사원", "");
        const rel = r.find((c) => ["본인", "가족", "타인", "친인척"].includes(c)) || "";
        d.execs.push({ name: r[1], pos, rel, reg: r.includes("등기") });
        continue;
      }
      if (sec === "related") {
        if (r[0] === "기업명" || /해당사항/.test(r[0]) || /기준일자/.test(r.join(" "))) continue;
        d.related.push({ name: r[0], biz: r[1] || "", rel: r[2] || "" });
        continue;
      }
      if (sec === "buy" || sec === "sell") {
        const list = sec === "buy" ? d.buyers : d.sellers;
        if (r[0] === "기업명" || /기준일자|단위/.test(r.join(" "))) continue;
        const bi = r.findIndex((c) => BIZNO.test(c));
        if (bi === 1) {
          const rep = r[2] && isNaN(num(r[2])) === true && !isNumCell(r[2]) ? r[2] : "";
          const share = r.slice(2).map(num).find((v) => v !== null && !isNaN(v));
          list.push({ name: r[0], bizno: r[1], rep, share: share ?? null });
        } else if (r[0] === "기타" && r[1]) list.push({ name: "기타", bizno: "", rep: "", share: num(r[1]) });
        continue;
      }
      if (sec === "hist") {
        if (/^\d{4}-\d{2}(-\d{2})?$/.test(h) && r[1]) d.history.push({ date: h, text: r.slice(1).join(" ") });
        continue;
      }
      if (sec === "purp") {
        const m = h.match(/^(\d+)\.?\s*(.*)$/);
        if (m && m[2]) d.purposes.push({ no: parseInt(m[1], 10), text: m[2] });
        continue;
      }
    }
    return d;
  }

  // ───────────────────────── 분석 ─────────────────────────
  function analyze(d, now = new Date()) {
    // ── 표시 연도 결정: 기본 3개년, 설립 연도에 따라 2개년·1개년
    const allYears = d.years.length ? d.years : d.sumYears;
    const estM = String(d.info["설립년월"] || "").match(/(\d{4})/);
    const estY0 = estM ? +estM[1] : null;
    const closingY = +(d.closingYear || last(allYears) || 0) || null;
    const byAge = estY0 && closingY ? Math.max(1, Math.min(MAX_YEARS, closingY - estY0 + 1)) : MAX_YEARS;
    let n = allYears.length ? Math.min(byAge, allYears.length) : MAX_YEARS;
    const align = (vals, div) => {
      if (!vals) return null;
      const v = vals.slice(-n).map((x) => (x === null || isNaN(x) ? null : x / div));
      while (v.length < n) v.unshift(null);
      return v;
    };
    const g = (sec, names, sumSec, sumNames) => {
      const tbl = d.fs[sec];
      if (tbl) for (const nm of names) for (const k of [nm, nm + "(*)"]) if (tbl[k]) return align(tbl[k], 1000);
      if (sumSec && d.fs[sumSec]) for (const nm of sumNames || names) if (d.fs[sumSec][nm]) return align(d.fs[sumSec][nm], 1);
      return null;
    };
    const add = (...arrs) => {
      const a = arrs.filter(Boolean); if (!a.length) return null;
      return Array.from({ length: n }, (_, i) => { const xs = a.map((x) => x[i]).filter((x) => x !== null); return xs.length ? xs.reduce((s, x) => s + x, 0) : null; });
    };
    const r = (name) => (d.ratio[name] ? align(d.ratio[name], 1) : null);
    // 설립 전·자료 없는 앞쪽 연도(매출·자산 모두 0 또는 없음) 제외
    const emptyLead = () => {
      const sv = g("IS", ["매출액"], "sIS", ["매출액"]) || [], av = g("BS", ["자산", "자산총계"], "sBS", ["자산총계"]) || [];
      return !sv[0] && !av[0];
    };
    while (allYears.length && n > 1 && emptyLead()) n--;
    const years = allYears.slice(allYears.length - n);

    const F = {
      sales: g("IS", ["매출액"], "sIS", ["매출액"]),
      cost: g("IS", ["매출원가"]),
      sga: g("IS", ["판매비와관리비"]),
      op: g("IS", ["영업이익(손실)", "영업이익"], "sIS", ["영업이익"]),
      ni: g("IS", ["당기순이익(순손실)", "당기순이익"], "sIS", ["당기순이익"]),
      pretax: g("IS", ["법인세비용차감전순손익", "법인세비용차감전순이익"]),
      tax: g("IS", ["법인세비용", "법인세등"]),
      assets: g("BS", ["자산", "자산총계"], "sBS", ["자산총계"]),
      liab: g("BS", ["부채", "부채총계"], "sBS", ["부채총계"]),
      equity: g("BS", ["자본", "자본총계"], "sBS", ["자본총계"]),
      capital: g("BS", ["자본금"], "sBS", ["자본금"]),
      retained: g("BS", ["이익잉여금", "이익잉여금(결손금)"]),
      cash: g("BS", ["현금 및 현금성자산", "현금및현금성자산"]),
      karyo: g("BS", ["가지급금", "주임종단기채권", "주.임.종단기채권", "주임종단기대여금"]),
      vehicles: g("BS", ["차량운반구"]),
      borrow: add(g("BS", ["단기차입금"]), g("BS", ["장기차입금"]), g("BS", ["유동성장기부채"]), g("BS", ["사채"])),
      salary: g("IS", ["급여"]),
      mcLabor: g("MC", ["노동관계비용", "노무비"]),
      retire: add(g("IS", ["퇴직급여"]), g("MC", ["퇴직급여"])),
      insurance: add(g("IS", ["보험료"]), g("MC", ["보험료"])),
      entertain: g("IS", ["접대비", "기업업무추진비"]),
      rnd: add(g("IS", ["경상개발비", "연구개발비"]), g("MC", ["경상개발비", "연구개발비"])),
      construction: g("IS", ["공사수입"]),
      rent: add(g("IS", ["임차료"]), g("MC", ["임차료"])),
      donation: g("IS", ["기부금"]),
      pyaRE: g("RE", ["전기오류수정이익"]),
      pyaLoss: g("IS", ["전기오류수정손실"]),
    };
    F.labor = add(F.salary, F.mcLabor, F.retire);
    const pyaArr = Array.from({ length: n }, (_, i) => {
      return F.pyaRE && F.pyaRE[i] !== null && F.pyaRE[i] < 0 ? -F.pyaRE[i] : 0;
    });
    F.pya = pyaArr;
    const pyaTotal = pyaArr.reduce((s, x) => s + x, 0);

    const ACC = [
      ["현금및현금성자산", ["현금 및 현금성자산", "현금및현금성자산"]], ["단기예금", ["단기예금(단기금융상품)", "단기금융상품", "단기예금"]],
      ["매출채권", ["매출채권"]], ["미수금", ["미수금"]], ["대손충당금", ["(대손충당금)"]], ["선급금", ["선급금"]],
      ["가지급금", ["가지급금", "주임종단기채권", "주.임.종단기채권"]], ["단기대여금", ["단기대여금"]], ["재고자산", ["재고자산"]],
      ["보증금", ["보증금 등", "보증금", "임차보증금"]], ["매입채무", ["매입채무"]], ["단기차입금", ["단기차입금"]],
      ["장기차입금", ["장기차입금"]], ["미지급법인세", ["미지급법인세"]],
    ];
    const accounts = ACC.map(([label, names]) => ({ label, v: g("BS", names) }))
      .filter((a) => a.v && a.v.some((x) => x && Math.abs(x) >= 0.5));

    const info = d.info;
    const emp = (() => { const m = String(info["종업원수"] || "").match(/(\d+)/); return m ? parseInt(m[1], 10) : null; })();
    const est = String(info["설립년월"] || "").match(/(\d{4})-(\d{2})/);
    const estYear = est ? parseInt(est[1], 10) : null;
    const age = est ? Math.floor((now - new Date(+est[1], +est[2] - 1, 1)) / (365.25 * 864e5)) : null;
    const indCode = (String(info["표준산업분류(10차)"] || info["표준산업분류(11차)"] || "").match(/\(([A-Z])\d+/) || [])[1] || "";
    const indName = String(info["표준산업분류(10차)"] || "").replace(/^\([^)]*\)\s*/, "");
    const reps = d.execs.filter((e) => /대표/.test(e.pos));
    const shares = d.holders.reduce((s, h) => s + (h.shares || 0), 0) || null;

    const cur = (a) => last(a), pv = (a) => prev(a);
    const opm = r("영업이익률") ? last(r("영업이익률")) : (cur(F.op) !== null && cur(F.sales) ? cur(F.op) / cur(F.sales) * 100 : null);
    const opmPrev = r("영업이익률") ? prev(r("영업이익률")) : null;
    const debtRatio = r("부채비율") ? last(r("부채비율")) : (cur(F.liab) !== null && cur(F.equity) ? cur(F.liab) / cur(F.equity) * 100 : null);
    const indAvg = (k) => (d.ind[k] ? d.ind[k].avg : null);
    const taxBurden = r("법인세부담률");
    const payout = r("배당성향") ? last(r("배당성향")) : null;

    // 특수관계 의심: 거래처 대표자명 == 임원/주주명
    const people = new Map();
    d.execs.forEach((e) => people.set(e.name, `${e.pos}${e.rel ? "(" + e.rel + ")" : ""}`));
    d.holders.forEach((h) => { if (!people.has(h.name)) people.set(h.name, `주주(${fmt(h.pct, 1)}%)`); });
    const specials = [...d.buyers.map((b) => ({ ...b, kind: "구매처" })), ...d.sellers.map((s) => ({ ...s, kind: "판매처" }))]
      .filter((t) => t.rep && people.has(t.rep)).map((t) => ({ ...t, role: people.get(t.rep) }));

    const topSeller = d.sellers.filter((s) => s.name !== "기타").sort((a, b) => (b.share || 0) - (a.share || 0))[0] || null;
    const topBuyer = d.buyers.filter((s) => s.name !== "기타").sort((a, b) => (b.share || 0) - (a.share || 0))[0] || null;
    const certKeys = ["벤처", "이노비즈", "메인비즈", "연구개발전담부서", "부설연구소"];
    const certHeld = certKeys.filter((k) => d.certs[k] && !/미인증|-/.test(d.certs[k]) && /인증|확인|보유/.test(d.certs[k]));
    const patents = d.ip["특허"] || 0;
    const purposeNos = d.purposes.map((p) => p.no);
    const allSame = purposeNos.length > 1 && purposeNos.every((x) => x === purposeNos[0]);
    const dupNos = allSame ? [] : [...new Set(purposeNos.filter((x, i) => purposeNos.indexOf(x) !== i))];
    const maxNo = Math.max(0, ...purposeNos);
    const gapNos = Array.from({ length: maxNo }, (_, i) => i + 1).filter((x) => !purposeNos.includes(x) && maxNo > 3);
    const hasRnDPurpose = d.purposes.some((p) => /연구|개발업|기술용역/.test(p.text));
    const familyHolders = d.holders.filter((h) => /본인|가족|친인척/.test(h.rel));

    const certKnown = Object.keys(d.certs).length > 0, ipKnown = Object.keys(d.ip).length > 0;
    const tradeKnown = d.buyers.length + d.sellers.length > 0, purpKnown = d.purposes.length > 0;
    const yearInfo = { estYear: estY0, closingYear: closingY, count: years.length, byAge,
      text: years.length ? `${years.length}개년(${years[0]}${years.length > 1 ? "~" + last(years) : ""})${estY0 ? ` · 설립 ${estY0}년 기준` : ""}` : "재무제표 없음" };
    const A = { yearInfo, certKnown, ipKnown, tradeKnown, purpKnown, years, n, F, accounts, emp, age, estYear, indCode, indName, reps, shares, opm, opmPrev, debtRatio, indAvg,
      taxBurden, payout, specials, topSeller, topBuyer, certHeld, certKeys, patents, dupNos, gapNos, hasRnDPurpose, pyaTotal,
      familyHolders, ratio: r };

    // 상태 판정
    const S = {};
    const salesD = [pv(F.sales), cur(F.sales)];
    const cashChg = pct(pv(F.cash), cur(F.cash));
    S.재무 = {
      st: cur(F.op) !== null && cur(F.op) < 0 ? "위험" : (opm !== null && (indAvg("영업이익률") === null ? opm >= 5 : opm >= indAvg("영업이익률")) && (debtRatio === null || debtRatio < 200) ? "양호" : "주의"),
      key: `매출 ${fmt(cur(F.sales))}백만`, d: salesD,
      t: [opm !== null ? `영업이익률 ${fmt(opm, 1)}%${indAvg("영업이익률") !== null ? `(업종 ${fmt(indAvg("영업이익률"), 1)}%)` : ""}` : null,
        debtRatio !== null ? `부채비율 ${fmt(debtRatio, 1)}%` : null,
        cashChg !== null && cashChg <= -30 ? `현금 ▼${fmt(-cashChg, 0)}%` : null].filter(Boolean).join(" · "),
    };
    const karyoCur = cur(F.karyo) || 0;
    const karyoShare = cur(F.assets) ? karyoCur / cur(F.assets) : 0;
    const pyaShare = cur(F.equity) ? pyaTotal / Math.abs(cur(F.equity)) : 0;
    A.taxMode = karyoCur > 0 && (karyoShare >= 0.01 || pyaShare < 0.05) ? "karyo" : (pyaShare >= 0.05 ? "pya" : (karyoCur > 0 ? "karyo" : "re"));
    if (A.taxMode === "karyo") {
      S.세무 = { st: cur(F.assets) && karyoCur / cur(F.assets) >= 0.05 ? "위험" : "주의", key: `가지급금 ${fmt(karyoCur)}백만`, d: [pv(F.karyo), karyoCur],
        t: `인정이자·상여처분 리스크${cur(F.retained) ? `, 이익잉여금 ${fmt(cur(F.retained))}백만 누적` : ""}` };
    } else if (A.taxMode === "pya") {
      S.세무 = { st: "위험", key: `전기오류수정 ${fmt(pyaTotal)}백만`, d: [pv(F.retained), cur(F.retained)], dLabel: "이익잉여금 ",
        t: "손익을 거치지 않고 잉여금에서 직접 차감 — 세무상 손금 처리 점검" };
    } else {
      S.세무 = { st: "주의", key: `이익잉여금 ${fmt(cur(F.retained))}백만`, d: [pv(F.retained), cur(F.retained)],
        t: `배당성향 ${fmt(payout || 0)}% — 잉여금 누적에 따른 주식가치 상승, 공제·감면 점검` };
    }
    const noRetire = !F.retire || F.retire.every((x) => !x);
    S.노무 = { st: noRetire && emp >= 5 ? "위험" : "주의", key: `인건비 ${fmt(cur(F.labor))}백만`, d: [pv(F.labor), cur(F.labor)],
      t: [emp !== null ? `종업원 ${emp}명` : null, noRetire ? "퇴직급여 비용 미계상" : null,
        emp >= 5 ? "5인 이상 사업장 의무(가산수당·중대재해) 점검" : "인력 확충 시 의무 증가 대비"].filter(Boolean).join(", ") };
    const insRatio = cur(F.insurance) && cur(F.sales) ? cur(F.insurance) / cur(F.sales) * 100 : null;
    S.보험 = { st: "주의", key: `보험료 ${fmt(cur(F.insurance))}백만`, d: [pv(F.insurance), cur(F.insurance)],
      t: `${insRatio !== null ? `매출 대비 ${fmt(insRatio, 2)}% — ` : ""}대표 유고·${indCode === "C" ? "생산물배상·" : indCode === "E" ? "환경책임·" : ""}사업장 재산 보장 점검` };
    const aoiIssues = [dupNos.length || gapNos.length ? "사업목적 번호 정비" : null, reps.length > 1 ? `대표 ${reps.length}인 체제 규정` : null,
      "임원 보수·퇴직금", payout === 0 ? "배당 조항" : null, purpKnown && !hasRnDPurpose ? "연구개발 목적 추가" : null].filter(Boolean);
    S.정관 = { st: "주의", key: purpKnown ? `사업목적 ${d.purposes.length}개` : "정관 원본 확인 필요", d: null, t: aoiIssues.slice(0, 4).join(", ") + " 등 정비 필요" };
    S.인증 = !certKnown ? { st: "주의", key: "인증 현황 확인 필요", d: null, t: "보고서에 기업인증 정보가 없어 인증서·KOITA 등록 여부 확인 필요" } : { st: certHeld.length === 0 ? "개선필요" : certHeld.length <= 2 ? "주의" : "양호",
      key: `인증 ${certHeld.length}건 · 특허 ${ipKnown ? patents + "건" : "확인 필요"}`, d: null,
      t: `${certKeys.filter((k) => !certHeld.includes(k)).join("·")} 미인증${d.bid && d.bid.bids ? `, 낙찰률 ${fmt(d.bid.wins / d.bid.bids * 100, 1)}%` : ""}` };
    const topShare = topSeller ? topSeller.share || 0 : 0;
    S.법무 = {
      st: topShare >= 80 ? "위험" : (topShare >= 30 || specials.length || !tradeKnown) ? "주의" : "양호",
      key: topShare >= 30 ? `판매처 1곳 ${fmt(topShare, 1)}%` : specials.length ? `특수관계 의심 거래 ${specials.length}건` : !tradeKnown ? "거래처 정보 확인 필요" : d.holders.length ? `최대주주 ${fmt(d.holders[0].pct, 1)}%` : "지분 구조 확인 필요",
      d: null,
      t: [topShare >= 30 ? "단일 거래처 계약 안정성" : null, specials.length ? `${specials[0].kind} 대표자명이 당사 ${specials[0].role.replace(/\(.*\)/, "")}와 동일` : null,
        d.related.length ? `관계회사 ${d.related[0].name} 거래 구조` : null, "지분·임원 등기 관리"].filter(Boolean).slice(0, 2).join(", ") + " 점검",
    };
    A.S = S;
    return A;
  }

  // ───────────────────────── PPT 생성 ─────────────────────────
  const COL = { NAVY: "0F2A44", TEAL: "1B9AAA", AMBER: "F2A93B", UP: "D9443A", DOWN: "2F6FDB", GOOD: "2E9E6B", TEXT: "1A2230",
    BG2: "EEF2F6", WHITE: "FFFFFF", GRAY: "6B7785", GRID: "E3E8EE", LINE: "D5DCE4", SOFT: "FDF1E1", LGRAY: "B8C2CC" };
  const STC = { "양호": COL.GOOD, "우수": COL.GOOD, "주의": COL.AMBER, "보통": COL.AMBER, "위험": COL.UP, "개선필요": COL.UP, "미흡": COL.UP };

  function delta(p, c, mode = "pct", unit) {
    if (p === null || p === undefined || c === null || c === undefined || isNaN(p) || isNaN(c)) return { text: "-", color: COL.GRAY };
    const diff = c - p;
    let v;
    if (mode === "pp") v = `${fmt(Math.abs(diff), 1)}${unit || "%p"}`;
    else if (mode === "abs") v = fmt(Math.abs(diff)) + "백만";
    else v = p === 0 ? "신규" : `${fmt(Math.abs(diff / Math.abs(p)) * 100, 1)}%`;
    if (diff > 0) return { text: `▲ 증가 ${v}`, color: COL.UP };
    if (diff < 0) return { text: `▼ 감소 ${v}`, color: COL.DOWN };
    return { text: "― 유지", color: COL.GRAY };
  }

  const dtext = (p, c, mode, unit) => { const x = delta(p, c, mode, unit); return x.text === "-" ? "" : x.text; };

  // 작성자 정보: 컨설턴트·소속은 입력값 그대로, 바닥글이 비어 있으면 "소속 + 컨설턴트"
  const AUTHOR_DEFAULTS = { consultant: "최세림 GFC", brand: "삼성센터법인지점" };
  function defaultFooter(brand, consultant) { return [brand, consultant].map((x) => String(x || "").trim()).filter(Boolean).join(" "); }
  function normalizeAuthor(o) {
    const consultant = String(o.consultant ?? AUTHOR_DEFAULTS.consultant).trim() || AUTHOR_DEFAULTS.consultant;
    const brand = String(o.brand ?? AUTHOR_DEFAULTS.brand).trim();
    const footer = String(o.footer || "").trim() || defaultFooter(brand, consultant);
    const byline = [brand, consultant].filter(Boolean).join("  |  ");
    return Object.assign({}, o, { consultant, brand, footer, byline });
  }

  function buildDeck(PptxGenJS, d, A, opts = {}) {
    const o = normalizeAuthor(Object.assign({ date: new Date() }, opts));
    const C = COL;
    const F = A.F, Y = A.years.length ? A.years : ["", "", ""];
    const company = d.info["기업명"] || "대상 기업";
    const ym = `${o.date.getFullYear()}. ${String(o.date.getMonth() + 1).padStart(2, "0")}`;
    const SRC = `출처: 한국평가데이터 기업종합보고서${o.queryDate ? ` (${o.queryDate} 조회)` : ""}`;
    const FONT = "맑은 고딕";
    const pres = new PptxGenJS();
    pres.layout = "LAYOUT_WIDE";
    pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
    pres.title = `${company} 경영 종합진단 컨설팅 보고서`;
    pres.author = o.consultant;
    pres.company = o.brand.replace(/&/g, "&amp;");

    pres.defineSlideMaster({ title: "COVER", background: { color: C.NAVY }, objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.35, w: 11.7, h: 1.3, fontSize: 40, bold: true, color: C.WHITE, valign: "top", align: "left", margin: 0, fontFace: FONT }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 3.75, w: 11.7, h: 0.9, fontSize: 18, color: "CADCEB", valign: "top", align: "left", margin: 0, fontFace: FONT }, text: "" } },
    ] });
    pres.defineSlideMaster({ title: "CONTENT", background: { color: C.WHITE }, margin: [0.5, 0.6, 0.6, 0.6], objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.88, w: 12.1, h: 0.8, fontSize: 24, bold: true, color: C.TEXT, valign: "middle", align: "left", margin: 0, fontFace: FONT }, text: "" } },
      { text: { text: `${o.footer}  |  ${company} 경영 종합진단`, options: { x: 0.6, y: 7.0, w: 8, h: 0.3, fontSize: 9, color: C.GRAY, margin: 0, fontFace: FONT } } },
    ], slideNumber: { x: 12.2, y: 7.0, w: 0.5, h: 0.3, fontSize: 9, color: C.GRAY, align: "right" } });
    pres.defineSlideMaster({ title: "CLOSING", background: { color: C.NAVY }, objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 0.7, w: 11.7, h: 0.9, fontSize: 32, bold: true, color: C.WHITE, valign: "middle", align: "left", margin: 0, fontFace: FONT }, text: "" } },
    ] });

    let sid = 0; const nm = (p) => `${p}_${++sid}`;
    const T = (s, text, op) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontFace: FONT, objectName: nm("t") }, op));
    const chip = (s, no, label) => T(s, `${no}  ${label}`, { x: 0.6, y: 0.42, w: 1.9, h: 0.38, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.19, fill: { color: C.TEAL }, color: C.WHITE, fontSize: 12, bold: true, align: "center", valign: "middle" });
    const source = (s, txt) => T(s, txt, { x: 0.6, y: 6.68, w: 12.1, h: 0.28, fontSize: 9, color: C.GRAY });
    const box = (s, x, y, w, h, fill, line, shadow) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({ x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: line || fill, width: 1 }, objectName: nm("box") },
      shadow ? { shadow: { type: "outer", color: "000000", opacity: 0.08, blur: 6, offset: 2, angle: 90 } } : {}));
    function kpi(s, x, y, w, h, label, value, unit, dl, sub) {
      box(s, x, y, w, h, C.BG2);
      T(s, label, { x: x + 0.2, y: y + 0.15, w: w - 0.4, h: 0.3, fontSize: 11, color: C.GRAY });
      T(s, [{ text: value, options: { fontSize: value.length > 6 ? 22 : 26, bold: true, color: C.NAVY } }, { text: unit ? ` ${unit}` : "", options: { fontSize: 12, color: C.GRAY } }],
        { x: x + 0.2, y: y + 0.48, w: w - 0.4, h: 0.6, valign: "middle" });
      if (dl && dl.text === "-") dl = null;
      if (dl && h >= 1.4) T(s, dl.text, { x: x + 0.2, y: y + 1.12, w: w - 0.4, h: 0.3, fontSize: 12, bold: true, color: dl.color });
      else if (dl) T(s, dl.text, { x: x + w - 1.6, y: y + 0.15, w: 1.45, h: 0.3, fontSize: 10, bold: true, color: dl.color, align: "right" });
      if (sub && h >= 1.9) T(s, sub, { x: x + 0.2, y: y + 1.45, w: w - 0.4, h: h - 1.55, fontSize: 10, color: C.GRAY, valign: "top" });
    }
    const pill = (s, x, y, w, text, color) => T(s, text, { x, y, w, h: 0.3, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.15, fill: { color }, color: C.WHITE, fontSize: 10, bold: true, align: "center", valign: "middle" });
    function insight(s, x, y, w, h, title, body) {
      box(s, x, y, w, h, C.NAVY);
      T(s, [{ text: title, options: { fontSize: 13, bold: true, color: C.AMBER, breakLine: true } }, { text: body, options: { fontSize: 11.5, color: C.WHITE } }],
        { x: x + 0.25, y: y + 0.12, w: w - 0.5, h: h - 0.24, valign: "middle", paraSpaceAfter: 4 });
    }
    const hdr = (t, align = "center") => ({ text: t, options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, align, valign: "middle" } });
    const cell = (t, op = {}) => ({ text: t, options: Object.assign({ color: C.TEXT, valign: "middle" }, op) });
    const tbl = (s, rows, op) => s.addTable(rows, Object.assign({ fontFace: FONT, border: { type: "solid", pt: 0.75, color: C.LINE }, margin: [0, 0.1, 0, 0.1], objectName: nm("tbl") }, op));
    const chartBase = () => ({ catAxisLabelColor: C.GRAY, valAxisLabelColor: C.GRAY, catAxisLabelFontFace: FONT, valAxisLabelFontFace: FONT, dataLabelFontFace: FONT, legendFontFace: FONT, titleFontFace: FONT,
      catAxisLabelFontSize: 11, valAxisLabelFontSize: 10, dataLabelFontSize: 10, legendFontSize: 11, valGridLine: { color: C.GRID, size: 0.75 }, catGridLine: { style: "none" },
      catAxisLineShow: false, valAxisLineShow: false, titleColor: C.TEXT, titleFontSize: 13, showTitle: true, objectName: nm("chart") });
    const z = (a) => (a || []).map((v) => (v === null || isNaN(v) ? 0 : Math.round(v)));
    const col = (s, series, op) => s.addChart(pres.charts.BAR, series, Object.assign(chartBase(), { barDir: "col", barGapWidthPct: 55, showValue: true, dataLabelPosition: "outEnd", dataLabelColor: C.TEXT, dataLabelFormatCode: "#,##0", valAxisLabelFormatCode: "#,##0", showLegend: series.length > 1, legendPos: "t" }, op));
    const cur = (a) => last(a), pv = (a) => prev(a);
    const S = A.S;
    const order = ["재무", "세무", "노무", "보험", "정관", "인증", "법무"];

    // 1. 표지
    {
      const s = pres.addSlide({ masterName: "COVER" });
      T(s, `CONSULTING REPORT  ·  ${ym}`, { x: 0.8, y: 1.7, w: 8, h: 0.4, fontSize: 13, bold: true, color: C.AMBER, charSpacing: 2 });
      s.addText(`${company} 경영 종합진단`, { placeholder: "title" });
      s.addText("재무 · 세무 · 노무 · 보험 · 정관 · 인증 · 법무 7대 영역 점검 및 실행 제안", { placeholder: "body" });
      const worst = order.find((k) => ["위험", "개선필요"].includes(S[k].st)) || "재무";
      order.forEach((a, i) => T(s, a, { x: 0.8 + i * 1.05, y: 5.0, w: 0.9, h: 0.9, shape: pres.shapes.OVAL, fill: { color: a === worst ? C.AMBER : "1B3A5C" }, line: { color: "3D6A8F", width: 1 }, color: C.WHITE, fontSize: 13, bold: true, align: "center", valign: "middle" }));
      T(s, o.byline, { x: 0.8, y: 6.45, w: 8, h: 0.35, fontSize: 12, color: "CADCEB" });
      T(s, `기초자료: 한국평가데이터 기업종합보고서${o.queryDate ? ` (${o.queryDate} 조회)` : ""}`, { x: 7.0, y: 6.45, w: 5.5, h: 0.35, fontSize: 10, color: "8FA6BC", align: "right" });
    }

    // 2. 기업 개요
    {
      const s = pres.addSlide({ masterName: "CONTENT" });
      chip(s, "01", "기업 개요");
      const top = d.holders[0];
      const famPct = A.familyHolders.reduce((x, h) => x + h.pct, 0);
      s.addText(`${A.age !== null ? `업력 ${A.age}년차 ` : ""}${A.indName || "기업"}${top ? `, 최대주주 ${top.name} ${fmt(top.pct, 1)}%${famPct >= 99 ? " (특수관계인 100%)" : ""}` : ""}`, { placeholder: "title" });
      const execTxt = d.execs.map((e) => `${e.pos} ${e.name}`).slice(0, 6).join(" · ");
      const creditBad = Object.entries(d.credit).filter(([k, v]) => /연체|신용도|공공|정지|회생|행정/.test(k) && v && !/해당사항없음|없음/.test(v));
      const rows = [
        ["기업명", `${company}${d.info["영문기업명"] ? ` (${d.info["영문기업명"]})` : ""}`],
        ["대표이사", A.reps.length ? A.reps.map((e) => e.name).join(" · ") : (d.info["대표자명"] || "-")],
        ["설립일", d.info["설립년월"] || "-"],
        ["업종", d.info["표준산업분류(10차)"] || "-"],
        ["주요제품", d.info["주요제품(상품)"] || "-"],
        ["소재지", `${(d.info["주소"] || "-").replace(/^\(\d+\)\s*/, "")}${d.info["자가소유 여부"] ? ` (${d.info["자가소유 여부"] === "소유" ? "자가" : "임차"})` : ""}`],
        ["기업규모", `${d.info["기업규모"] || "-"}${A.emp !== null ? ` · 종업원 ${A.emp}명` : ""}`],
        ["자본금", cur(F.capital) !== null ? `${fmt(cur(F.capital))}백만원${A.shares ? ` (${fmt(A.shares)}주)` : ""}` : "-"],
        ["등기임원", execTxt || "-"],
        ["신용정보", creditBad.length ? creditBad.map(([k, v]) => `${k}: ${v}`).join(", ") : "연체·공공정보·행정처분 등 해당없음"],
      ].map(([k, v]) => [cell(k, { bold: true, color: C.NAVY, fill: { color: C.BG2 } }), cell(v, { fill: { color: C.WHITE }, color: k === "신용정보" && creditBad.length ? C.UP : C.TEXT })]);
      tbl(s, rows, { x: 0.6, y: 1.95, w: 6.7, colW: [1.5, 5.2], rowH: 0.44, fontSize: 11.5, margin: [0, 0.12, 0, 0.12] });
      if (d.holders.length) {
        const hs = d.holders.slice(0, 6);
        const rest = 100 - hs.reduce((x, h) => x + h.pct, 0);
        const labels = hs.map((h) => `${h.name}${h.rel ? `(${h.rel})` : ""}`), vals = hs.map((h) => h.pct);
        if (rest > 0.5) { labels.push("기타"); vals.push(+rest.toFixed(2)); }
        s.addChart(pres.charts.DOUGHNUT, [{ name: "지분율", labels, values: vals }], Object.assign(chartBase(), { x: 7.6, y: 1.85, w: 5.1, h: 3.15, holeSize: 55, title: "주주 구성 (%)",
          chartColors: [C.NAVY, C.TEAL, "6CC3CE", "B5E2E7", C.AMBER, C.LGRAY, "DDE3EA"].slice(0, labels.length), showLegend: true, legendPos: "r", showValue: false, showPercent: true, dataLabelColor: "FFFFFF", dataLabelFontSize: 10 }));
      } else {
        box(s, 7.6, 1.95, 5.1, 3.0, C.BG2);
        T(s, "주주 현황 정보 없음", { x: 7.6, y: 1.95, w: 5.1, h: 3.0, fontSize: 13, color: C.GRAY, align: "center", valign: "middle" });
      }
      kpi(s, 7.6, 5.15, 1.6, 1.45, "업력", A.age !== null ? String(A.age) : "-", "년");
      kpi(s, 9.35, 5.15, 1.6, 1.45, "종업원", A.emp !== null ? String(A.emp) : "-", "명");
      kpi(s, 11.1, 5.15, 1.6, 1.45, "업계 매출순위", d.rank ? fmt(d.rank) : "-", "위");
      source(s, `${SRC} · 업계순위는 동일 업종 최근 결산 매출 기준`);
    }

    // 3. 종합 진단
    {
      const s = pres.addSlide({ masterName: "CONTENT" });
      chip(s, "02", "종합 진단");
      const good = order.filter((k) => S[k].st === "양호"), bad = order.filter((k) => ["위험", "개선필요"].includes(S[k].st));
      s.addText(`${good.length ? good.join("·") + " 영역은 양호" : "전 영역 점검 필요"}${bad.length ? `, ${bad.join("·")} 영역은 우선 개선 필요` : ", 세부 관리 체계 보완 필요"}`, { placeholder: "title" });
      const w = 2.86, h = 2.2, gap = 0.22;
      order.forEach((a, i) => {
        const cd = S[a]; const col_ = i % 4, row = Math.floor(i / 4);
        const x = 0.6 + col_ * (w + gap), y = 1.95 + row * (h + 0.2);
        box(s, x, y, w, h, C.WHITE, C.LINE, true);
        T(s, a, { x: x + 0.2, y: y + 0.15, w: 1.4, h: 0.4, fontSize: 16, bold: true, color: C.NAVY, valign: "middle" });
        pill(s, x + w - 1.2, y + 0.2, 1.0, cd.st, STC[cd.st]);
        T(s, cd.key, { x: x + 0.2, y: y + 0.65, w: w - 0.4, h: 0.4, fontSize: cd.key.length > 14 ? 13.5 : 15, bold: true, color: C.TEXT, valign: "middle" });
        const dl = cd.d ? delta(cd.d[0], cd.d[1]) : null;
        const hasD = dl && dl.text !== "-";
        if (hasD) T(s, `${cd.dLabel || "전년 대비 "}${dl.text}`, { x: x + 0.2, y: y + 1.05, w: w - 0.4, h: 0.3, fontSize: 11, bold: true, color: dl.color });
        T(s, cd.t, { x: x + 0.2, y: y + (hasD ? 1.38 : 1.1), w: w - 0.4, h: hasD ? 0.75 : 1.0, fontSize: 10.5, color: C.GRAY, valign: "top" });
      });
      const rank = { "위험": 0, "개선필요": 0, "주의": 1, "양호": 2 };
      const pri = [...order].sort((a, b) => rank[S[a].st] - rank[S[b].st]).slice(0, 4).join(" → ");
      insight(s, 0.6 + 3 * (w + gap), 1.95 + (h + 0.2), w, h, "읽는 법", `▲ 증가(적색) / ▼ 감소(청색)는 전년 대비 수치 방향입니다.\n상태: 양호 · 주의 · 위험/개선필요\n\n우선순위: ${pri}`);
      source(s, `${SRC} · 재무 ${A.yearInfo.text} (단위: 백만원)`);
    }

    // 4. 재무① 손익
    {
      const s = pres.addSlide({ masterName: "CONTENT" });
      chip(s, "03", "재무 ①");
      const sd = delta(pv(F.sales), cur(F.sales)), od = delta(pv(F.op), cur(F.op));
      const growth = A.n >= 3 && F.sales && F.sales[0] ? cur(F.sales) / F.sales[0] : null;
      s.addText(`매출 ${fmt(cur(F.sales))}백만원${sd.text === "-" ? "" : " " + sd.text.replace(/ 증가| 감소/, "")}${A.opm !== null ? `, 영업이익률 ${fmt(A.opm, 1)}%` : ""}${A.indAvg("영업이익률") !== null ? ` (업종 ${fmt(A.indAvg("영업이익률"), 1)}%)` : ""}`, { placeholder: "title" });
      col(s, [{ name: "매출액", labels: Y, values: z(F.sales) }, { name: "영업이익", labels: Y, values: z(F.op) }, { name: "당기순이익", labels: Y, values: z(F.ni) }],
        { x: 0.6, y: 1.9, w: 7.3, h: 4.7, chartColors: [C.NAVY, C.TEAL, C.AMBER], title: "손익 추이 (백만원)", dataLabelFontSize: 9 });
      const costRate = A.ratio("매출원가율");
      kpi(s, 8.2, 1.95, 2.15, 2.25, `매출액 (${last(Y)})`, fmt(cur(F.sales)), "백만원", sd, growth ? `${Y[0]}년 대비 ${fmt(growth, 1)}배` : null);
      kpi(s, 10.55, 1.95, 2.15, 2.25, `영업이익 (${last(Y)})`, fmt(cur(F.op)), "백만원", od, null);
      kpi(s, 8.2, 4.35, 2.15, 2.25, "영업이익률", fmt(A.opm, 1), "%", delta(A.opmPrev, A.opm, "pp"), A.indAvg("영업이익률") !== null ? `업종평균 ${fmt(A.indAvg("영업이익률"), 1)}%` : null);
      if (costRate) kpi(s, 10.55, 4.35, 2.15, 2.25, "매출원가율", fmt(last(costRate), 1), "%", delta(prev(costRate), last(costRate), "pp"), "낮을수록 수익성 개선");
      else kpi(s, 10.55, 4.35, 2.15, 2.25, "당기순이익", fmt(cur(F.ni)), "백만원", delta(pv(F.ni), cur(F.ni)), null);
      source(s, `${SRC} · 손익계산서 · 재무 ${A.yearInfo.text} (단위: 백만원)`);
    }

    // 5. 재무② 재무상태·주요 계정
    {
      const s = pres.addSlide({ masterName: "CONTENT" });
      chip(s, "03", "재무 ②");
      const cashD = pct(pv(F.cash), cur(F.cash));
      const niPos = cur(F.ni) > 0;
      const rising = A.accounts.filter((a) => !/현금|예금|차입|채무|법인세|충당금/.test(a.label)).map((a) => ({ label: a.label, inc: (last(a.v) || 0) - (prev(a.v) || 0) }))
        .filter((a) => a.inc > 0).sort((a, b) => b.inc - a.inc).slice(0, 3);
      s.addText(cashD !== null && cashD < -20 && niPos
        ? `이익은 났지만 현금은 ▼${fmt(-cashD, 0)}% — ${rising.length ? rising.map((r) => r.label).slice(0, 2).join("·") + "으로 자금이 묶이는 구조" : "자금 흐름 점검 필요"}`
        : `자산 ${fmt(cur(F.assets))}백만원, 부채비율 ${fmt(A.debtRatio, 1)}% — ${A.debtRatio !== null && A.debtRatio < 100 ? "안정적 재무구조" : "재무구조 개선 필요"}`, { placeholder: "title" });
      col(s, [{ name: "자산총계", labels: Y, values: z(F.assets) }, { name: "부채총계", labels: Y, values: z(F.liab) }, { name: "자본총계", labels: Y, values: z(F.equity) }],
        { x: 0.6, y: 1.9, w: 5.2, h: 4.7, chartColors: [C.NAVY, C.UP, C.TEAL], title: "재무상태 (백만원)", dataLabelFontSize: 9 });
      const accs = A.accounts.slice(0, 8);
      const rows = [[hdr("주요 계정", "left"), ...Y.map((y) => hdr(y)), hdr("전년 대비")]];
      accs.forEach((a) => {
        const p = prev(a.v), c = last(a.v);
        const big = p !== null && p !== 0 && Math.abs((c - p) / p) > 3;
        const dl = p === null || p === undefined ? { text: "-", color: C.GRAY } : big || p === 0 ? delta(p, c, "abs") : delta(p, c);
        const ch = pct(p, c);
        const hi = /가지급금|현금/.test(a.label) || (ch !== null && Math.abs(ch) >= 50 && Math.abs(c - p) >= (cur(F.assets) || 1) * 0.03);
        rows.push([cell(a.label, { bold: hi, fill: { color: hi ? C.SOFT : C.WHITE } }), ...a.v.map((v, i) => cell(v ? fmt(v) : "-", { align: "right", bold: i === a.v.length - 1 })),
          cell(dl.text, { color: dl.color, bold: true, align: "center" })]);
      });
      const cw = Y.length === 3 ? [1.9, 0.95, 0.95, 0.95, 1.85] : [1.9, ...Y.map(() => 2.85 / Y.length), 1.85];
      if (accs.length) tbl(s, rows, { x: 6.1, y: 1.95, w: 6.6, colW: cw, rowH: 0.36, fontSize: 10.5 });
      const borrowC = cur(F.borrow);
      const allowAcc = A.accounts.find((a) => a.label === "대손충당금");
      const allow = allowAcc ? (last(allowAcc.v) || 0) - (prev(allowAcc.v) || 0) : 0;
      const body = cashD !== null && cashD < 0 && niPos && rising.length
        ? `${last(Y)}년 순이익 ${fmt(cur(F.ni))}백만 발생에도 현금은 ${fmt(Math.abs(cur(F.cash) - pv(F.cash)))}백만 감소. ${rising.map((r) => `${r.label}(+${fmt(r.inc)})`).join("·")}로 자금이 유출되었습니다.${allow ? ` 이 중 대손충당금 ${fmt(allow)}백만이 설정되어 회수가 불확실합니다.` : ""} 회수 조건과 자금 계획 점검이 필요합니다.`
        : `현금성자산 ${fmt(cur(F.cash))}백만${dtext(pv(F.cash), cur(F.cash)) ? ` (${dtext(pv(F.cash), cur(F.cash))})` : ""}, 차입금 ${fmt(borrowC || 0)}백만. 부채비율 ${fmt(A.debtRatio, 1)}%${A.indAvg("부채비율") !== null ? ` (업종 ${fmt(A.indAvg("부채비율"), 1)}%)` : ""}로 재무 안정성을 판단할 수 있습니다.`;
      insight(s, 6.1, 1.95 + 0.36 * (accs.length + 1) + 0.3, 6.6, 6.6 - (1.95 + 0.36 * (accs.length + 1) + 0.3), "자금 흐름 진단", body);
      source(s, `${SRC} · 재무상태표 (단위: 백만원)`);
    }

    // 6. 재무③ 업종 비교
    {
      const s = pres.addSlide({ masterName: "CONTENT" });
      chip(s, "03", "재무 ③");
      const LOWER = ["부채비율", "차입금의존도", "차입금/매출액"];
      const keys = ["매출액증가율", "영업이익증가율", "총자산증가율", "영업이익률", "EBITDA마진율", "총자산순이익률(ROA)", "부채비율", "유동비율", "차입금의존도", "이자보상배수(배)", "매출채권회전율(회)", "총자본회전율(회)", "자기자본회전율(회)"]
        .filter((k) => d.ind[k] && d.ind[k].cur !== null).slice(0, 11);
      const evalOf = (k, avg, c) => {
        if (avg === null || c === null) return "보통";
        const lower = LOWER.includes(k);
        const better = lower ? c <= avg : c >= avg;
        const strong = lower ? c <= avg * 0.5 : (avg > 0 ? c >= avg * 1.5 : c - avg >= 5);
        return better ? (strong ? "우수" : "양호") : (lower ? c <= avg * 1.1 : c >= avg * 0.9) ? "보통" : "미흡";
      };
      const evals = keys.map((k) => evalOf(k, d.ind[k].avg, d.ind[k].cur));
      const good = evals.filter((e) => e === "우수" || e === "양호").length;
      s.addText(keys.length ? `업종 비교 ${keys.length}개 지표 중 ${good}개가 업종평균 이상${good >= keys.length * 0.7 ? " — 정책자금·보증 심사에 유리" : " — 약점 지표 관리 필요"}` : "업종 비교 지표", { placeholder: "title" });
      const pk = ["하위25%", "평균", "상위25%", "조회기업"];
      if (d.peer["평균"] && d.peer["조회기업"]) {
        const labs = ["하위25%", "평균", "상위25%", "당사"];
        col(s, [{ name: "매출액", labels: labs, values: z(pk.map((k) => d.peer[k] ? d.peer[k]["매출액"] : 0)) }, { name: "영업이익", labels: labs, values: z(pk.map((k) => d.peer[k] ? d.peer[k]["영업이익"] : 0)) }],
          { x: 0.6, y: 1.9, w: 5.0, h: 4.7, chartColors: [C.NAVY, C.AMBER], title: "동종업계 규모 비교 (백만원)", dataLabelFontSize: 9 });
      } else if (keys.length) {
        const ks = keys.filter((k) => /이익률|마진|ROA/.test(k)).slice(0, 4);
        s.addChart(pres.charts.BAR, [{ name: "업종평균", labels: ks, values: ks.map((k) => d.ind[k].avg || 0) }, { name: "당사", labels: ks, values: ks.map((k) => d.ind[k].cur || 0) }],
          Object.assign(chartBase(), { x: 0.6, y: 1.9, w: 5.0, h: 4.7, barDir: "bar", chartColors: [C.LGRAY, C.TEAL], showLegend: true, legendPos: "t", showValue: true, dataLabelPosition: "outEnd", dataLabelColor: C.TEXT, dataLabelFormatCode: "0.0", title: "업종평균 대비 (%)", catAxisOrientation: "maxMin" }));
      }
      const rows = [[hdr("지표", "left"), hdr("업종평균"), hdr(`당사 ${last(Y)}`), hdr("전년 대비"), hdr("평가")]];
      keys.forEach((k, i) => {
        const rv = A.ratio(k) || A.ratio(k.replace(/\((회|배)\)/, ""));
        const p = rv ? prev(rv) : null, c = d.ind[k].cur;
        const unit = /\(회\)/.test(k) ? "회" : /\(배\)/.test(k) ? "배" : "%p";
        const dl = delta(p, c, "pp", unit);
        rows.push([cell(k.replace("총자산순이익률(ROA)", "ROA(%)")), cell(fmt(d.ind[k].avg, 1), { align: "right", color: C.GRAY }), cell(fmt(c, 1), { align: "right", bold: true }),
          cell(dl.text, { align: "center", color: dl.color, bold: true }), cell(evals[i], { align: "center", bold: true, color: STC[evals[i]] })]);
      });
      if (keys.length) tbl(s, rows, { x: 5.9, y: 1.95, w: 6.8, colW: [2.0, 1.05, 1.1, 1.75, 0.9], rowH: 0.385, fontSize: 10.5, margin: [0, 0.08, 0, 0.08] });
      else T(s, "보고서에 업종 비교(재무진단) 자료가 없습니다.", { x: 5.9, y: 3.5, w: 6.8, h: 0.6, fontSize: 14, color: C.GRAY, align: "center" });
      source(s, `${SRC} · 재무진단·동종업계 경영규모 비교 · ${Y.length >= 2 ? `전년 대비는 ${Y[Y.length - 2]}→${last(Y)} 변동폭` : "전년 비교 자료 없음(1개년)"}`);
    }

    // 7. 세무① 핵심 리스크
    {
      const s = pres.addSlide({ masterName: "CONTENT" });
      chip(s, "04", "세무 ①");
      let risks, chartSeries, chartTitle, colors, k1, k2, adv;
      if (A.taxMode === "karyo") {
        const kc = cur(F.karyo), interest = kc * 0.046;
        s.addText(`가지급금 ${fmt(kc)}백만원${dtext(pv(F.karyo), kc) ? " " + dtext(pv(F.karyo), kc).replace(/ 증가| 감소/, "") : ""} — 가장 시급한 세무 리스크`, { placeholder: "title" });
        chartSeries = [{ name: "가지급금", labels: Y, values: z(F.karyo) }]; chartTitle = "가지급금 잔액 (백만원)"; colors = [C.UP];
        k1 = ["총자산 대비", fmt(kc / cur(F.assets) * 100, 1), "%"]; k2 = ["인정이자(연, 추정)", `약 ${fmt(interest)}`, "백만원"];
        risks = [["인정이자 익금산입", "대표에게 이자를 받지 않으면 당좌대출이자율(4.6%) 기준 이자를 법인 수익으로 보고 과세, 같은 금액은 대표 상여로 처분"],
          ["대표 소득세 증가", "미수 인정이자는 대표 상여로 처리되어 근로소득세·4대보험 부담 동반"],
          ["대외 신용도 하락", "금융기관·보증기관 심사 시 부실자산으로 보아 차감 평가 — 정책자금 한도에 불리"],
          ["폐업·양도 시 상여처분", "회수되지 않은 가지급금은 최종적으로 대표 상여로 처분되어 일시에 과세"]];
        adv = ["해소 방향", "급여·상여 재설계 및 배당으로 상환 재원 마련 → 임원 퇴직금 규정 정비 → 연차별 상환 계획 수립 (세무대리인과 금액 확정)"];
      } else if (A.taxMode === "pya") {
        s.addText(`${A.n}년간 ${fmt(A.pyaTotal)}백만원이 손익을 거치지 않고 잉여금에서 직접 차감 — 최우선 점검`, { placeholder: "title" });
        chartSeries = [{ name: "당기순이익", labels: Y, values: z(F.ni) }, { name: "전기오류수정 차감액", labels: Y, values: z(F.pya) }];
        chartTitle = "순이익 vs 잉여금 직접 차감 (백만원)"; colors = [C.TEAL, C.UP];
        k1 = ["누적 차감액", fmt(A.pyaTotal), "백만원"]; k2 = ["이익잉여금", fmt(cur(F.retained)), "백만원", delta(pv(F.retained), cur(F.retained))];
        risks = [["세무상 손금 인정 여부", "전기오류수정 손실은 해당 사업연도 손금 요건(귀속시기·증빙)을 충족해야 인정 — 경정청구·수정신고 대상 여부 검토"],
          ["채권 대손 요건", "대손충당금·채권 상각은 채무자 파산·소멸시효 등 대손 사유와 증빙 확보 필요"],
          ["특수관계자 채권 여부", "상대방이 특수관계인이면 대손 손금 불인정 및 인정이자 문제 발생"],
          ["대외 신뢰도", "반복적 오류수정은 금융기관·조달 심사에서 회계 신뢰성 감점 요인"]];
        adv = ["권고", "연도별 전기오류수정 내역서와 채권 원장을 확보해 세무조정 내역과 대조 → 경정청구 가능 여부 및 회계처리 기준 확정"];
      } else {
        s.addText(`이익잉여금 ${fmt(cur(F.retained))}백만원 누적 — 주식가치 상승 전 절세·지분 전략 필요`, { placeholder: "title" });
        chartSeries = [{ name: "이익잉여금", labels: Y, values: z(F.retained) }, { name: "법인세비용", labels: Y, values: z(F.tax) }];
        chartTitle = "잉여금·법인세 (백만원)"; colors = [C.NAVY, C.AMBER];
        k1 = ["배당성향", fmt(A.payout || 0), "%"]; k2 = ["법인세비용", fmt(cur(F.tax)), "백만원", delta(pv(F.tax), cur(F.tax))];
        risks = [["비상장주식 가치 상승", "잉여금이 쌓일수록 주식 평가액이 올라 지분 이동·상속 시 세 부담 증가"],
          ["미처분 잉여금 활용", "배당·임원 보상 정책 없이 유보만 지속되면 주주 소득 실현 경로가 제한"],
          ["공제·감면 누락", "연구개발·고용 관련 세액공제 등 적용 가능 항목 점검 필요"],
          ["가업승계 준비", "지분 구조·후계 계획에 맞춘 사전 증여·승계 특례 검토"]];
        adv = ["권고", "주식가치 사전 평가 → 배당·급여 정책 수립 → 지분 이동 시점·방식 결정 (세무대리인과 협의)"];
      }
      col(s, chartSeries, { x: 0.6, y: 1.9, w: 4.9, h: 3.4, chartColors: colors, title: chartTitle });
      kpi(s, 0.6, 5.45, 2.35, 1.15, k1[0], k1[1], k1[2], k1[3] || null);
      kpi(s, 3.15, 5.45, 2.35, 1.15, k2[0], k2[1], k2[2], k2[3] || null);
      risks.forEach(([t, dsc], i) => {
        const y = 1.95 + i * 0.86;
        T(s, String(i + 1), { x: 5.8, y: y + 0.12, w: 0.5, h: 0.5, shape: pres.shapes.OVAL, fill: { color: C.UP }, color: C.WHITE, bold: true, fontSize: 14, align: "center", valign: "middle" });
        T(s, [{ text: t, options: { bold: true, fontSize: 13, color: C.TEXT, breakLine: true } }, { text: dsc, options: { fontSize: 10.5, color: C.GRAY } }], { x: 6.5, y, w: 6.2, h: 0.8, valign: "middle" });
      });
      insight(s, 5.8, 5.45, 6.9, 1.15, adv[0], adv[1]);
      source(s, `${SRC} · 재무상태표·이익잉여금처분계산서 (단위: 백만원)${A.taxMode === "karyo" ? " · 인정이자는 기말잔액 × 4.6% 단순 추정" : ""}`);
    }

    // 8. 세무② 공제·감면
    {
      const s = pres.addSlide({ masterName: "CONTENT" });
      chip(s, "04", "세무 ②");
      const tb = A.taxBurden;
      s.addText(tb && last(tb) !== null ? `법인세 부담률 ${prev(tb) !== null ? `${fmt(prev(tb), 1)}% → ` : ""}${fmt(last(tb), 1)}%, 공제·감면 적용 여부 전면 점검` : "법인세 공제·감면 적용 여부 전면 점검", { placeholder: "title" });
      const tbi = tb ? tb.map((v, i) => (v === null ? -1 : i)).filter((i) => i >= 0) : [];
      if (tbi.length >= 2) s.addChart(pres.charts.LINE, [{ name: "법인세 부담률", labels: tbi.map((i) => Y[i]), values: tbi.map((i) => tb[i]) }], Object.assign(chartBase(), { x: 0.6, y: 1.9, w: 4.9, h: 3.4, chartColors: [C.UP], lineSize: 3, lineDataSymbol: "circle", lineDataSymbolSize: 9, showLegend: false, showValue: true, dataLabelPosition: "t", dataLabelColor: C.TEXT, dataLabelFormatCode: "0.0", valAxisMinVal: 0, title: "법인세 부담률 (%)" }));
      else col(s, [{ name: "법인세비용", labels: Y, values: z(F.tax) }], { x: 0.6, y: 1.9, w: 4.9, h: 3.4, chartColors: [C.UP], title: "법인세비용 (백만원)" });
      const nav = A.shares && cur(F.equity) ? cur(F.equity) * 1e6 / A.shares : null;
      kpi(s, 0.6, 5.45, 2.35, 1.15, nav ? "1주당 순자산(장부)" : "이익잉여금", nav ? fmt(nav) : fmt(cur(F.retained)), nav ? "원" : "백만원");
      kpi(s, 3.15, 5.45, 2.35, 1.15, "배당성향", fmt(A.payout || 0), "%");
      const rnd = cur(F.rnd) || 0;
      const lab = A.certHeld.includes("연구개발전담부서") || A.certHeld.includes("부설연구소");
      const rows = [[hdr("항목", "left"), hdr("현황", "left"), hdr("검토 포인트", "left")],
        [cell("법인세 변동", { bold: true }), cell([{ text: (pv(F.tax) !== null ? `법인세비용 ${fmt(pv(F.tax))}→${fmt(cur(F.tax))}백만 ${delta(pv(F.tax), cur(F.tax)).text.replace(/증가 |감소 /, "")}` : `법인세비용 ${fmt(cur(F.tax))}백만`), options: { color: delta(pv(F.tax), cur(F.tax)).color, bold: true } }]), cell("감면·공제 적용 또는 일시 요인인지 원인 확인")],
        [cell(A.age !== null && A.age <= 5 ? "창업중소기업 세액감면" : "중소기업 특별세액감면", { bold: true }), cell(`${A.estYear || "-"}년 설립 · ${d.info["기업규모"] || "-"}`), cell(A.age !== null && A.age <= 5 ? "대표 연령·소재지 요건 충족 시 5년간 감면 — 적용 여부 확인" : "업종·규모별 감면율 적용 여부 확인")],
        [cell("연구·인력개발비 세액공제", { bold: true }), cell(rnd ? `연구개발비 ${fmt(rnd)}백만${lab ? " · 연구조직 보유" : ""}` : [{ text: "연구개발비 0원", options: { color: C.UP, bold: true } }]), cell(lab ? "연구원 인건비 공제 누락 여부, 연구노트·증빙 관리" : "전담부서 설립 시 연구원 인건비 공제(중소기업 25%)")],
        [cell("통합고용세액공제", { bold: true }), cell(A.emp !== null ? `종업원 ${A.emp}명` : "-"), cell("상시근로자 증가분 공제, 청년·정규직 전환 우대")],
        [cell("업무용승용차", { bold: true }), cell(cur(F.vehicles) ? `차량운반구 ${fmt(cur(F.vehicles))}백만 (취득가)` : "차량 보유·리스 확인"), cell("운행기록부·전용보험, 감가상각 한도 관리")],
        [cell("접대비·기부금", { bold: true }), cell(`접대비 ${fmt(cur(F.entertain) || 0)}백만${cur(F.donation) ? ` · 기부금 ${fmt(cur(F.donation))}백만` : ""}`), cell("적격증빙 수취 및 손금 한도 관리")]];
      tbl(s, rows, { x: 5.8, y: 1.95, w: 6.9, colW: [2.05, 2.35, 2.5], rowH: 0.55, fontSize: 10.5, margin: [0, 0.08, 0, 0.08] });
      insight(s, 5.8, 5.95, 6.9, 0.65, "핵심", `자본 ${fmt(cur(F.equity))}백만 · 이익잉여금 ${fmt(cur(F.retained))}백만 — 지분 이동 전 주식가치 평가가 필요합니다.`);
      source(s, `${SRC} · 손익계산서·재무비율 · 감면 적용 여부는 법인세 신고서(세무조정계산서) 확인 필요`);
    }

    // 9. 노무
    {
      const s = pres.addSlide({ masterName: "CONTENT" });
      chip(s, "05", "노무");
      const emp = A.emp || 0;
      const salesG = pct(pv(F.sales), cur(F.sales)), laborG = pct(pv(F.labor), cur(F.labor));
      s.addText(`${salesG !== null && laborG !== null ? `매출 ${salesG >= 0 ? "▲" : "▼"}${fmt(Math.abs(salesG), Math.abs(salesG) < 10 ? 1 : 0)}% 대비 인건비 ${laborG >= 0 ? "▲" : "▼"}${fmt(Math.abs(laborG), Math.abs(laborG) < 10 ? 1 : 0)}% — ` : ""}${emp >= 5 ? `${emp}인 사업장 의무 점검·보상체계 정비` : "인력 확충 대비 노무 체계 준비"}`, { placeholder: "title" });
      const series = [F.salary && { name: "판관 급여", labels: Y, values: z(F.salary) }, F.mcLabor && { name: "제조 노무비", labels: Y, values: z(F.mcLabor) }, F.retire && { name: "퇴직급여", labels: Y, values: z(F.retire) }].filter(Boolean);
      if (series.length) s.addChart(pres.charts.BAR, series, Object.assign(chartBase(), { x: 0.6, y: 1.9, w: 4.9, h: 3.4, barDir: "col", barGrouping: "stacked", barGapWidthPct: 55, chartColors: [C.NAVY, C.TEAL, C.AMBER].slice(0, series.length), showLegend: true, legendPos: "t", showValue: true, dataLabelPosition: "ctr", dataLabelColor: "FFFFFF", dataLabelFormatCode: "#,##0;;;", dataLabelFontSize: 9, valAxisLabelFormatCode: "#,##0", title: "인건비 구성 (백만원)" }));
      const perHead = A.ratio("종업원1인당인건비(백만원)");
      const lr = A.ratio("노무비/매출액");
      if (perHead) kpi(s, 0.6, 5.45, 2.35, 1.15, "1인당 인건비", fmt(last(perHead), 1), "백만원", delta(prev(perHead), last(perHead)));
      else kpi(s, 0.6, 5.45, 2.35, 1.15, "종업원", String(A.emp ?? "-"), "명");
      if (lr) kpi(s, 3.15, 5.45, 2.35, 1.15, "노무비/매출액", fmt(last(lr), 2), "%", delta(prev(lr), last(lr), "pp"));
      else kpi(s, 3.15, 5.45, 2.35, 1.15, "인건비 합계", fmt(cur(F.labor)), "백만원", delta(pv(F.labor), cur(F.labor)));
      const noRetire = !F.retire || F.retire.every((x) => !x);
      const stc = { "시급": C.UP, "점검": C.AMBER, "준비": C.TEAL, "기회": C.GOOD };
      const rows = [[hdr("점검 항목", "left"), hdr("근거·현황", "left"), hdr("점검 내용", "left"), hdr("상태")]];
      const items = [
        ["퇴직급여 적립", noRetire ? "재무제표에 퇴직급여 비용·충당부채 없음" : `퇴직급여 ${fmt(cur(F.retire))}백만 비용처리`, "퇴직연금(DC/DB) 도입·적립 현황 확인", noRetire ? "시급" : "점검"],
        ["중대재해처벌법", emp >= 5 ? "5인 이상 사업장 적용(2024.1.27~)" : "5인 미만 — 적용 제외", "안전보건관리체계·위험성평가 문서화", emp >= 5 ? "시급" : "준비"],
        ["근로계약·임금명세서", `종업원 ${A.emp ?? "-"}명`, "서면 계약서 작성·교부, 임금명세서 발급", "점검"],
        ["가산수당·연차휴가", emp >= 5 ? "5인 이상 근로기준법 전면 적용" : "5인 미만 일부 적용", "연장·야간·휴일수당, 포괄임금 운영 적정성", emp >= 5 ? "점검" : "준비"],
        ["취업규칙", emp >= 10 ? "10인 이상 작성·신고 의무 대상" : "10인 이상 작성·신고 의무", emp >= 10 ? "신고 여부·최신 법령 반영 확인" : "인원 증가 대비 사전 정비", emp >= 10 ? "점검" : "준비"],
        ["고용 지원·공제", "신규 채용·정규직 전환 시", "청년·고령자 고용지원금, 고용세액공제 연계", "기회"],
      ];
      items.forEach(([a, b, c2, st]) => rows.push([cell(a, { bold: true }), cell(b), cell(c2), cell(st, { color: stc[st], bold: true, align: "center" })]));
      tbl(s, rows, { x: 5.8, y: 1.95, w: 6.9, colW: [1.7, 2.4, 2.2, 0.6], rowH: 0.6, fontSize: 10.5, margin: [0, 0.08, 0, 0.08] });
      source(s, `${SRC} · 손익계산서·제조원가명세서 · 종업원 수는 보고서 조회 기준`);
    }

    // 10. 보험
    {
      const s = pres.addSlide({ masterName: "CONTENT" });
      chip(s, "06", "보험");
      const top = d.holders[0], ts = A.topSeller;
      s.addText(`${A.familyHolders.length >= d.holders.length && d.holders.length ? "대표·가족 중심 지배구조" : "주요 주주·임원 중심 경영"} — 경영 공백·사업장·배상 리스크를 보장으로 분산`, { placeholder: "title" });
      const insR = cur(F.insurance) && cur(F.sales) ? cur(F.insurance) / cur(F.sales) * 100 : null;
      kpi(s, 0.6, 1.95, 2.9, 1.5, `연간 보험료 (${last(Y)})`, fmt(cur(F.insurance)), "백만원", delta(pv(F.insurance), cur(F.insurance)));
      kpi(s, 0.6, 3.6, 2.9, 1.5, "매출 대비 보험료", fmt(insR, 2), "%");
      insight(s, 0.6, 5.25, 2.9, 1.35, insR !== null && insR < 0.5 ? "보장 공백 신호" : "점검 포인트", insR !== null && insR < 0.5 ? `매출 ${fmt(cur(F.sales))}백만 대비 보험료 지출이 낮음` : "가입 증권의 중복·공백을 함께 확인");
      const cands = [
        [1, "대표·주요주주 유고", `${top ? `최대주주 ${top.name} ${fmt(top.pct, 1)}%` : "대표 중심 경영"}${ts && ts.share >= 30 ? `, 판매처 ${fmt(ts.share, 1)}% 의존` : ""}`, "경영인정기보험(유고 시 운영자금), 임원 퇴직금 재원"],
        A.indCode === "C" && [1, "제품 결함 배상", `${d.info["주요제품(상품)"] || "제품"} 제조·납품`, "생산물배상책임보험(PL)"],
        (A.indCode === "E" || /폐기물|환경/.test(A.indName)) && [1, "환경오염 사고", `${A.indName} 시설 운영`, "환경책임보험 가입 대상·보상한도 확인"],
        (cur(F.construction) > 0 || A.indCode === "F") && [2, "공사·제3자 배상", cur(F.construction) ? `공사수입 ${fmt(cur(F.construction))}백만` : "건설 공사 수행", "건설공사보험·영업배상책임"],
        cur(F.vehicles) >= 100 && [2, "차량·중장비 사고", `차량운반구 ${fmt(cur(F.vehicles))}백만 (취득가)`, "화물차·건설기계 보험, 적재물배상책임"],
        [2, "사업장 재산 손해", `${d.info["자가소유 여부"] === "소유" ? "자가" : "임차"} 사업장${d.info["담보제공여부"] === "담보제공" ? " · 담보 제공 중" : ""}`, "화재·재산종합보험" + (d.info["자가소유 여부"] === "소유" ? "" : ", 임차자배상책임")],
        [2, "근로자 재해", `종업원 ${A.emp ?? "-"}명`, "근로자재해보장책임(산재 초과분), 단체상해"],
        [3, "임원 책임", `등기임원 ${d.execs.length}명`, "임원배상책임(D&O) — 규모 확대 시 검토"],
      ].filter(Boolean).sort((a, b) => a[0] - b[0]).slice(0, 5);
      const rows = [[hdr("리스크", "left"), hdr("회사 현황 (근거)", "left"), hdr("권장 보장", "left"), hdr("우선순위")]];
      const pc = { 1: C.UP, 2: C.AMBER, 3: C.TEAL };
      cands.forEach(([p, a, b, c2]) => rows.push([cell(a, { bold: true }), cell(b), cell(c2), cell(`${p}순위`, { bold: true, color: pc[p], align: "center" })]));
      tbl(s, rows, { x: 3.8, y: 1.95, w: 8.9, colW: [1.95, 3.15, 2.9, 0.9], rowH: 0.78, fontSize: 11 });
      source(s, `${SRC} · 보험료 = 판관비 + 제조원가 보험료 · 현재 가입 증권 확인 후 중복·공백 분석 필요`);
    }

    // 11. 정관
    {
      const s = pres.addSlide({ masterName: "CONTENT" });
      chip(s, "07", "정관");
      s.addText(`${A.reps.length > 1 ? `대표 ${A.reps.length}인 체제·` : ""}보수·퇴직금·배당·사업목적 조항 정비가 다른 과제의 전제`, { placeholder: "title" });
      const hs = d.holders.map((h) => h.name).join("·");
      const purpIssue = [A.dupNos.length ? `번호 중복(${A.dupNos.join(",")})` : null, A.gapNos.length ? `누락(${A.gapNos.slice(0, 3).join(",")})` : null].filter(Boolean).join("·");
      const items = [
        A.reps.length > 1 && ["대표이사 체제", `대표이사 ${A.reps.map((e) => e.name).join("·")} (${A.reps.length}인)`, "각자·공동대표 권한 범위, 이사회 결의사항 명시", "의사결정 충돌 예방", false],
        ["임원 보수·퇴직금", "지급규정 유무 확인 필요", "주총 결의 지급규정, 지급배수·산정기준 명시", "퇴직금 손금 인정·보상 재원", false],
        ["이익배당", `배당성향 ${fmt(A.payout || 0)}% · 잉여금 ${fmt(cur(F.retained))}백만`, "중간배당·현물배당 근거 조항", "잉여금 유동화, 주주 소득 분산", !A.payout],
        A.purpKnown ? ["사업목적", `${d.purposes.length}개${purpIssue ? " · " + purpIssue : ""}`, `미영위 업종 정리·순번 정비${A.hasRnDPurpose ? "" : ", 연구개발업 추가"}`, "인증·R&D·정책자금 요건 충족", !!purpIssue || !A.hasRnDPurpose]
          : ["사업목적", "보고서에 정보 없음 — 등기부 확인", "영위 사업과 목적 일치 여부 점검", "인증·R&D·정책자금 요건 충족", false],
        ["주식양도 제한", d.holders.length ? `${hs} ${d.holders.length}인 보유` : "주주 구성 확인", "이사회 승인·우선매수권, 주주 간 계약 연계", "외부 지분 유입 방지", false],
        ["자기주식 취득", "조항 유무 확인 필요", "상법 절차에 따른 취득 근거 마련", "주주 지분 정리 수단", false],
        ["유족보상·경조금", "규정 유무 확인 필요", "임원 유족보상금 규정 신설", "유고 시 보상 재원 (보험 연계)", false],
      ].filter(Boolean).slice(0, 6);
      const rows = [[hdr("정비 항목", "left"), hdr("현재 상태", "left"), hdr("개정 방향", "left"), hdr("기대 효과", "left")]];
      items.forEach(([a, b, c2, e, red]) => rows.push([cell(a, { bold: true }), cell(red ? [{ text: b, options: { color: C.UP, bold: true } }] : b), cell(c2), cell(e)]));
      tbl(s, rows, { x: 0.6, y: 1.95, w: 12.1, colW: [2.0, 3.0, 3.6, 3.5], rowH: 0.5, fontSize: 11, margin: [0, 0.12, 0, 0.12] });
      ["정관·등기부 확보", "조항별 전면 검토", "개정안·규정 작성", "주주총회 특별결의", "변경 등기·비치"].forEach((t, i) => {
        const x = 0.6 + i * 2.46;
        T(s, `${i + 1}. ${t}`, { x, y: 6.1, w: 2.2, h: 0.45, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, fill: { color: i === 3 ? C.AMBER : C.BG2 }, color: i === 3 ? C.WHITE : C.NAVY, bold: true, fontSize: 11, align: "center", valign: "middle" });
        if (i < 4) T(s, "›", { x: x + 2.2, y: 6.1, w: 0.26, h: 0.45, fontSize: 18, bold: true, color: C.GRAY, align: "center", valign: "middle" });
      });
      source(s, `${SRC} · 사업목적·경영진·주주 현황 · 정관 원본 대조 필요`);
    }

    // 12. 인증
    {
      const s = pres.addSlide({ masterName: "CONTENT" });
      chip(s, "08", "인증");
      const bid = d.bid && d.bid.bids ? d.bid : null;
      s.addText(!A.certKnown ? "기업인증 현황 확인 필요 — 인증 로드맵 사전 검토" : bid ? `입찰 ${fmt(bid.bids)}건 중 낙찰 ${fmt(bid.wins)}건 — 인증 가점으로 공공조달 경쟁력 강화` : `기업인증 ${A.certHeld.length}건 · 특허 ${A.patents}건 — ${A.certHeld.length < 2 ? "인증 로드맵 착수 적기" : "추가 인증으로 지원 범위 확대"}`, { placeholder: "title" });
      const tiles = [["벤처", "벤처기업"], ["이노비즈", "이노비즈"], ["메인비즈", "메인비즈"], ["연구개발전담부서", "연구개발전담부서"], ["부설연구소", "기업부설연구소"]];
      tiles.forEach(([k, label], i) => {
        const x = 0.6 + i * 2.05, held = A.certHeld.includes(k);
        box(s, x, 1.95, 1.85, 1.2, C.BG2);
        T(s, label, { x, y: 2.05, w: 1.85, h: 0.4, fontSize: 12, bold: true, color: C.NAVY, align: "center" });
        T(s, !A.certKnown ? "확인 필요" : held ? "✓  인증" : "✕  미인증", { x, y: 2.5, w: 1.85, h: 0.5, fontSize: 16, bold: true, color: !A.certKnown ? C.GRAY : held ? C.GOOD : C.UP, align: "center", valign: "middle" });
      });
      box(s, 0.6 + 5 * 2.05, 1.95, 1.85, 1.2, C.BG2);
      T(s, "특허", { x: 0.6 + 5 * 2.05, y: 2.05, w: 1.85, h: 0.4, fontSize: 12, bold: true, color: C.NAVY, align: "center" });
      T(s, !A.ipKnown ? "확인 필요" : A.patents ? `✓  ${A.patents}건` : "✕  0건", { x: 0.6 + 5 * 2.05, y: 2.5, w: 1.85, h: 0.5, fontSize: 16, bold: true, color: !A.ipKnown ? C.GRAY : A.patents ? C.GOOD : C.UP, align: "center", valign: "middle" });
      const has = (k) => A.certHeld.includes(k);
      const yr3 = A.age !== null && A.age >= 3 ? "업력 3년 충족" : `업력 3년 도래 (${A.estYear ? A.estYear + 3 : "-"}년)`;
      let steps = [
        !has("연구개발전담부서") && !has("부설연구소") && ["연구개발전담부서", "KOITA 신고 · R&D 세액공제 시작점"],
        has("연구개발전담부서") && !has("부설연구소") && ["기업부설연구소", "전담부서 전환 — 연구전담요원 요건 충족 시"],
        !A.patents && [A.ipKnown ? "특허 출원" : "특허 현황 확인·출원", "핵심 기술 권리화 · 인증 심사 가점"],
        !has("메인비즈") && ["메인비즈", `경영혁신형 중소기업 · ${yr3}`],
        !has("이노비즈") && ["이노비즈", `기술혁신형 중소기업 · ${yr3}`],
        !has("벤처") && ["벤처기업 확인", "정책자금·보증 우대, 조달 가점"],
        A.patents && ["특허 추가 출원", "공정·제품 개선 기술 권리화"],
      ].filter(Boolean).slice(0, 4);
      while (steps.length < 4) steps.push(["인증 유지·갱신", "유효기간·사후관리 일정 관리"]);
      const periods = ["0~3개월", "1~3개월", "3~6개월", "6~12개월"];
      const x0 = bid ? 3.25 : 0.6, sw = bid ? 2.2 : 2.85, step = bid ? 2.38 : 3.08;
      if (bid) {
        kpi(s, 0.6, 3.4, 2.4, 1.5, "나라장터 입찰 (최근 1년)", fmt(bid.bids), "건");
        kpi(s, 0.6, 5.05, 2.4, 1.5, `낙찰 ${fmt(bid.wins)}건${bid.amount ? ` (${fmt(bid.amount)}백만원)` : ""}`, fmt(bid.wins / bid.bids * 100, 1), "%");
      }
      steps.forEach(([t, dsc], i) => {
        const x = x0 + i * step;
        box(s, x, 3.4, sw, 3.15, C.WHITE, C.LINE, true);
        T(s, String(i + 1), { x: x + 0.18, y: 3.58, w: 0.55, h: 0.55, shape: pres.shapes.OVAL, fill: { color: i === 0 ? C.AMBER : C.TEAL }, color: C.WHITE, bold: true, fontSize: 16, align: "center", valign: "middle" });
        T(s, periods[i], { x: x + 0.8, y: 3.65, w: sw - 0.95, h: 0.4, fontSize: 10.5, bold: true, color: C.GRAY, align: "right", valign: "middle" });
        T(s, t, { x: x + 0.18, y: 4.3, w: sw - 0.3, h: 0.5, fontSize: 15, bold: true, color: C.NAVY, valign: "middle" });
        T(s, dsc, { x: x + 0.18, y: 4.85, w: sw - 0.3, h: 1.5, fontSize: 11, color: C.GRAY, valign: "top" });
      });
      source(s, `${SRC} · 기업인증·산업재산권${bid ? "·나라장터 입찰정보" : ""} · 소요기간은 일반적 준비기간 예시`);
    }

    // 13. 법무
    {
      const s = pres.addSlide({ masterName: "CONTENT" });
      chip(s, "09", "법무");
      const ts = A.topSeller, sp = A.specials[0];
      s.addText(!A.tradeKnown ? "거래처·지분·임원 등기 관련 법무 리스크 사전 점검" : ts && ts.share >= 50 ? `매출의 ${fmt(ts.share, 1)}%가 한 곳에서 — 거래계약·특수관계·지분 구조를 문서로 지켜야` : sp ? "거래처는 분산 구조, 특수관계 의심 거래와 지분 구조는 문서화 필요" : "거래·지분·임원 등기 관련 법무 리스크 사전 점검", { placeholder: "title" });
      const barChart = (list, x, title, highlight) => {
        const L = list.filter((t) => t.name !== "기타" && t.share).slice(0, 4);
        if (!L.length) { box(s, x, 1.95, 2.9, 2.9, C.BG2); T(s, `${title}\n정보 없음`, { x, y: 1.95, w: 2.9, h: 2.9, fontSize: 12, color: C.GRAY, align: "center", valign: "middle" }); return; }
        const short = (n_) => { const t = n_.replace(/[(（]\s*(주|합자|유|사)\s*[)）]|주식회사/g, "").trim(); return t.length > 9 ? t.slice(0, 8) + "…" : t; };
        s.addChart(pres.charts.BAR, [{ name: title, labels: L.map((t) => short(t.name)), values: L.map((t) => t.share) }], Object.assign(chartBase(), { x, y: 1.9, w: 2.95, h: 3.0, barDir: "bar", catAxisOrientation: "maxMin",
          chartColors: L.map((t) => (A.specials.find((p) => p.name === t.name) ? C.UP : highlight)), showLegend: false, showValue: true, dataLabelPosition: "outEnd", dataLabelColor: C.TEXT, dataLabelFormatCode: "0.0", catAxisLabelFontSize: 9, valAxisHidden: true, valGridLine: { style: "none" }, title: `${title} (%)` }));
      };
      barChart(d.sellers, 0.6, "판매처 상위", C.NAVY);
      barChart(d.buyers, 3.65, "구매처 상위", C.TEAL);
      const ins = sp ? ["특수관계 거래 점검", `${sp.kind} ${sp.name}(${fmt(sp.share, 1)}%)의 대표자명이 당사 ${sp.role}와 같은 '${sp.rep}'. 동일인이면 시가 거래·계약서 구비가 필수입니다(부당행위계산부인 대비).`]
        : ts && ts.share >= 30 ? ["거래 집중 리스크", `판매처 ${ts.name} 비중 ${fmt(ts.share, 1)}%. 거래기본계약(단가·대금지급·하자책임)과 기술자료·지식재산권 귀속 조항을 서면으로 확보해야 합니다.`]
        : !A.tradeKnown ? ["거래처 확인 필요", "보고서에 거래처 정보가 없습니다. 주요 매출·매입처와 거래 비중, 계약서 보유 여부를 별도로 확인해야 합니다."]
        : ["거래 구조", "주요 거래처가 분산되어 있습니다. 거래기본계약·비밀유지계약 체계를 정비해 두면 분쟁 대응력이 높아집니다."];
      insight(s, 0.6, 5.1, 6.0, 1.5, ins[0], ins[1]);
      const fam = A.familyHolders.reduce((x, h) => x + h.pct, 0);
      const cards = [
        ts && ts.share >= 30 && ["거래계약 정비", `${ts.name} 거래기본계약서·발주서 체계 점검, 기술자료 제공 시 비밀유지계약(NDA) 체결`],
        sp && ["특수관계자 거래", `${sp.name} 거래 시 시가 기준·계약서 구비, 동일인 여부는 등기부로 확인`],
        d.related.length && ["관계회사 거래", `${d.related.map((r_) => r_.name).slice(0, 2).join(", ")} (${d.related[0].rel || "관계회사"}) — 거래 조건 시가 유지·일감몰아주기 요건 점검`],
        d.holders.length && ["주주·지분 관리", fam >= 99 ? `대표·가족 ${d.holders.length}인 100% — 주주명부·주식 취득자금 근거 정비, 지분 이동 계획 수립` : `주주 ${d.holders.length}인 지분 분산 — 의결권·양도·퇴임 시 매입가 산정을 주주 간 계약으로 정리`],
        A.reps.length > 1 && ["대표이사 체제", `대표 ${A.reps.length}인 — 권한 분장·대표권 제한 등기 여부 확인`],
        d.info["담보제공여부"] === "담보제공" && ["담보·자산 관리", `사업장 담보 제공 중 — 차입금 ${fmt(cur(F.borrow) || 0)}백만 약정 조건과 담보 범위 확인`],
        ["임원 임기·등기", "등기임원 임기 만료 시점 확인 및 중임·변경 등기 기한 관리"],
      ].filter(Boolean).slice(0, 4);
      cards.forEach(([t, dsc], i) => {
        const y = 1.95 + i * 1.18;
        box(s, 6.9, y, 5.8, 1.05, C.BG2);
        T(s, String(i + 1), { x: 7.05, y: y + 0.27, w: 0.5, h: 0.5, shape: pres.shapes.OVAL, fill: { color: C.NAVY }, color: C.WHITE, bold: true, fontSize: 13, align: "center", valign: "middle" });
        T(s, [{ text: t, options: { bold: true, fontSize: 13, color: C.TEXT, breakLine: true } }, { text: dsc, options: { fontSize: 10.5, color: C.GRAY } }], { x: 7.75, y: y + 0.08, w: 4.8, h: 0.9, valign: "middle" });
      });
      source(s, `${SRC} · 거래처·관계회사·주주·경영진 현황`);
    }

    // 14. 로드맵
    {
      const s = pres.addSlide({ masterName: "CONTENT" });
      chip(s, "10", "실행 로드맵");
      s.addText("즉시 진단 → 3개월 내 제도 정비 → 1년 내 인증·지분 전략 완성", { placeholder: "title" });
      const taxNow = A.taxMode === "karyo" ? "가지급금 내역·인정이자 산정" : A.taxMode === "pya" ? "전기오류수정·채권 대손 내역 검토" : "공제·감면 적용 내역 검토";
      const certFirst = A.certHeld.length === 0 ? "연구개발전담부서 설립" : !A.certHeld.includes("메인비즈") ? "메인비즈 신청" : "추가 인증 준비";
      const cols = [
        ["즉시 (1개월)", C.AMBER, [["세무", taxNow], ["정관", "정관·등기부·주주명부 확보 및 검토"], ["노무", "근로계약·퇴직급여·안전 현황 점검"], ["보험", "가입 증권 보장 분석"]]],
        ["단기 (3개월)", C.TEAL, [["인증", certFirst], ["정관", "개정안 작성 · 주주총회 결의"], ["노무", "퇴직연금 정비 · 안전보건 문서화"], ["보험", "핵심 리스크 보장 설계"], ["법무", A.specials.length ? "특수관계 거래 계약 정비" : "거래기본계약·NDA 정비"]]],
        ["중기 (6~12개월)", C.NAVY, [["인증", "특허 출원 → 벤처·이노비즈·메인비즈"], ["세무", A.taxMode === "karyo" ? "급여·배당 정책으로 가지급금 단계적 축소" : "주식가치 평가 및 지분 이동 계획"], ["재무", "인증 기반 정책자금·보증 활용"], ["법무", "지분 이동·승계 플랜 수립"]]],
      ];
      cols.forEach(([t, c2, items], ci) => {
        const x = 0.6 + ci * 4.1, w = 3.85;
        T(s, t, { x, y: 1.95, w, h: 0.55, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08, fill: { color: c2 }, color: C.WHITE, bold: true, fontSize: 15, align: "center", valign: "middle" });
        items.forEach(([area, txt], ii) => {
          const y = 2.65 + ii * 0.8;
          box(s, x, y, w, 0.68, C.BG2);
          T(s, area, { x: x + 0.12, y: y + 0.17, w: 0.7, h: 0.34, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.17, fill: { color: C.WHITE }, line: { color: c2, width: 1 }, color: c2, bold: true, fontSize: 10, align: "center", valign: "middle" });
          T(s, txt, { x: x + 0.95, y, w: w - 1.05, h: 0.68, fontSize: 11.5, color: C.TEXT, valign: "middle" });
        });
        if (ci < 2) T(s, "›", { x: x + w, y: 1.95, w: 0.25, h: 0.55, fontSize: 20, bold: true, color: C.GRAY, align: "center", valign: "middle" });
      });
      source(s, "일정은 자료 확보 및 회사 의사결정 일정에 따라 조정");
    }

    // 15. 요청 자료
    {
      const s = pres.addSlide({ masterName: "CLOSING" });
      s.addText("다음 단계 — 정밀 진단을 위한 요청 자료", { placeholder: "title" });
      const req = [
        ["최근 3개년 법인세 신고서", "세무조정계산서 · 공제/감면 내역"],
        A.taxMode === "karyo" ? ["가지급금 원장", "발생·상환 내역, 약정이자 여부"] : A.taxMode === "pya" ? ["전기오류수정 내역서", "연도별 수정 사유 및 증빙"] : ["주식 평가 기초자료", "부동산 시가·주주명부"],
        ["현행 정관 · 법인등기부등본", "임원 임기 및 조항 확인"],
        ["주주명부", "지분 취득 경위·시점"],
        ["근로계약서 · 임금대장", "퇴직급여·수당 산정 방식"],
        ["보험 가입 증권", "법인·대표 명의 전체"],
        ["주요 거래 계약서", A.topSeller ? `${A.topSeller.name.slice(0, 12)} 등 주요 거래처` : "주요 매출·매입처"],
        ["연구 인력·공간 현황", "인증·연구조직 요건 검토용"],
      ];
      req.forEach(([t, dsc], i) => {
        const c2 = i % 2, row = Math.floor(i / 2);
        const x = 0.8 + c2 * 6.0, y = 1.95 + row * 1.05;
        T(s, String(i + 1).padStart(2, "0"), { x, y, w: 0.7, h: 0.7, shape: pres.shapes.OVAL, fill: { color: "1B3A5C" }, line: { color: C.AMBER, width: 1.5 }, color: C.AMBER, bold: true, fontSize: 14, align: "center", valign: "middle" });
        T(s, [{ text: t, options: { bold: true, fontSize: 15, color: C.WHITE, breakLine: true } }, { text: dsc, options: { fontSize: 11, color: "A9BDD1" } }], { x: x + 0.9, y: y - 0.05, w: 4.9, h: 0.8, valign: "middle" });
      });
      T(s, "본 자료는 공개·신용정보 기반 예비 진단으로, 세무·법률 판단은 원본 자료 확인 후 확정됩니다.", { x: 0.8, y: 6.15, w: 11.7, h: 0.3, fontSize: 10, color: "8FA6BC" });
      T(s, o.byline, { x: 0.8, y: 6.55, w: 8, h: 0.35, fontSize: 12, color: "CADCEB" });
    }
    return pres;
  }

  function queryDateOf(rows) {
    for (const r of rows) for (const c of r) { const m = String(c).match(/조회일시\s*:\s*(\d{4}-\d{2}-\d{2})/); if (m) return m[1].replace(/-/g, "."); }
    return null;
  }

  function fileNameFor(d) {
    const nm = String((d.info && d.info["기업명"]) || "기업").replace(/[(（]\s*주\s*[)）]|주식회사|\s/g, "").replace(/[\\/:*?"<>|]/g, "");
    return `${nm || "기업"}_경영종합진단_컨설팅.pptx`;
  }

  const api = { fileNameFor, MAX_YEARS, AUTHOR_DEFAULTS, defaultFooter, normalizeAuthor, num, fmt, cleanRows, rowsFromWorkbook, rowsFromPdf, parseReport, analyze, buildDeck, delta, queryDateOf };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.BizReport = api;
})(typeof window !== "undefined" ? window : globalThis);
