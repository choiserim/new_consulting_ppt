(function () {
  "use strict";
  const B = window.BizReport;
  const $ = (s) => document.getElementById(s);
  const AREAS = ["재무", "세무", "노무", "보험", "정관", "인증", "법무"];
  const OUTLINE = [["표지", ""], ["기업 개요", "주주 구성"], ["종합 진단", "7대 영역 상태"], ["재무 ①", "손익 추이"], ["재무 ②", "재무상태·주요 계정"], ["재무 ③", "업종 비교"],
    ["세무 ①", "핵심 세무 리스크"], ["세무 ②", "공제·감면 점검"], ["노무", "인건비·의무 점검"], ["보험", "리스크별 보장"], ["정관", "정비 항목"], ["인증", "현황·로드맵"],
    ["법무", "거래처·특수관계"], ["실행 로드맵", "즉시·단기·중기"], ["요청 자료", "다음 단계"]];
  let state = null;

  if (window.pdfjsLib) {
    try { window.pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js"; } catch (e) {}
  }
  const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const step = (n) => [1, 2, 3].forEach((i) => $("st" + i).classList.toggle("on", i === n));
  function msg(id, text, err) { const el = $(id); el.hidden = !text; el.textContent = text || ""; el.classList.toggle("err", !!err); }

  // ── 작성자 정보: 바닥글은 '소속 + 컨설턴트'와 자동 연동(직접 고치면 연동 해제)
  const STORE = "plana-bizreport-author";
  let footerTouched = false;
  function author() {
    return B.normalizeAuthor({ consultant: $("consultant").value, brand: $("brand").value, footer: $("footer").value });
  }
  function syncAuthor() {
    if (!footerTouched) $("footer").value = B.defaultFooter($("brand").value, $("consultant").value);
    $("footerReset").hidden = !footerTouched;
    $("footerHint").textContent = footerTouched ? "(직접 입력)" : "(소속 + 컨설턴트 이름 자동 연동)";
    const o = author();
    $("pvByline").textContent = o.byline;
    $("pvFooter").textContent = o.footer;
    try { localStorage.setItem(STORE, JSON.stringify({ consultant: $("consultant").value, brand: $("brand").value, footer: $("footer").value, footerTouched })); } catch (e) {}
  }
  function loadAuthor() {
    try {
      const v = JSON.parse(localStorage.getItem(STORE) || "null");
      if (v) { $("consultant").value = v.consultant ?? $("consultant").value; $("brand").value = v.brand ?? $("brand").value; footerTouched = !!v.footerTouched; if (footerTouched) $("footer").value = v.footer || ""; }
    } catch (e) {}
    syncAuthor();
  }

  function renderOutline() {
    $("outline").innerHTML = OUTLINE.map(([a, b]) => `<li><div>${esc(a)}${b ? ` <span>· ${esc(b)}</span>` : ""}</div></li>`).join("");
    $("areasEmpty").innerHTML = AREAS.map((a) => `<div class="area"><div class="hd"><span class="nm">${a}</span><span class="pill s-대기">대기</span></div><div class="key" style="color:var(--muted)">보고서 필요</div></div>`).join("");
  }

  function dcls(t) { return /^▲/.test(t) ? "up" : /^▼/.test(t) ? "down" : "flat"; }

  function render() {
    const { d, A } = state;
    $("emptyState").hidden = true; $("filled").hidden = false;
    $("coName").textContent = d.info["기업명"] || "(기업명 미확인)";
    $("coMeta").textContent = [A.indName, d.info["설립년월"] ? `설립 ${d.info["설립년월"]}` : "", A.emp !== null ? `종업원 ${A.emp}명` : "", state.queryDate ? `${state.queryDate} 조회` : ""].filter(Boolean).join(" · ");
    $("areas").innerHTML = AREAS.map((a) => {
      const s = A.S[a];
      const dl = s.d ? B.delta(s.d[0], s.d[1]) : null;
      return `<div class="area"><div class="hd"><span class="nm">${a}</span><span class="pill s-${esc(s.st)}">${esc(s.st)}</span></div>
        <div class="key">${esc(s.key)}</div>${dl && dl.text !== "-" ? `<div class="dl ${dcls(dl.text)}">${esc((s.dLabel || "") + dl.text)}</div>` : ""}</div>`;
    }).join("");
    const F = A.F, Y = A.years;
    $("yearNote").textContent = `(단위: 백만원 · ${A.yearInfo.text} · 전년 대비 ▲ 증가 · ▼ 감소)`;
    const lines = [["매출액", F.sales], ["영업이익", F.op], ["당기순이익", F.ni], ["자산총계", F.assets], ["부채총계", F.liab], ["자본총계", F.equity], ["현금성자산", F.cash],
      ["가지급금", F.karyo], ["이익잉여금", F.retained], ["인건비", F.labor], ["보험료", F.insurance]].filter(([, v]) => v && v.some((x) => x));
    $("fin").innerHTML = `<thead><tr><th>계정</th>${Y.map((y) => `<th>${esc(y)}</th>`).join("")}<th>전년 대비</th></tr></thead><tbody>` +
      lines.map(([k, v]) => { const dl = B.delta(v.length > 1 ? v[v.length - 2] : null, v[v.length - 1]); return `<tr><td>${k}</td>${v.map((x) => `<td class="n">${B.fmt(x)}</td>`).join("")}<td class="${dcls(dl.text)}"><b>${esc(dl.text)}</b></td></tr>`; }).join("") + "</tbody>";
    const checks = [
      ["재무제표 " + A.yearInfo.text, Y.length >= 1],
      ["업종 비교 지표 " + Object.keys(d.ind).length + "개", Object.keys(d.ind).length > 0],
      ["동종업계 규모 비교", !!d.peer["평균"]],
      ["주주 " + d.holders.length + "명", d.holders.length > 0],
      ["임원 " + d.execs.length + "명", d.execs.length > 0],
      ["구매처·판매처 " + (d.buyers.length + d.sellers.length) + "곳", d.buyers.length + d.sellers.length > 0],
      ["기업인증 현황", Object.keys(d.certs).length > 0],
      ["산업재산권", Object.keys(d.ip).length > 0],
      ["사업목적 " + d.purposes.length + "개", d.purposes.length > 0],
      ["나라장터 입찰정보", !!(d.bid && d.bid.bids)],
      ["업계 순위", !!d.rank],
      ["특수관계 의심 거래 " + A.specials.length + "건", true],
    ];
    $("checks").innerHTML = checks.map(([t, ok]) => `<li class="${ok ? "" : "miss"}">${esc(t)}</li>`).join("");
    $("raw").textContent = state.rows.slice(0, 400).map((r) => r.join(" | ")).join("\n");
  }

  async function handle(file) {
    if (!file) return;
    msg("msg", "보고서를 읽는 중입니다…"); msg("genMsg", ""); $("gen").disabled = true;
    state = null; $("filled").hidden = true; $("emptyState").hidden = false; step(1); // 이전 결과가 새 파일 결과로 보이지 않도록 비운다
    $("fileName").hidden = false; $("fileName").textContent = file.name;
    try {
      const buf = await file.arrayBuffer();
      let rows;
      if (/\.pdf$/i.test(file.name) || file.type === "application/pdf") {
        if (!window.pdfjsLib) throw new Error("PDF 처리 라이브러리를 불러오지 못했습니다. 인터넷 연결을 확인한 뒤 새로고침해 주세요.");
        rows = await B.rowsFromPdf(window.pdfjsLib, buf);
        if (rows.length < 20) throw new Error("PDF에서 글자를 거의 읽지 못했습니다. 스캔 이미지 PDF일 수 있으니 엑셀(.xls) 보고서를 올려 주세요.");
      } else {
        if (!window.XLSX) throw new Error("엑셀 처리 라이브러리를 불러오지 못했습니다. 인터넷 연결을 확인한 뒤 새로고침해 주세요.");
        rows = B.rowsFromWorkbook(window.XLSX, buf);
      }
      const d = B.parseReport(rows);
      if (!d.info["기업명"] || (!d.years.length && !d.sumYears.length)) throw new Error("한국평가데이터 기업종합보고서 형식을 찾지 못했습니다. 보고서 원본 파일인지 확인해 주세요.");
      const A = B.analyze(d);
      state = { rows, d, A, queryDate: B.queryDateOf(rows), file: file.name };
      render();
      const warn = [!A.F.sales && "매출액", !Object.keys(d.ind).length && "업종 비교", !d.holders.length && "주주 현황", !A.certKnown && "기업인증", !A.tradeKnown && "거래처"].filter(Boolean);
      const lastY = A.years[A.years.length - 1];
      const yearGap = d.closingYear && lastY && +lastY < +d.closingYear;
      const text = yearGap
        ? `주의: 보고서 결산일자는 ${d.closingYear}년인데 재무제표는 ${lastY}년까지만 읽혔습니다. PDF의 오른쪽 열이 잘렸을 수 있으니 엑셀(.xls) 보고서로 다시 올리는 것을 권장합니다.`
        : warn.length ? `읽기 완료. 일부 항목(${warn.join(", ")})을 보고서에서 찾지 못해 해당 슬라이드는 '확인 필요'로 표시됩니다.` : "읽기 완료. 오른쪽 결과를 확인한 뒤 PPT를 만드세요.";
      msg("msg", text, !!yearGap);
      $("gen").disabled = false; step(2);
    } catch (e) {
      state = null; $("filled").hidden = true; $("emptyState").hidden = false;
      msg("msg", e.message || "파일을 읽지 못했습니다.", true); step(1);
    }
  }

  async function generate() {
    if (!state) return;
    if (!window.PptxGenJS) { msg("genMsg", "PPT 라이브러리를 불러오지 못했습니다. 새로고침해 주세요.", true); return; }
    $("gen").disabled = true; msg("genMsg", "PPT를 만드는 중입니다…");
    try {
      const pres = B.buildDeck(window.PptxGenJS, state.d, state.A, {
        ...author(), queryDate: state.queryDate,
      });
      const name = B.fileNameFor(state.d);
      const inArtifact = window.claude && typeof window.claude.use === "function";
      if (inArtifact) {
        const dl = await window.claude.use("downloads");
        if (!dl) throw { code: "unavailable" };
        const blob = await pres.write({ outputType: "blob" });
        await dl.save({ filename: name, data: blob });
        msg("genMsg", `${name} 저장을 요청했습니다.`);
      } else {
        await pres.writeFile({ fileName: name });
        msg("genMsg", `${name} 파일을 내려받았습니다.`);
      }
      step(3);
    } catch (e) {
      const code = e && e.code;
      const text = code === "declined" ? "저장을 취소했습니다. 다시 누르면 저장할 수 있습니다."
        : code === "rate_limited" ? "저장 창이 이미 열려 있습니다. 잠시 후 다시 눌러 주세요."
        : code === "unavailable" || code === "not_granted" ? "이 화면에서는 파일 저장을 쓸 수 없습니다. 페이지를 새 창에서 열어 주세요."
        : "PPT를 만들지 못했습니다: " + ((e && e.message) || code || "알 수 없는 오류");
      msg("genMsg", text, code !== "declined");
    } finally { $("gen").disabled = false; }
  }

  renderOutline();
  loadAuthor();
  $("consultant").addEventListener("input", syncAuthor);
  $("brand").addEventListener("input", syncAuthor);
  $("footer").addEventListener("input", () => { footerTouched = $("footer").value.trim() !== B.defaultFooter($("brand").value, $("consultant").value); syncAuthor(); });
  $("footerReset").addEventListener("click", () => { footerTouched = false; syncAuthor(); });
  $("file").addEventListener("change", (e) => handle(e.target.files[0]));
  const drop = $("drop");
  ["dragenter", "dragover"].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add("drag"); }));
  ["dragleave", "drop"].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.remove("drag"); }));
  drop.addEventListener("drop", (e) => handle(e.dataTransfer.files[0]));
  $("gen").addEventListener("click", generate);
})();
