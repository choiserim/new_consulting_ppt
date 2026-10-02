"use strict";
/* 문서 일관성 테스트 — docs/CONSISTENCY.md의 규칙을 자동으로 확인한다.  실행: npm test */
const test = require("node:test");
const assert = require("node:assert/strict");
const { B, analyzeFile, deckOf } = require("./helpers");

const AUTHOR = { consultant: "최세림 GFC", brand: "삼성센터법인지점" };
const FILES = ["sample_3y.xlsx", "sample_2y.xlsx", "sample_1y.xlsx", "sample_2y_emptycol.xlsx"];
const decks = {};
const get = async (f) => (decks[f] = decks[f] || (await deckOf(f, AUTHOR)));

test("규칙 1: 재무 연도는 기본 3개년, 설립 연도에 따라 2개년·1개년", async () => {
  const want = { "sample_3y.xlsx": ["2023", "2024", "2025"], "sample_2y.xlsx": ["2024", "2025"], "sample_1y.xlsx": ["2025"], "sample_2y_emptycol.xlsx": ["2024", "2025"] };
  for (const [f, years] of Object.entries(want)) {
    const { A } = await analyzeFile(f);
    assert.deepEqual(A.years, years, f);
    for (const [k, v] of Object.entries(A.F)) if (Array.isArray(v)) assert.equal(v.length, years.length, `${f} ${k} 길이`);
  }
});

test("규칙 2: 모든 연도 차트의 가로축은 같은 표시 연도를 쓴다", async () => {
  for (const f of FILES) {
    const { A, slides } = await get(f);
    for (const s of slides) for (const c of s.charts) {
      if (c.cats.some((x) => /^\d{4}$/.test(x))) {
        if (/법인세 부담률/.test(c.name)) assert.ok(c.cats.every((y) => A.years.includes(y)), `${f} ${c.name}`);
        else assert.deepEqual(c.cats, A.years, `${f} ${c.name}`);
      }
    }
  }
});

test("규칙 3: 같은 수치는 모든 슬라이드에서 같은 값으로 표기한다", async () => {
  for (const f of FILES) {
    const { A, slides } = await get(f);
    const sales = B.fmt(A.F.sales[A.F.sales.length - 1]);
    assert.ok(slides[2].text.includes(`매출 ${sales}백만`), `${f} 종합진단 매출`);
    assert.ok(slides[3].text.includes(sales), `${f} 재무① 매출`);
    const salesChart = slides[3].charts.find((c) => c.name === "매출액");
    assert.deepEqual(salesChart.vals, A.F.sales.map((v) => Math.round(v)), `${f} 매출 차트`);
    const assetChart = slides[4].charts.find((c) => c.name === "자산총계");
    assert.deepEqual(assetChart.vals, A.F.assets.map((v) => Math.round(v)), `${f} 자산 차트`);
  }
});

test("규칙 4: 작성자 정보는 표지·마지막 장·바닥글에 같은 이름으로 들어간다", async () => {
  const { slides, layoutsXml } = await get("sample_3y.xlsx");
  const byline = "삼성센터법인지점  |  최세림 GFC";
  assert.ok(slides[0].text.includes(byline), "표지");
  assert.ok(slides[slides.length - 1].text.includes(byline), "마지막 장");
  assert.ok(layoutsXml.includes("삼성센터법인지점 최세림 GFC"), "바닥글 = 소속 + 컨설턴트");
  const o = B.normalizeAuthor({ consultant: "홍길동 GFC", brand: "가상지점", footer: "" });
  assert.equal(o.footer, "가상지점 홍길동 GFC");
  assert.equal(B.normalizeAuthor({ consultant: "홍길동 GFC", brand: "가상지점", footer: "직접 입력" }).footer, "직접 입력");
});

test("규칙 5: 엑셀과 PDF로 올린 같은 보고서는 같은 PPT 내용이 나온다", async () => {
  const x = await deckOf("sample_3y.xlsx", AUTHOR);
  const p = await deckOf("sample_3y.pdf", AUTHOR);
  assert.equal(p.slides.length, x.slides.length);
  x.slides.forEach((s, i) => {
    assert.equal(p.slides[i].text, s.text, `슬라이드 ${i + 1} 글자`);
    assert.deepEqual(p.slides[i].charts, s.charts, `슬라이드 ${i + 1} 차트`);
  });
});

test("규칙 6: 증감 표기는 ▲=증가, ▼=감소로만 쓰고 계산 오류 글자가 없다", async () => {
  for (const f of FILES) {
    const { slides } = await get(f);
    slides.forEach((s, i) => {
      assert.ok(!/▲ 감소|▼ 증가/.test(s.text), `${f} 슬라이드 ${i + 1} 방향 모순`);
      assert.ok(!/undefined|NaN|Infinity|null|\[object/.test(s.text), `${f} 슬라이드 ${i + 1}: ${(s.text.match(/.{0,20}(undefined|NaN|Infinity|null).{0,20}/) || [""])[0]}`);
    });
  }
});

test("규칙 7: 1개년 보고서는 전년 대비를 만들지 않는다", async () => {
  const { slides } = await get("sample_1y.xlsx");
  assert.ok(!/전년 대비는 \d{4}→/.test(slides[5].text));
  assert.ok(!/\d{4}년 대비 [\d.]+배/.test(slides[3].text));
  assert.ok(slides[2].text.includes("1개년(2025)") || slides[2].text.includes("재무 1개년"));
});

test("규칙 8: 같은 입력이면 매번 같은 결과(결정성)", async () => {
  const a = await deckOf("sample_2y.xlsx", AUTHOR);
  const b = await deckOf("sample_2y.xlsx", AUTHOR);
  assert.deepEqual(a.slides, b.slides);
});

test("규칙 9: 15장 구성과 순서는 입력과 관계없이 같다", async () => {
  const chips = ["01  기업 개요", "02  종합 진단", "03  재무 ①", "03  재무 ②", "03  재무 ③", "04  세무 ①", "04  세무 ②", "05  노무", "06  보험", "07  정관", "08  인증", "09  법무", "10  실행 로드맵"];
  for (const f of FILES) {
    const { slides } = await get(f);
    assert.equal(slides.length, 15, f);
    chips.forEach((c, i) => assert.ok(slides[i + 1].text.startsWith(c), `${f} ${i + 2}장: ${c}`));
  }
});

test("파일 이름: 기업명에서 (주)·주식회사·공백을 뺀다", () => {
  assert.equal(B.fileNameFor({ info: { "기업명": "(주)샘플정밀" } }), "샘플정밀_경영종합진단_컨설팅.pptx");
  assert.equal(B.fileNameFor({ info: { "기업명": "주식회사 가상 테크" } }), "가상테크_경영종합진단_컨설팅.pptx");
});

test("숫자 읽기: 콤마·△·괄호 음수·단위", () => {
  assert.equal(B.num("1,234"), 1234);
  assert.equal(B.num("△1,032,684"), -1032684);
  assert.equal(B.num("(500)"), -500);
  assert.equal(B.num("26명"), 26);
  assert.equal(B.num("-"), null);
});
