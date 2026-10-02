/* 테스트용 가상 기업종합보고서(xlsx) 생성기
 * 실제 고객 보고서는 개인정보·영업정보가 있어 저장소에 넣지 않습니다.
 * 이 스크립트는 한국평가데이터 기업종합보고서와 같은 배치의 "가상 기업" 파일을 만듭니다.
 *   node scripts/make-fixtures.js            → tests/fixtures/*.xlsx
 */
"use strict";
const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");

const OUT = path.join(__dirname, "..", "tests", "fixtures");

// 연도별 기본 재무 데이터(단위: 천원) — 가상 수치
const BASE = {
  2023: { sales: 3200000, cost: 2300000, sga: 420000, nonOpInc: 12000, nonOpExp: 8000, tax: 60000, cash: 520000, ar: 410000, prepaid: 90000, karyo: 0, deposit: 50000, machine: 380000, apay: 260000, ltBorrow: 600000, capital: 100000, salary: 310000, mcLabor: 280000, retire: 40000, insurance: 14000, entertain: 18000, rnd: 0, vehicle: 70000 },
  2024: { sales: 4100000, cost: 2950000, sga: 520000, nonOpInc: 15000, nonOpExp: 9000, tax: 85000, cash: 610000, ar: 520000, prepaid: 120000, karyo: 85000, deposit: 50000, machine: 520000, apay: 310000, ltBorrow: 500000, capital: 100000, salary: 360000, mcLabor: 330000, retire: 46000, insurance: 16000, entertain: 21000, rnd: 35000, vehicle: 70000 },
  2025: { sales: 4800000, cost: 3350000, sga: 610000, nonOpInc: 18000, nonOpExp: 7000, tax: 120000, cash: 450000, ar: 690000, prepaid: 260000, karyo: 210000, deposit: 80000, machine: 640000, apay: 340000, ltBorrow: 400000, capital: 100000, salary: 420000, mcLabor: 360000, retire: 52000, insurance: 19000, entertain: 24000, rnd: 48000, vehicle: 95000 },
};

function derive(years) {
  // 연도별 손익·재무상태를 계산해 서로 맞물리게 만든다(이익잉여금 = 전기 + 당기순이익)
  let re = 150000;
  const out = {};
  for (const y of years) {
    const b = BASE[y];
    const op = b.sales - b.cost - b.sga;
    const pretax = op + b.nonOpInc - b.nonOpExp;
    const ni = pretax - b.tax;
    re += ni;
    const curAssets = b.cash + b.ar + b.prepaid + b.karyo;
    const nonCur = b.deposit + b.machine + b.vehicle;
    const assets = curAssets + nonCur;
    const equity = b.capital + re;
    const curLiab = b.apay + Math.round(b.tax * 0.6);
    const liab = curLiab + b.ltBorrow;
    // 자산 = 부채 + 자본이 되도록 맞춘다(모자라면 현금, 남으면 미지급금)
    const gap = liab + equity - assets;
    const addCash = Math.max(0, gap), payable = Math.max(0, -gap);
    out[y] = Object.assign({}, b, { op, pretax, ni, re, cash: b.cash + addCash, payable, curAssets: curAssets + addCash, nonCur,
      assets: assets + addCash, equity, curLiab: curLiab + payable, liab: liab + payable });
  }
  return out;
}

const k = (v) => (v === null || v === undefined ? "-" : v === 0 ? "-" : String(v) + ".0");
const r2 = (v) => (v === null || !isFinite(v) ? "-" : (Math.round(v * 100) / 100).toString());

function build({ name, years, est, emptyLeadYear = null, file }) {
  const allCols = emptyLeadYear ? [emptyLeadYear, ...years] : years;
  const D = derive(years);
  const val = (y, f) => (D[y] ? f(D[y]) : null);
  const row = (label, f) => [label, ...allCols.map((y) => k(val(y, f)))];
  const dates = allCols.map((y) => `${y}-12-31`);
  const last = years[years.length - 1];
  const pages = [];
  const page = (rows) => pages.push([["조회일시 : 2026-10-02 10:00:00"], ...rows, [`COPYRIGHT 2022 BY KOREA RATING & DATA. ALL RIGHTS RESERVED.`, `${pages.length + 1}/20`]]);

  page([["기업 종합 보고서"], ["- 기업명 :", name], ["- 사업자번호 :", "000-00-00000"], ["- 대표자 :", "홍길동"]]);
  page([
    ["기업개요"],
    ["기업명", name, "영문기업명", "SAMPLE PRECISION"],
    ["사업자번호", "000-00-00000", "법인(주민)번호", "000000-0000000"],
    ["대표자명", "홍길동", "종업원수", "12명"],
    ["설립형태", "신규설립(개업)", "설립년월", est],
    ["기업유형", "일반법인", "기업규모", "소기업"],
    ["결산월", "12월"],
    ["주소", "(00000) 충남 가상시 예시면 샘플로 1"],
    ["표준산업분류(10차)", "(C29199) 그 외 기타 일반 목적용 기계 제조업"],
    ["주요제품(상품)", "산업용 정밀부품"],
    ["주채권기관", "가상은행"],
    ["기업신용등급", "EW 등급"],
    [`결산일자 : ${last}-12-31`],
  ]);
  page([
    ["신용정보"],
    ["단기연체정보(신용정보원)", "신용도판단정보", "공공정보", "상거래연체정보", "당좌거래정지 발생이력"],
    ["해당사항없음", "해당사항없음", "해당사항없음", "해당사항없음", "해당사항없음"],
    ["휴폐업정보", "법인등기정보", "회생절차/워크아웃정보", "행정처분정보", "당좌개설/카드발급 정보"],
    ["부가가치세 일반과세자", "정상", "해당사항없음", "해당사항없음", "해당사항없음"],
  ]);
  page([
    ["기술력", "기업인증・산업재산권 현황"],
    ["기업인증"],
    ["벤처", "이노비즈", "메인비즈", "연구개발전담부서", "부설연구소"],
    ["미인증", "미인증", "미인증", "인증", "미인증"],
    ["산업재산권"],
    ["특허", "실용신안", "디자인", "상표권"],
    ["평가기준연도 정보없음"],
    ["2건", "0건", "0건", "1건"],
  ]);
  // 재무진단(업종 비교)
  const ratio = (y) => {
    const t = D[y]; if (!t) return null;
    const p = D[y - 1];
    return {
      "매출액증가율": p ? (t.sales / p.sales - 1) * 100 : null,
      "영업이익증가율": p ? (t.op / p.op - 1) * 100 : null,
      "총자산증가율": p ? (t.assets / p.assets - 1) * 100 : null,
      "매출원가율": t.cost / t.sales * 100,
      "영업이익률": t.op / t.sales * 100,
      "EBITDA마진율": (t.op + 60000) / t.sales * 100,
      "총자산순이익률(ROA)": t.ni / t.assets * 100,
      "부채비율": t.liab / t.equity * 100,
      "유동비율": t.curAssets / t.curLiab * 100,
      "차입금의존도": t.ltBorrow / t.assets * 100,
      "법인세부담률": t.tax / t.pretax * 100,
      "배당성향": 0,
      "노무비/매출액": t.mcLabor / t.sales * 100,
      "종업원1인당인건비(백만원)": (t.salary + t.mcLabor + t.retire) / 1000 / 12,
      "매출채권회전율(회)": t.sales / t.ar,
      "총자본회전율(회)": t.sales / t.assets,
    };
  };
  const R = Object.fromEntries(years.map((y) => [y, ratio(y)]));
  const ratioRow = (label) => ["성장성", label, ...allCols.map((y) => r2(R[y] ? R[y][label] : null))];
  page([
    ["요약 재무상태표", "단위 : 백만원"],
    ["구분", ...allCols.map(String)],
    ["자산총계", ...allCols.map((y) => (D[y] ? Math.round(D[y].assets / 1000).toLocaleString("en-US") : "-"))],
    ["부채총계", ...allCols.map((y) => (D[y] ? Math.round(D[y].liab / 1000).toLocaleString("en-US") : "-"))],
    ["자본총계", ...allCols.map((y) => (D[y] ? Math.round(D[y].equity / 1000).toLocaleString("en-US") : "-"))],
  ]);
  page([
    ["동종업계내 경영규모 비교", "기준년도 : " + last, "단위 : 백만원"],
    ["구분", "총자산", "자본총계", "납입자본금", "매출액", "영업이익", "당기순이익"],
    ["조회기업", ...[D[last].assets, D[last].equity, D[last].capital, D[last].sales, D[last].op, D[last].ni].map((v) => String(Math.round(v / 1000)) + ".0")],
    ["상위25%", "9500.0", "5200.0", "500.0", "8800.0", "620.0", "510.0"],
    ["평균", "6400.0", "3500.0", "320.0", "5600.0", "330.0", "270.0"],
    ["하위25%", "1800.0", "700.0", "100.0", "1500.0", "20.0", "10.0"],
    ["업계순위"],
    ["412위", name, String(D[last].sales / 1000) + ".0", "12월", "000-00-00000", "홍길동"],
  ]);
  page([
    ["경영진현황", "기준일자 : 2026-08-01"],
    ["구분", "성명", "등기여부", "직위", "담당업무", "근속년수", "경영실권자와의 관계", "주식소유현황"],
    ["대표이사및사원", "홍길동", "등기", "대표이사", "총괄", "본인"],
    ["사내이사", "김예시", "등기", "사내이사", "생산", "타인"],
    ["감사", "홍예시", "등기", "감사", "관리", "가족"],
    ["주요주주현황", "기준일자 : 2026-08-01", "단위 : 주, %"],
    ["주주명", "구분", "소유주식수", "지분율", "경영실권자와의 관계", "회사와의 관계"],
    ["보통주", "우선주", "합계", "보통주", "우선주", "합계"],
    ["홍길동", "최대주주", "12000.0", "0.0", "12000.0", "60.0", "0.0", "60.0", "본인", "대표이사"],
    ["홍예시", "5%이상주주", "8000.0", "0.0", "8000.0", "40.0", "0.0", "40.0", "가족", "임원"],
    ["관계회사현황", "기준일자 : 2026-08-01", "단위 : 백만원, %"],
    ["기업명", "사업내용", "관계내용", "지분율", "결산년도"],
    ["해당사항 없음"],
  ]);
  page([
    ["사업장 현황"],
    ["자가소유 여부", "비소유", "실제가동 여부", "실제가동"],
    ["담보제공여부", "담보비제공"],
    ["거래처현황"],
    ["구매처현황", "기준일자 : 2026-08-01", "단위 : 백만원, %"],
    ["기업명", "사업자번호", "대표자명", "거래비중"],
    ["가상소재(주)", "000-00-00001", "김예시", "41.5"],
    ["예시철강", "000-00-00002", "박샘플", "22.0"],
    ["기타", "36.5"],
    ["판매처현황", "단위 : 백만원, %"],
    ["기업명", "사업자번호", "대표자명", "거래비중"],
    ["(주)샘플모터스", "000-00-00003", "최가상", "58.2"],
    ["가상기계(주)", "000-00-00004", "정예시", "17.4"],
    ["기타", "24.4"],
    ["매출구성", "기준일자 : 2026-08-01"],
  ]);
  page([
    ["사업목적"],
    ["내용"],
    ["1. 산업용 정밀부품 제조업"],
    ["2. 기계 가공업"],
    ["3. 도소매업"],
    ["4. 각 호에 관련된 부대사업 일체"],
    ["종합의견"],
    ["연혁"],
    ["연혁일자", "내용"],
    [est, "법인설립"],
  ]);
  // 상세 재무제표(천원)
  page([
    ["재무상태표", "단위 : 천원"],
    ["계정명", ...dates],
    ["감사의견"],
    row("자산(*)", (t) => t.assets),
    row("유동자산(*)", (t) => t.curAssets),
    row("현금 및 현금성자산(*)", (t) => t.cash),
    row("매출채권(*)", (t) => t.ar),
    row("선급금(*)", (t) => t.prepaid),
    row("가지급금", (t) => t.karyo),
    row("비유동자산(*)", (t) => t.nonCur),
    row("기계장치", (t) => t.machine),
    row("차량운반구", (t) => t.vehicle),
    row("보증금 등", (t) => t.deposit),
    row("부채(*)", (t) => t.liab),
    row("유동부채(*)", (t) => t.curLiab),
    row("매입채무(*)", (t) => t.apay),
    row("미지급금", (t) => t.payable),
    row("미지급법인세", (t) => Math.round(t.tax * 0.6)),
    row("비유동부채(*)", (t) => t.ltBorrow),
    row("장기차입금(*)", (t) => t.ltBorrow),
    row("자본(*)", (t) => t.equity),
    row("자본금(*)", (t) => t.capital),
    row("이익잉여금(*)", (t) => t.re),
  ]);
  page([
    ["손익계산서", "단위 : 천원"],
    ["계정명", ...dates],
    ["감사의견"],
    row("매출액(*)", (t) => t.sales),
    row("매출원가(*)", (t) => t.cost),
    row("매출총이익(손실)", (t) => t.sales - t.cost),
    row("판매비와관리비(*)", (t) => t.sga),
    row("급여(*)", (t) => t.salary),
    row("퇴직급여", (t) => t.retire),
    row("보험료", (t) => t.insurance),
    row("접대비", (t) => t.entertain),
    row("경상개발비", (t) => t.rnd),
    row("영업이익(손실)", (t) => t.op),
    row("영업외수익(*)", (t) => t.nonOpInc),
    row("영업외비용(*)", (t) => t.nonOpExp),
    row("법인세비용차감전순손익", (t) => t.pretax),
    row("법인세비용", (t) => t.tax),
    row("당기순이익(순손실)", (t) => t.ni),
  ]);
  page([
    ["제조원가명세서", "단위 : 천원"],
    ["계정명", ...dates],
    row("노동관계비용(*)", (t) => t.mcLabor),
    row("보험료", () => 6000),
  ]);
  page([
    ["재무비율", "단위 : %"],
    ["구분", "계정명", ...dates],
    ...["매출액증가율", "영업이익증가율", "총자산증가율", "매출원가율", "영업이익률", "EBITDA마진율", "총자산순이익률(ROA)", "부채비율", "유동비율", "차입금의존도", "법인세부담률", "배당성향", "노무비/매출액", "종업원1인당인건비(백만원)", "매출채권회전율(회)", "총자본회전율(회)"].map(ratioRow),
  ]);
  const IND = { "매출액증가율": 4.1, "영업이익률": 5.8, "EBITDA마진율": 8.9, "총자산순이익률(ROA)": 4.2, "부채비율": 160.3, "유동비율": 150.2, "차입금의존도": 35.4, "매출채권회전율(회)": 6.1, "총자본회전율(회)": 1.1 };
  const diag = Object.entries(IND).map(([kk, avg]) => {
    const c = R[last] && R[last][kk], p = R[last - 1] && R[last - 1][kk];
    const dir = c === null || c === undefined || p === null || p === undefined ? "-" : c >= p ? "▲ 증가" : "▼ 감소";
    return [kk, String(avg), dir, r2(c), kk, ...allCols.map((y) => r2(R[y] ? R[y][kk] : null))];
  });
  page([
    ["재무진단"],
    ["성장성", "수익성", "재무구조", "부채상환능력", "활동성"],
    ["성장성", "기준일자 : " + last + "-12-31"],
    ["성장 역량은 양호함"],
    ["구분", "업종평균", "전년대비", "조회기업", "구분", ...dates],
    ...diag,
  ]);

  const wb = XLSX.utils.book_new();
  pages.forEach((rows, i) => {
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = Array.from({ length: 12 }, () => ({ wch: 24 })); // PDF로 바꿀 때 칸 사이가 붙지 않도록
    XLSX.utils.book_append_sheet(wb, ws, `Page ${i + 1}`);
  });
  fs.mkdirSync(OUT, { recursive: true });
  XLSX.writeFile(wb, path.join(OUT, file));
  return file;
}

const NAME = "(주)샘플정밀";
const made = [
  build({ name: NAME, years: [2023, 2024, 2025], est: "2019-06-01", file: "sample_3y.xlsx" }),
  build({ name: NAME, years: [2024, 2025], est: "2024-05-10", file: "sample_2y.xlsx" }),
  build({ name: NAME, years: [2025], est: "2025-03-02", file: "sample_1y.xlsx" }),
  // 보고서에는 3개 열이 있으나 2023년은 설립 전이라 비어 있는 경우
  build({ name: NAME, years: [2024, 2025], est: "2024-02-15", emptyLeadYear: 2023, file: "sample_2y_emptycol.xlsx" }),
];
console.log("fixtures:", made.join(", "));
