/* index.html(+src/*)을 한 파일로 묶는다.
 *   dist/기업진단_PPT_생성기.html : 더블클릭으로 여는 단일 HTML (라이브러리는 CDN)
 *   dist/artifact.html            : Claude 아티팩트 게시용 본문(문서 골격 없이)
 * index.html은 GitHub Pages에서 그대로 동작하므로 빌드 없이도 쓸 수 있다.
 */
"use strict";
const fs = require("fs");
const path = require("path");
const root = path.join(__dirname, "..");
const read = (p) => fs.readFileSync(path.join(root, p), "utf8");

const index = read("index.html");
const css = read("src/styles.css");
const core = read("src/core.js");
const app = read("src/app.js");

const inlined = index
  .replace('<link rel="stylesheet" href="src/styles.css">', () => `<style>\n${css}</style>`)
  .replace('<script src="src/core.js"></script>', () => `<script>\n${core}</script>`)
  .replace('<script src="src/app.js"></script>', () => `<script>\n${app}</script>`);
for (const must of ["<style>", "BizReport", "renderOutline"]) if (!inlined.includes(must)) throw new Error("inline failed: " + must);

const bodyStart = inlined.indexOf("<body>") + "<body>".length;
const bodyEnd = inlined.lastIndexOf("</body>");
const headInner = inlined.slice(inlined.indexOf("<head>") + 6, inlined.indexOf("</head>"))
  .replace(/<meta [^>]*>\s*/g, "");
const fragment = headInner.trim() + "\n" + inlined.slice(bodyStart, bodyEnd).trim() + "\n";

fs.mkdirSync(path.join(root, "dist"), { recursive: true });
fs.writeFileSync(path.join(root, "dist", "기업진단_PPT_생성기.html"), inlined);
fs.writeFileSync(path.join(root, "dist", "artifact.html"), fragment);
console.log("built dist/기업진단_PPT_생성기.html, dist/artifact.html");
