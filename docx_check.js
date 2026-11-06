"use strict";
/* Word download harness: runs the real converter and walker from
   index.html in a stub DOM, checks representative equations, censuses
   fallbacks over a full-bank sample draw, then builds a genuine
   sample-word.docx from a real practice sheet.
   Run: node docx_check.js */
const fs = require("fs"), vm = require("vm"), path = require("path");
const APP = __dirname;
const katex = require(path.join(APP, "katex", "katex.min.js"));
/* ---------- minimal markup parsers (stub DOM for the harness) ---------- */
const ENT = {
  amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", nbsp: "\u00A0",
  minus: "\u2212", hellip: "\u2026", mdash: "\u2014", ndash: "\u2013", times: "\u00D7",
  deg: "\u00B0", middot: "\u00B7", le: "\u2264", ge: "\u2265", ne: "\u2260", alpha: "\u03B1",
  beta: "\u03B2", pi: "\u03C0", theta: "\u03B8", Delta: "\u0394", infty: "\u221E", rarr: "\u2192",
};
function decodeEnt(s) {
  return s.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z][a-zA-Z0-9]*);/g, (m, g) => {
    if (g[0] === "#") {
      const code = (g[1] === "x" || g[1] === "X") ? parseInt(g.slice(2), 16) : parseInt(g.slice(1), 10);
      return String.fromCodePoint(code);
    }
    return Object.prototype.hasOwnProperty.call(ENT, g) ? ENT[g] : m;
  });
}
function mkNode(tagLower, attrs, htmlMode) {
  const el = {
    nodeType: 1,
    localName: tagLower,
    tagName: htmlMode ? tagLower.toUpperCase() : tagLower,
    childNodes: [],
    attributes: attrs,
    getAttribute(name) {
      const a = attrs.find((x) => x.name === name.toLowerCase());
      return a ? a.value : null;
    },
  };
  Object.defineProperty(el, "textContent", { get() { return collectText(el); } });
  return el;
}
function makeText(v) { return { nodeType: 3, nodeValue: v, childNodes: [] }; }
function collectText(n) {
  if (n.nodeType === 3) return n.nodeValue;
  return (n.childNodes || []).map(collectText).join("");
}
const VOID = { br: 1, img: 1, input: 1, hr: 1, meta: 1, link: 1, col: 1, area: 1, base: 1, wbr: 1, source: 1 };
function parseMarkup(str, htmlMode) {
  const root = mkNode("#root", [], htmlMode);
  const stack = [root];
  const top = () => stack[stack.length - 1];
  let i = 0;
  const n = str.length;
  while (i < n) {
    const lt = str.indexOf("<", i);
    if (lt < 0) { top().childNodes.push(makeText(decodeEnt(str.slice(i)))); break; }
    if (lt > i) top().childNodes.push(makeText(decodeEnt(str.slice(i, lt))));
    if (str.startsWith("<!--", lt)) {
      const end = str.indexOf("-->", lt);
      i = end < 0 ? n : end + 3;
      continue;
    }
    if (str.startsWith("<!", lt) || str.startsWith("<?", lt)) {
      const end = str.indexOf(">", lt);
      i = end < 0 ? n : end + 1;
      continue;
    }
    if (str.startsWith("</", lt)) {
      const end = str.indexOf(">", lt);
      const name = str.slice(lt + 2, end < 0 ? n : end).trim().toLowerCase();
      for (let s = stack.length - 1; s > 0; s--) {
        if (stack[s].localName === name) { stack.length = s; break; }
      }
      i = end < 0 ? n : end + 1;
      continue;
    }
    /* open tag */
    let j = lt + 1;
    let name = "";
    while (j < n && /[A-Za-z0-9_:.-]/.test(str[j])) { name += str[j]; j++; }
    name = name.toLowerCase();
    const attrs = [];
    let selfClose = false;
    while (j < n) {
      while (j < n && /\s/.test(str[j])) j++;
      if (str[j] === ">") { j++; break; }
      if (str[j] === "/" && str[j + 1] === ">") { selfClose = true; j += 2; break; }
      let an = "";
      while (j < n && /[^\s=/>]/.test(str[j])) { an += str[j]; j++; }
      let av = "";
      while (j < n && /\s/.test(str[j])) j++;
      if (str[j] === "=") {
        j++;
        while (j < n && /\s/.test(str[j])) j++;
        const q = str[j];
        if (q === '"' || q === "'") {
          const e = str.indexOf(q, j + 1);
          av = str.slice(j + 1, e < 0 ? n : e);
          j = e < 0 ? n : e + 1;
        } else {
          const start = j;
          while (j < n && /[^\s>]/.test(str[j])) j++;
          av = str.slice(start, j);
        }
      }
      if (an) attrs.push({ name: an.toLowerCase(), value: decodeEnt(av) });
      if (j >= n) break;
    }
    const el = mkNode(name, attrs, htmlMode);
    top().childNodes.push(el);
    if (!selfClose && !(htmlMode && VOID[name])) {
      if (htmlMode && (name === "script" || name === "style")) {
        const close = str.toLowerCase().indexOf("</" + name + ">", j);
        i = close < 0 ? n : close + name.length + 3;
        continue;
      }
      stack.push(el);
    }
    i = j;
  }
  return root;
}
function parseXml(str) { return parseMarkup(str, false); }
function parseHtml(str) { return parseMarkup(str, true); }
/* Tiny stack-based XML well-formedness checker. */
function xmlErrors(xml) {
  const errs = [];
  const stack = [];
  const re = /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[[\s\S]*?\]\]>|<![^>]*>|<\/([A-Za-z_][\w:.-]*)\s*>|<([A-Za-z_][\w:.-]*)((?:[^>"']|"[^"]*"|'[^']*')*?)(\/?)>|([^<]+)/g;
  let m, lastEnd = 0;
  while ((m = re.exec(xml))) {
    lastEnd = re.lastIndex;
    if (m[5] !== undefined) {
      if (/&(?!amp;|lt;|gt;|quot;|apos;|#)/.test(m[5])) errs.push("raw & in text near: " + m[5].slice(0, 30));
      continue;
    }
    if (m[1]) {
      const open = stack.pop();
      if (open !== m[1]) errs.push("close mismatch: expected " + open + ", got " + m[1]);
    } else if (m[2]) {
      if (!m[4]) stack.push(m[2]);
    }
  }
  if (lastEnd !== xml.length) errs.push("unparsed tail at " + lastEnd);
  if (stack.length) errs.push("unclosed: " + stack.join(","));
  return errs;
}



/* ---------- stub DOM, same pattern as print_test.js ---------- */
const els = {}, listeners = {};
function parseChks(htmlStr, cls) {
  if (!htmlStr) return [];
  const re = new RegExp('<input[^>]*class="' + cls + '"[^>]*>', "g");
  const out = []; let m;
  while ((m = re.exec(htmlStr))) {
    const tag = m[0];
    out.push({ value: (tag.match(/value="([^"]*)"/) || [])[1], checked: / checked/.test(tag), addEventListener() {} });
  }
  return out;
}
function makeEl(id) {
  const e = { id, children: [], style: {}, dataset: {}, value: "", checked: false, _inner: "", _text: "",
    classList: { add() {}, remove() {}, toggle() {}, contains: () => false },
    addEventListener(t, f) { listeners[id] = { click: f }; },
    setAttribute() {}, getAttribute() { return null; }, removeAttribute() {},
    appendChild(c) { this.children.push(c); return c; },
    querySelector() { return null; }, querySelectorAll() { return []; } };
  Object.defineProperty(e, "innerHTML", { get() { return this._inner; }, set(v) { this._inner = String(v); } });
  Object.defineProperty(e, "textContent", { get() { return this._text; }, set(v) { this._text = String(v); } });
  return e;
}
["presetMid", "presetFinal", "presetAll", "presetUnit"].forEach(id => {
  const e = makeEl(id); e.dataset.preset = id.replace("preset", "").toLowerCase(); els[id] = e;
});
const documentStub = {
  getElementById(id) { return els[id] || (els[id] = makeEl(id)); },
  querySelectorAll(sel) {
    if (sel === ".preset") return ["presetMid", "presetFinal", "presetAll", "presetUnit"].map(i => els[i]);
    if (sel === ".unitChk") return parseChks(els.unitGrid && els.unitGrid.innerHTML, "unitChk");
    if (sel === ".unitChk:checked") return parseChks(els.unitGrid && els.unitGrid.innerHTML, "unitChk").filter(c => c.checked);
    if (sel === ".indChk") return parseChks(els.indBox && els.indBox.innerHTML, "indChk");
    if (sel === ".indChk:checked") return parseChks(els.indBox && els.indBox.innerHTML, "indChk").filter(c => c.checked);
    return [];
  },
  querySelector() { return null; },
  createElement() { return makeEl("anon"); },
  createTreeWalker() { return { nextNode() { return null; } }; },
  body: makeEl("body"),
};
const sandbox = { console, URLSearchParams, document: documentStub, TextEncoder: require("util").TextEncoder,
  location: { search: "", pathname: "/index.html", origin: "http://test", href: "http://test/index.html" },
  history: { replaceState() {} }, navigator: {}, setTimeout, clearTimeout, print() {}, NodeFilter: { SHOW_TEXT: 4 } };
sandbox.window = sandbox; sandbox.globalThis = sandbox;
vm.createContext(sandbox);
const html = fs.readFileSync(path.join(APP, "index.html"), "utf8");
const blocks = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map(m => m[1]);
blocks.forEach((b, i) => vm.runInContext(b, sandbox, { filename: "block" + i }));
const S = sandbox;
const fails = [];
const ok = (c, m) => { if (!c) fails.push(m); };

/* ---------- helpers ---------- */
function mathElOf(tex, disp) {
  const s = katex.renderToString(tex, { displayMode: !!disp, throwOnError: true });
  const root = parseXml(s.slice(s.indexOf("<math"), s.indexOf("</math>") + 7));
  return root.childNodes.find(c => c.nodeType === 1);
}
function typesetHtml(htmlStr) {
  return htmlStr
    .replace(/\\\[([\s\S]*?)\\\]/g, (m, tex) => katex.renderToString(tex, { displayMode: true, throwOnError: true }))
    .replace(/\\\(([\s\S]*?)\\\)/g, (m, tex) => katex.renderToString(tex, { displayMode: false, throwOnError: true }));
}
/* All Word text runs of a document, concatenated in order. */
function wtText(xml) {
  return [...xml.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)].map(m => m[1]).join("");
}

/* ---------- 1. generator functions exist ---------- */
ok(typeof S.mathToOmml === "function", "mathToOmml missing");
ok(typeof S.makeWordDocx === "function", "makeWordDocx missing");
ok(typeof S.downloadWordFile === "function", "downloadWordFile missing");
ok(typeof S.zipStore === "function", "zipStore missing");
{
  const z = S.zipStore([{ name: "a.txt", data: new sandbox.TextEncoder().encode("hello") }]);
  ok(z[0] === 80 && z[1] === 75, "zipStore output does not start with PK");
}

/* ---------- 2. representative equations ---------- */
const reps = [
  ["stacked fraction", "\\frac{x+3}{x-2}", false, ["<m:f>"]],
  ["radical", "\\sqrt{x-3}", false, ["<m:rad>", '<m:degHide m:val="1"/>']],
  ["cube root", "\\sqrt[3]{x}", false, ["<m:rad>"]],
  ["limit with under-script", "\\lim_{x\\to 3^{-}} f(x)", true, ["<m:limLow>"]],
  ["aligned solution", "\\begin{aligned} f'(x) &= 3x^2+2x \\\\ &= x(3x+2) \\end{aligned}", true, ["<m:eqArr>"]],
  ["superscript", "x^2", false, ["<m:sSup>"]],
  ["subscript", "x_i", false, ["<m:sSub>"]],
  ["sub+sup", "x_i^2", false, ["<m:sSubSup>"]],
  ["integral", "\\int_0^1 x^2\\,dx", true, ["<m:nary>", '<m:chr m:val="∫"/>']],
  ["sum", "\\sum_{n=1}^{\\infty}\\frac{1}{n^2}", true, ["<m:nary>", '<m:limLoc m:val="undOvr"/>']],
];
let repOmml = "";
for (const [name, tex, disp, wants] of reps) {
  const o = S.mathToOmml(mathElOf(tex, disp), { fb: {} });
  repOmml += o;
  for (const w of wants) ok(o.includes(w), name + ": missing " + w + " in " + o.slice(0, 160));
  const errs = xmlErrors(S.docxDocumentXml("<w:p><m:oMath>" + o + "</m:oMath></w:p>"));
  ok(errs.length === 0, name + ": document.xml not well-formed: " + errs.join("; "));
}
ok(!repOmml.includes("annotation"), "annotation leaked into OMML");
ok(!/\\[a-zA-Z]{2,}/.test(repOmml), "raw TeX command leaked into OMML: " + (repOmml.match(/\\[a-zA-Z]{2,}/) || [""])[0]);

/* ---------- 3. full-bank fallback census ---------- */
function collectStrings(v, out) {
  if (typeof v === "string") out.push(v);
  else if (Array.isArray(v)) v.forEach(x => collectStrings(x, out));
  else if (v && typeof v === "object") Object.values(v).forEach(x => collectStrings(x, out));
}
{
  const seen = new Set(), st = { fb: {} };
  let segs = 0, renderErrs = 0;
  for (const e of S.BANK) {
    for (let t = 0; t < 25; t++) {
      let g; try { g = e.gen(); } catch (err) { continue; }
      const strs = []; collectStrings(g, strs);
      for (const s of strs) {
        const re = /\\\(([\s\S]*?)\\\)|\\\[([\s\S]*?)\\\]/g; let m;
        while ((m = re.exec(s))) {
          const tex = m[1] !== undefined ? m[1] : m[2], disp = m[1] === undefined;
          const key = (disp ? "D:" : "I:") + tex;
          if (seen.has(key)) continue;
          seen.add(key); segs++;
          try { S.mathToOmml(mathElOf(tex, disp), st); } catch (err) { renderErrs++; }
        }
      }
    }
  }
  const fbTotal = Object.values(st.fb).reduce((a, b) => a + b, 0);
  console.log("bank census: " + segs + " unique segments, converter errors " + renderErrs + ", fallbacks " + JSON.stringify(st.fb));
  ok(renderErrs === 0, "converter threw on " + renderErrs + " bank segments");
  ok(fbTotal === 0, "fallback-degraded elements in bank output: " + JSON.stringify(st.fb));
}

/* ---------- 4. real sheet -> sample-word.docx ---------- */
const el = id => documentStub.getElementById(id);
el("countIn").value = "10"; el("lv1").value = "20"; el("lv2").value = "60"; el("lv3").value = "20";
el("wsIn").value = "1"; el("orderIn").value = "unit"; el("chkIn").value = "each";
el("mcCountIn").value = "6"; el("ncCountIn").value = "4"; el("keyChk").checked = true;
(listeners.presetMid || {}).click && listeners.presetMid.click();
(listeners.goSheet || {}).click && listeners.goSheet.click();
const rawSheet = els.printArea.innerHTML;
ok(rawSheet.includes("Math 3208: Practice Sheet"), "sheet did not render for the sample");
const rootEl = parseHtml(typesetHtml(rawSheet));
const stDoc = { fb: {} };
const docXml = S.docxDocumentXml(S.docxBodyXml(rootEl, stDoc));
const docErrs = xmlErrors(docXml);
ok(docErrs.length === 0, "sample document.xml not well-formed: " + docErrs.join("; "));
ok(docXml.includes("Math 3208: Practice Sheet"), "sample missing the sheet title");
ok(docXml.includes("Name:"), "sample missing the Name line");
ok((docXml.match(/<m:oMath>/g) || []).length >= 20, "sample has too few equations: " + (docXml.match(/<m:oMath>/g) || []).length);
ok(docXml.includes("<m:oMathPara>"), "sample missing a display equation paragraph");
ok(docXml.includes('<w:br w:type="page"/>'), "sample missing the answer key page break");
ok(docXml.includes("Answer Key, with worked solutions"), "sample missing the answer key");
ok(docXml.includes("<w:pBdr>"), "sample missing a workspace box paragraph");
ok(!docXml.includes("○"), "sheet document.xml still carries an open-circle option marker");
ok(/Answer: [A-D]\. /.test(wtText(docXml)), "sheet key cites no option letter in any Answer line");
ok(!docXml.includes("annotation"), "sample document.xml contains annotation");
ok(!/\\[a-zA-Z]{2,}/.test(docXml), "sample document.xml contains a raw TeX command");
console.log("sample doc: oMath " + (docXml.match(/<m:oMath>/g) || []).length +
  ", oMathPara " + (docXml.match(/<m:oMathPara>/g) || []).length +
  ", walker fallbacks " + JSON.stringify(stDoc.fb));
const bytes = S.makeWordDocx(rootEl);
const buf = Buffer.from(bytes);
ok(buf[0] === 0x50 && buf[1] === 0x4b, "makeWordDocx output is not a ZIP");
ok(buf.length > 5000, "sample docx suspiciously small: " + buf.length);
const outPath = path.join(APP, "sample-word.docx");
fs.writeFileSync(outPath, buf);
console.log("wrote " + outPath + " (" + buf.length + " bytes)");

/* ---------- 4b. synthetic sheet: option lettering and key letters, exact ---------- */
{
  const synHtml =
    '<div><h1 style="font-size:1.3rem;">Math 3208: Practice Sheet</h1>' +
    '<div class="pq"><p><b>1.</b></p><div class="prompt">Pick the even number.</div>' +
    '<div class="opts"><label style="display:block;">○ 3</label><label style="display:block;">○ 4</label><label style="display:block;">○ 5</label><label style="display:block;">○ 7</label></div></div>' +
    '<div class="pq"><p><b>2.</b></p><div class="prompt">Pick the root.</div>' +
    '<div class="opts"><label style="display:block;">○ \\(x=1\\)</label><label style="display:block;">○ \\(x=2\\)</label><label style="display:block;">○ \\(x=3\\)</label><label style="display:block;">○ \\(x=4\\)</label></div></div>' +
    '<div class="pq"><p><b>3.</b></p><div class="prompt">Compute the value.</div></div>' +
    '<div class="keybreak"><h1 style="font-size:1.3rem;">Answer Key, with worked solutions</h1>' +
    '<div class="solp"><p><b>1.</b> Answer: <b>4</b></p><div style="line-height:1.5;">Because 4 is even.</div></div>' +
    '<div class="solp"><p><b>2.</b> Answer: <b>\\(x=3\\)</b></p><div style="line-height:1.5;">Substitution.</div></div>' +
    '<div class="solp"><p><b>3.</b> Answer: <b>9</b></p><div style="line-height:1.5;">Workings.</div></div>' +
    '</div></div>';
  const synRoot = parseHtml(typesetHtml(synHtml));
  const synXml = S.docxDocumentXml(S.docxBodyXml(synRoot, { fb: {} }));
  const synErrs = xmlErrors(synXml);
  ok(synErrs.length === 0, "synthetic sheet document.xml not well-formed: " + synErrs.join("; "));
  ok(!synXml.includes("○"), "synthetic sheet still carries an open-circle marker");
  const st = wtText(synXml);
  const iA = st.indexOf("A. 3"), iB = st.indexOf("B. 4"), iC = st.indexOf("C. 5"), iD = st.indexOf("D. 7");
  ok(iA >= 0 && iB > iA && iC > iB && iD > iC, "sheet options not lettered A to D in order: " + st.slice(0, 200));
  ok((st.match(/A\. /g) || []).length >= 2, "math option did not get its A. letter");
  ok(st.includes("Answer: B. 4"), "key line for question 1 does not cite letter B: " + st.slice(st.indexOf("Answer:"), st.indexOf("Answer:") + 40));
  ok(st.includes("Answer: C. "), "key line for the math question does not cite letter C");
  const q2para = synXml.split("<w:p>").find(p => p.includes("Answer:") && p.includes("C. "));
  ok(!!q2para && q2para.indexOf("C. ") < q2para.indexOf("<m:oMath>"), "key letter C does not precede the answer equation");
  ok(st.includes("Answer: 9") && !/Answer: [A-D]\. 9/.test(st), "non-MC key line must stay unlettered");
}

/* ---------- 5. real exam: answer record grids become tables ---------- */
(listeners.presetFinal || {}).click && listeners.presetFinal.click();
(listeners.goExam || {}).click && listeners.goExam.click();
const examRoot = parseHtml(typesetHtml(els.session.innerHTML));
function findById(n, id) {
  if (!n || n.nodeType !== 1) return null;
  if (n.getAttribute && n.getAttribute("id") === id) return n;
  for (const c of n.childNodes || []) { const r = findById(c, id); if (r) return r; }
  return null;
}
const prev = findById(examRoot, "sheetPrev");
ok(!!prev, "exam preview element not found");
if (prev) {
  const stEx = { fb: {} };
  const exXml = S.docxDocumentXml(S.docxBodyXml(prev, stEx));
  const exErrs = xmlErrors(exXml);
  ok(exErrs.length === 0, "exam document.xml not well-formed: " + exErrs.join("; "));
  ok(exXml.includes("Math 3208: Practice Exam"), "exam sample missing its title");
  ok((exXml.match(/<w:tbl>/g) || []).length >= 2, "exam answer record grids missing: " + (exXml.match(/<w:tbl>/g) || []).length);
  ok(exXml.includes('<w:u w:val="single"/>'), "exam key grid letters not underlined");
  ok(exXml.includes("Answer Key"), "exam sample missing the answer key");
  const letterRuns = [...exXml.matchAll(/<w:t[^>]*>([A-D]\.)<\/w:t>/g)].map(m => m[1]);
  ok(letterRuns.slice(0, 4).join(",") === "A.,B.,C.,D.", "exam Section A first question options not lettered in order: " + letterRuns.slice(0, 4).join(","));
  ok(letterRuns.length >= 100, "exam Section A option letters largely missing: " + letterRuns.length + " letter runs");
  ok(!/\\[a-zA-Z]{2,}/.test(exXml), "exam document.xml contains a raw TeX command");
  console.log("exam doc: oMath " + (exXml.match(/<m:oMath>/g) || []).length +
    ", tables " + (exXml.match(/<w:tbl>/g) || []).length +
    ", walker fallbacks " + JSON.stringify(stEx.fb));
}

if (fails.length) { console.log("DOCX CHECK FAILURES:\n" + fails.join("\n")); process.exit(1); }
console.log("DOCX CHECK: ALL PASS");
