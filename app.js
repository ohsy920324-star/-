"use strict";

/* ------------------------------------------------------------------ *
 *  강의록 → 강의자료 정리기 (누적형)
 *  - 스크립트를 하나씩 정리해 "조각"으로 누적
 *  - 섹션이 끝나면 표지+목차+범례+모든 조각을 합친 완성본 .docx 생성
 * ------------------------------------------------------------------ */

/* 박스 라벨 → 색상 (예시 파일과 동일한 색상) */
const BOX = {
  "용어":                    { cls: "green",  bg: "E2F0D9", bd: "70AD47", label: "538135" },
  "사장님께 이렇게 말하세요": { cls: "yellow", bg: "FFF2CC", bd: "BF9000", label: "7F6000" },
  "바로 쓰는 실전 멘트":     { cls: "pink",   bg: "FCE4EC", bd: "C5407E", label: "AD1457" },
  "핵심":                    { cls: "blue",   bg: "DEEBF7", bd: "2E75B6", label: "1F4E79" },
  "주의":                    { cls: "orange", bg: "FCE4D6", bd: "C55A11", label: "833C00" },
};
const BOX_LABELS = Object.keys(BOX);
const NAVY = "1F3864";

/* "이 자료 보는 법" 범례 (예시 파일 문구) */
const LEGEND = [
  ["용어", "마케팅 용어를 쉽게 풀어 설명합니다. 헷갈리는 단어는 여기서 확인하세요."],
  ["사장님께 이렇게 말하세요", "사장님(광고주/대표님)에게 그대로 말로 풀어서 설명할 때 쓰는 멘트 예시입니다."],
  ["바로 쓰는 실전 멘트", "카톡·문의·블로그에 바로 복붙해서 쓸 수 있는 실전 멘트·스크립트, 그리고 사례 글의 원문입니다."],
  ["핵심", "절대 잊으면 안 되는 핵심 포인트입니다."],
  ["주의", "실수하기 쉬운 부분, 꼭 조심해야 할 내용입니다."],
];

/* 구조화 출력(structured outputs)용 JSON 스키마 — 조각 1개 분량 */
const SCHEMA = {
  type: "object",
  additionalProperties: false,
  properties: {
    blocks: {
      type: "array",
      items: {
        type: "object",
        additionalProperties: false,
        properties: {
          type:  { type: "string", enum: ["h1", "h2", "para", "bullet", "box"] },
          text:  { type: "string" },
          label: { type: "string", enum: BOX_LABELS },
        },
        required: ["type", "text"],
      },
    },
  },
  required: ["blocks"],
};

const SYSTEM_PROMPT = `당신은 마케팅 강의 스크립트를 신입 직원도 이해할 수 있는 강의자료로 정리하는 전문 에디터입니다.
지금 받는 스크립트는 더 큰 '섹션'을 이루는 한 조각(한 회차/한 주제)입니다. 이 조각 하나를 정리합니다.

[절대 규칙]
1. 강의 내용을 하나도 빠짐없이, 강의가 진행된 순서(타임라인)대로 정리합니다. 요약·압축하지 말고 원문 디테일을 100% 살립니다.
2. 모든 마케팅/광고 전문용어는 [용어] 박스로 쉽고 풀어서 설명합니다.
3. 숫자·금액·비율·날짜·구체적 사례·멘트는 절대 축약하거나 생략하지 않습니다. 원문의 수치와 표현을 그대로 살립니다.
4. 이 스크립트 전체를 하나의 대단원으로 보고, 맨 처음 블록은 반드시 type "h1"(대단원 제목)으로 둡니다. 그 안을 필요에 따라 h2(소단원)로 나눕니다.

[5가지 박스]
- box + label "용어": 첫 줄에 용어 이름만 쓰고, 다음 줄부터 쉬운 설명을 적습니다.
- box + label "사장님께 이렇게 말하세요": 광고주(사장님/대표님)에게 그대로 말하는 멘트. 큰따옴표(" ")로 실제 말투를 살립니다.
- box + label "바로 쓰는 실전 멘트": 카톡·문의·블로그에 그대로 복붙할 수 있는 스크립트, 또는 사례 글 원문(원문이면 머리에 "(사례 글 원문 — …)" 식으로 표시).
- box + label "핵심": 절대 잊으면 안 되는 핵심 포인트.
- box + label "주의": 실수하기 쉬운 부분·주의사항.
그 외 일반 설명은 "para"(문단) 또는 "bullet"(목록 항목)을 사용합니다.

[출력 형식]
- 반드시 정해진 JSON 스키마(blocks 배열)로만 출력합니다.
- 표지·목차·범례·문서 제목은 만들지 않습니다(시스템이 자동으로 붙입니다).
- 첫 블록은 반드시 type "h1" 입니다.
- box 블록은 반드시 label 을 포함하고, box가 아닌 블록에는 label 을 넣지 않습니다.
- 박스 text 안에서 줄바꿈이 필요하면 \\n 으로 표현합니다.
- 한국어로 작성합니다.`;

/* ---------------------------------------------------------------- 상태 */
const $ = (id) => document.getElementById(id);
let pieces = []; // [{ name, blocks:[...] }]

function saveState() {
  localStorage.setItem("ck_pieces", JSON.stringify(pieces));
}
function loadState() {
  try { pieces = JSON.parse(localStorage.getItem("ck_pieces") || "[]") || []; }
  catch { pieces = []; }
}

window.addEventListener("DOMContentLoaded", () => {
  $("apiKey").value = localStorage.getItem("ck_apiKey") || "";
  $("model").value = localStorage.getItem("ck_model") || "claude-opus-4-8";
  $("docTitle").value = localStorage.getItem("ck_title") || "";
  $("docSubtitle").value = localStorage.getItem("ck_subtitle") || "";
  loadState();
  renderPieceList();
  renderPreview();
});

const persistMap = { apiKey: "ck_apiKey", model: "ck_model", docTitle: "ck_title", docSubtitle: "ck_subtitle" };
Object.keys(persistMap).forEach((id) => {
  $(id).addEventListener("input", () => {
    localStorage.setItem(persistMap[id], $(id).value);
    if (id === "docTitle" || id === "docSubtitle") renderPreview();
  });
});

/* ---------------------------------------------------------------- 버튼 */
$("addBtn").addEventListener("click", addPiece);
$("downloadBtn").addEventListener("click", () => buildDocx());
$("resetBtn").addEventListener("click", () => {
  if (!pieces.length || confirm("정리된 조각을 모두 지우고 새 섹션을 시작할까요?")) {
    pieces = []; saveState(); renderPieceList(); renderPreview();
    setStatus("초기화했습니다. 새 섹션을 시작하세요.");
  }
});

/* ---------------------------------------------------------------- 조각 추가 */
async function addPiece() {
  const apiKey = $("apiKey").value.trim();
  const model = $("model").value;
  const script = $("script").value.trim();

  if (!apiKey) return setStatus("API 키를 입력하세요.", true);
  if (!script) return setStatus("정리할 강의 스크립트를 붙여넣으세요.", true);

  $("addBtn").disabled = true;
  $("downloadBtn").disabled = true;
  setStatus("정리 중… (분량이 길면 1~3분 걸릴 수 있어요)");

  const userText =
    `다음은 한 섹션의 ${pieces.length + 1}번째 스크립트입니다. 규칙대로 정리해 주세요.\n\n` +
    `--- 강의 스크립트 시작 ---\n${script}\n--- 강의 스크립트 끝 ---`;

  try {
    const { text, stopReason } = await callClaude({
      apiKey, model, system: SYSTEM_PROMPT, userText,
      onProgress: (n) => setStatus(`정리 중… (${n.toLocaleString()}자 생성됨)`),
    });

    let parsed;
    try { parsed = JSON.parse(text); }
    catch {
      throw new Error(stopReason === "max_tokens"
        ? "출력이 최대 길이를 초과해 잘렸습니다. 스크립트를 둘로 나눠 다시 추가해 주세요."
        : "결과 JSON 파싱 실패. 다시 시도해 주세요.");
    }

    const blocks = (parsed.blocks || []).filter((b) => b && b.type && typeof b.text === "string");
    if (!blocks.length) throw new Error("정리 결과가 비어 있습니다. 다시 시도해 주세요.");

    const h1 = blocks.find((b) => b.type === "h1");
    const name = (h1 ? h1.text : `조각 ${pieces.length + 1}`).slice(0, 60);
    pieces.push({ name, blocks });
    saveState();
    renderPieceList();
    renderPreview();
    $("script").value = "";
    localStorage.removeItem("ck_script");

    setStatus(
      (stopReason === "max_tokens" ? "⚠️ 일부가 잘렸을 수 있어요(최대 길이 도달). " : "") +
      `“${name}” 추가 완료! 다음 스크립트를 붙여넣거나, 끝났으면 다운로드하세요.`
    );
  } catch (err) {
    console.error(err);
    setStatus(err.message || String(err), true);
  } finally {
    $("addBtn").disabled = false;
    $("downloadBtn").disabled = pieces.length === 0;
  }
}

function setStatus(msg, isError) {
  const el = $("status");
  el.textContent = msg;
  el.classList.toggle("error", !!isError);
}

/* ---------------------------------------------------------------- 조각 목록 UI */
function renderPieceList() {
  $("pieceCount").textContent = String(pieces.length);
  $("downloadBtn").disabled = pieces.length === 0;
  const ul = $("pieceList");
  ul.innerHTML = "";
  if (!pieces.length) {
    ul.innerHTML = '<li class="empty">아직 추가된 조각이 없습니다.</li>';
    return;
  }
  pieces.forEach((p, i) => {
    const li = document.createElement("li");
    li.className = "item";
    const num = document.createElement("span"); num.className = "num"; num.textContent = i + 1;
    const name = document.createElement("span"); name.className = "name"; name.title = p.name; name.textContent = p.name;
    const up = document.createElement("button"); up.textContent = "▲"; up.title = "위로"; up.disabled = i === 0;
    up.onclick = () => { [pieces[i - 1], pieces[i]] = [pieces[i], pieces[i - 1]]; saveState(); renderPieceList(); renderPreview(); };
    const del = document.createElement("button"); del.textContent = "삭제";
    del.onclick = () => { pieces.splice(i, 1); saveState(); renderPieceList(); renderPreview(); };
    li.append(num, name, up, del);
    ul.appendChild(li);
  });
}

/* ---------------------------------------------------------------- Claude API (브라우저 직접 호출 + 스트리밍) */
async function callClaude({ apiKey, model, system, userText, onProgress }) {
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "anthropic-dangerous-direct-browser-access": "true",
    },
    body: JSON.stringify({
      model, max_tokens: 32000, stream: true, system,
      output_config: { format: { type: "json_schema", schema: SCHEMA } },
      messages: [{ role: "user", content: userText }],
    }),
  });

  if (!res.ok || !res.body) {
    const t = await res.text().catch(() => "");
    throw new Error(`API 오류 ${res.status}: ${t.slice(0, 300)}`);
  }

  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "", text = "", stopReason = null;

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });
    let nl;
    while ((nl = buffer.indexOf("\n")) >= 0) {
      const line = buffer.slice(0, nl).trim();
      buffer = buffer.slice(nl + 1);
      if (!line.startsWith("data:")) continue;
      const data = line.slice(5).trim();
      if (!data) continue;
      let evt; try { evt = JSON.parse(data); } catch { continue; }
      if (evt.type === "content_block_delta" && evt.delta && evt.delta.type === "text_delta") {
        text += evt.delta.text;
        if (onProgress) onProgress(text.length);
      } else if (evt.type === "message_delta" && evt.delta && evt.delta.stop_reason) {
        stopReason = evt.delta.stop_reason;
      } else if (evt.type === "error") {
        const m = evt.error && (evt.error.message || JSON.stringify(evt.error));
        throw new Error("스트림 오류: " + (m || "알 수 없음"));
      }
    }
  }
  return { text, stopReason };
}

/* ---------------------------------------------------------------- 미리보기 (전체 누적) */
function allBlocks() {
  const out = [];
  pieces.forEach((p) => p.blocks.forEach((b) => out.push(b)));
  return out;
}

function renderPreview() {
  const root = $("preview");
  root.innerHTML = "";
  const title = $("docTitle").value.trim() || "강의 정리";
  const subtitle = $("docSubtitle").value.trim();

  if (!pieces.length) {
    root.innerHTML = '<p class="placeholder">정리 결과가 여기에 이어붙어 표시됩니다.</p>';
    return;
  }

  // 표지
  const cover = document.createElement("div");
  cover.className = "cover";
  cover.innerHTML = `<div class="title">${esc(title)}</div>` + (subtitle ? `<div class="subtitle">${esc(subtitle)}</div>` : "");
  root.appendChild(cover);

  // 목차
  add(root, "h2", "toc-title", "목차");
  const toc = document.createElement("div");
  toc.className = "toc";
  allBlocks().forEach((b) => {
    if (b.type === "h1" || b.type === "h2") {
      const d = document.createElement("div");
      if (b.type === "h2") d.className = "sub";
      d.textContent = (b.type === "h2" ? "· " : "") + b.text;
      toc.appendChild(d);
    }
  });
  root.appendChild(toc);

  // 범례
  add(root, "h2", "legend-title", "이 자료 보는 법");
  LEGEND.forEach(([label, desc]) => root.appendChild(previewBox(label, desc)));

  // 본문
  allBlocks().forEach((b) => {
    let el;
    if (b.type === "h1") { el = document.createElement("h3"); el.className = "h1"; el.textContent = b.text; }
    else if (b.type === "h2") { el = document.createElement("h4"); el.className = "h2"; el.textContent = b.text; }
    else if (b.type === "para") { el = document.createElement("p"); el.className = "para"; el.textContent = b.text; }
    else if (b.type === "bullet") { el = document.createElement("ul"); el.className = "bul"; const li = document.createElement("li"); li.textContent = b.text; el.appendChild(li); }
    else if (b.type === "box") { el = previewBox(b.label || "핵심", b.text); }
    if (el) root.appendChild(el);
  });
}

function previewBox(label, text) {
  const info = BOX[label] || BOX["핵심"];
  const el = document.createElement("div");
  el.className = "box " + info.cls;
  const lab = document.createElement("span"); lab.className = "label"; lab.textContent = "[" + label + "]";
  const body = document.createElement("div"); body.className = "body"; body.textContent = text;
  el.append(lab, body);
  return el;
}
function add(root, tag, cls, txt, append) {
  if (txt) { const e = document.createElement(tag); e.className = cls; e.textContent = txt; root.appendChild(e); }
}
function esc(s) { return s.replace(/[&<>]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c])); }

/* ---------------------------------------------------------------- 완성본 Word(.docx) 생성 */
function buildDocx() {
  const D = window.docx;
  if (!D) return setStatus("docx 라이브러리를 불러오지 못했습니다. 새로고침 후 다시 시도하세요.", true);
  if (!pieces.length) return setStatus("추가된 조각이 없습니다.", true);

  const title = $("docTitle").value.trim() || "강의 정리";
  const subtitle = $("docSubtitle").value.trim();
  const children = [];

  // 표지
  children.push(new D.Paragraph({
    alignment: D.AlignmentType.CENTER, spacing: { before: 1200, after: 200 },
    children: [new D.TextRun({ text: title, bold: true, size: 44, color: NAVY })],
  }));
  if (subtitle) {
    children.push(new D.Paragraph({
      alignment: D.AlignmentType.CENTER, spacing: { after: 200 },
      children: [new D.TextRun({ text: subtitle, size: 22, color: "595959" })],
    }));
  }
  children.push(new D.Paragraph({ children: [new D.PageBreak()] }));

  // 목차
  children.push(new D.Paragraph({ spacing: { after: 120 }, children: [new D.TextRun({ text: "목차", bold: true, size: 28, color: NAVY })] }));
  children.push(new D.TableOfContents("목차", { hyperlink: true, headingStyleRange: "1-2" }));
  children.push(new D.Paragraph({ children: [new D.PageBreak()] }));

  // 범례 ("이 자료 보는 법")
  children.push(new D.Paragraph({ text: "이 자료 보는 법", heading: D.HeadingLevel.HEADING_1, spacing: { after: 120 } }));
  LEGEND.forEach(([label, desc]) => {
    children.push(boxTable(D, { label, text: desc }));
    children.push(new D.Paragraph({ spacing: { after: 40 }, children: [] }));
  });
  children.push(new D.Paragraph({ children: [new D.PageBreak()] }));

  // 본문 (모든 조각)
  allBlocks().forEach((b) => {
    if (b.type === "h1") children.push(new D.Paragraph({ text: b.text, heading: D.HeadingLevel.HEADING_1, spacing: { before: 280, after: 120 } }));
    else if (b.type === "h2") children.push(new D.Paragraph({ text: b.text, heading: D.HeadingLevel.HEADING_2, spacing: { before: 200, after: 100 } }));
    else if (b.type === "para") children.push(new D.Paragraph({ spacing: { after: 120 }, children: [new D.TextRun({ text: b.text, size: 22 })] }));
    else if (b.type === "bullet") children.push(new D.Paragraph({ text: b.text, bullet: { level: 0 }, spacing: { after: 60 } }));
    else if (b.type === "box") {
      children.push(boxTable(D, b));
      children.push(new D.Paragraph({ spacing: { after: 60 }, children: [] }));
    }
  });

  const doc = new D.Document({
    styles: {
      default: { document: { run: { font: "맑은 고딕", size: 22 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { bold: true, size: 30, color: NAVY }, paragraph: { spacing: { before: 280, after: 120 } } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { bold: true, size: 25, color: "2A3B5E" }, paragraph: { spacing: { before: 200, after: 100 } } },
      ],
    },
    sections: [{ properties: {}, children }],
  });

  D.Packer.toBlob(doc).then((blob) => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = title.replace(/[\\/:*?"<>|]/g, "_") + ".docx";
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
    setStatus("완성본을 다운로드했습니다. (목차 페이지번호는 Word에서 F9로 업데이트)");
  });
}

/* 단일 셀 표 = 색상 박스 */
function boxTable(D, b) {
  const c = BOX[b.label] || BOX["핵심"];
  const lines = String(b.text).split("\n");
  const cell = [
    new D.Paragraph({ spacing: { after: 80 }, children: [new D.TextRun({ text: "[" + (b.label || "핵심") + "]", bold: true, size: 22, color: c.label })] }),
  ];
  lines.forEach((ln) => cell.push(new D.Paragraph({ spacing: { after: 20 }, children: [new D.TextRun({ text: ln, size: 22 })] })));

  const side = { style: D.BorderStyle.SINGLE, size: 6, color: c.bd };
  const left = { style: D.BorderStyle.SINGLE, size: 24, color: c.bd };
  return new D.Table({
    width: { size: 100, type: D.WidthType.PERCENTAGE },
    borders: { top: side, bottom: side, right: side, left, insideHorizontal: side, insideVertical: side },
    rows: [new D.TableRow({ children: [new D.TableCell({
      shading: { type: D.ShadingType.CLEAR, fill: c.bg, color: "auto" },
      margins: { top: 120, bottom: 120, left: 160, right: 160 },
      children: cell,
    })] })],
  });
}
