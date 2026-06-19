"use strict";

/* ------------------------------------------------------------------ *
 *  강의록 → 강의자료 정리기
 *  - 브라우저에서 사용자의 Claude API 키로 직접 호출
 *  - 결과를 5색 박스 형식으로 미리보기 + Word(.docx) 다운로드
 * ------------------------------------------------------------------ */

// 박스 라벨 → 색상 매핑 (미리보기 클래스 / docx 색상)
const BOX = {
  "용어":                    { cls: "green",  bg: "E6F4EA", bd: "34A853", label: "1E7E34" },
  "사장님께 이렇게 말하세요": { cls: "yellow", bg: "FFF4D6", bd: "F1C232", label: "9C6F00" },
  "바로 쓰는 실전 멘트":     { cls: "pink",   bg: "FDE7EF", bd: "E91E63", label: "AD1457" },
  "핵심":                    { cls: "blue",   bg: "E3F0FB", bd: "3D85C6", label: "1155CC" },
  "주의":                    { cls: "orange", bg: "FDEBD7", bd: "E69138", label: "B45309" },
};
const BOX_LABELS = Object.keys(BOX);

// Claude 구조화 출력(structured outputs)용 JSON 스키마
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
          type:  { type: "string", enum: ["title", "h1", "h2", "para", "bullet", "box"] },
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

[절대 규칙]
1. 강의 내용을 하나도 빠짐없이, 강의가 진행된 순서(타임라인)대로 정리합니다. 요약·압축하지 말고 모든 내용을 살립니다.
2. 모든 마케팅/광고 전문용어는 [용어] 박스로 쉽고 풀어서 설명합니다.
3. 숫자, 금액, 비율, 날짜, 구체적 사례, 멘트는 절대 축약하거나 생략하지 않습니다. 원문의 수치와 표현을 그대로 살립니다.
4. 내용을 적절히 대단원(h1)과 소단원(h2)으로 구조화합니다.

[블록 종류 — 5가지 박스]
- box + label "용어": 마케팅 용어를 쉽게 풀어 설명
- box + label "사장님께 이렇게 말하세요": 광고주(사장님)에게 설명할 때 그대로 쓸 멘트
- box + label "바로 쓰는 실전 멘트": 카톡 등에 그대로 복붙할 수 있는 스크립트
- box + label "핵심": 꼭 기억할 핵심 포인트
- box + label "주의": 자주 하는 실수·주의사항
그 외 일반 설명은 "para"(문단) 또는 "bullet"(목록 항목)을 사용합니다.

[출력 형식]
- 반드시 정해진 JSON 스키마(blocks 배열)로만 출력합니다.
- 첫 블록은 type "title" 로 문서 제목을 넣습니다(사용자가 제목을 지정한 경우 그 제목을 사용).
- box 블록은 반드시 label 을 포함합니다. box가 아닌 블록은 label 을 넣지 않습니다.
- 박스의 text 안에서 줄바꿈이 필요하면 \\n 으로 표현합니다.
- 한국어로 작성합니다.`;

// ---------------------------------------------------------------- 상태
const $ = (id) => document.getElementById(id);
let lastBlocks = null;
let lastTitle = "강의 정리";

// 저장된 값 복원
window.addEventListener("DOMContentLoaded", () => {
  $("apiKey").value = localStorage.getItem("ck_apiKey") || "";
  $("model").value = localStorage.getItem("ck_model") || "claude-opus-4-8";
  $("script").value = localStorage.getItem("ck_script") || "";
  $("docTitle").value = localStorage.getItem("ck_title") || "";
});
["apiKey", "model", "script", "docTitle"].forEach((id) => {
  $(id).addEventListener("input", () => {
    localStorage.setItem("ck_" + (id === "script" ? "script" : id === "docTitle" ? "title" : id), $(id).value);
  });
});

// ---------------------------------------------------------------- 실행
$("runBtn").addEventListener("click", run);
$("downloadBtn").addEventListener("click", () => {
  if (lastBlocks) buildDocx(lastBlocks, lastTitle);
});

async function run() {
  const apiKey = $("apiKey").value.trim();
  const model = $("model").value;
  const script = $("script").value.trim();
  const userTitle = $("docTitle").value.trim();

  if (!apiKey) return setStatus("API 키를 입력하세요.", true);
  if (!script) return setStatus("정리할 강의 스크립트를 붙여넣으세요.", true);

  $("runBtn").disabled = true;
  $("downloadBtn").disabled = true;
  setStatus("정리 중… (분량이 길면 1~3분 걸릴 수 있어요)");

  const userText =
    (userTitle ? `문서 제목: "${userTitle}"\n\n` : "") +
    `아래 강의 스크립트를 규칙대로 정리해 주세요.\n\n--- 강의 스크립트 시작 ---\n${script}\n--- 강의 스크립트 끝 ---`;

  try {
    const { text, stopReason } = await callClaude({ apiKey, model, system: SYSTEM_PROMPT, userText,
      onProgress: (n) => setStatus(`정리 중… (${n.toLocaleString()}자 생성됨)`) });

    let parsed;
    try {
      parsed = JSON.parse(text);
    } catch (e) {
      throw new Error(
        stopReason === "max_tokens"
          ? "출력이 최대 길이를 초과해 잘렸습니다. 스크립트를 둘로 나눠 다시 시도해 주세요."
          : "결과 JSON 파싱에 실패했습니다. 다시 시도해 주세요."
      );
    }

    const blocks = (parsed.blocks || []).filter((b) => b && b.type && typeof b.text === "string");
    if (!blocks.length) throw new Error("정리 결과가 비어 있습니다. 다시 시도해 주세요.");

    const titleBlock = blocks.find((b) => b.type === "title");
    lastTitle = userTitle || (titleBlock ? titleBlock.text : "강의 정리");
    lastBlocks = blocks;

    renderPreview(blocks, lastTitle);
    $("downloadBtn").disabled = false;
    setStatus(
      stopReason === "max_tokens"
        ? "⚠️ 일부가 잘렸을 수 있어요(최대 길이 도달). 결과를 확인하고, 필요하면 스크립트를 나눠 다시 시도하세요."
        : "완료! ‘Word(.docx) 다운로드’를 누르세요."
    );
  } catch (err) {
    console.error(err);
    setStatus(err.message || String(err), true);
  } finally {
    $("runBtn").disabled = false;
  }
}

function setStatus(msg, isError) {
  const el = $("status");
  el.textContent = msg;
  el.classList.toggle("error", !!isError);
}

// ---------------------------------------------------------------- Claude API (브라우저 직접 호출 + 스트리밍)
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
      model,
      max_tokens: 32000,
      stream: true,
      system,
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

      let evt;
      try { evt = JSON.parse(data); } catch { continue; }

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

// ---------------------------------------------------------------- 미리보기 렌더링
function renderPreview(blocks, title) {
  const root = $("preview");
  root.innerHTML = "";

  const h1 = document.createElement("h1");
  h1.className = "doc-title";
  h1.textContent = title;
  root.appendChild(h1);

  // 목차
  const tocTitle = document.createElement("h2");
  tocTitle.className = "toc-title";
  tocTitle.textContent = "목차";
  root.appendChild(tocTitle);
  const toc = document.createElement("div");
  toc.className = "toc";
  blocks.forEach((b) => {
    if (b.type === "h1" || b.type === "h2") {
      const d = document.createElement("div");
      d.textContent = (b.type === "h2" ? "    · " : "") + b.text;
      toc.appendChild(d);
    }
  });
  if (!toc.children.length) toc.innerHTML = "<div>(소제목 없음)</div>";
  root.appendChild(toc);

  // 본문
  blocks.forEach((b) => {
    if (b.type === "title") return;
    let el;
    if (b.type === "h1") { el = document.createElement("h3"); el.className = "h1"; el.textContent = b.text; }
    else if (b.type === "h2") { el = document.createElement("h4"); el.className = "h2"; el.textContent = b.text; }
    else if (b.type === "para") { el = document.createElement("p"); el.className = "para"; el.textContent = b.text; }
    else if (b.type === "bullet") {
      el = document.createElement("ul"); el.className = "bul";
      const li = document.createElement("li"); li.textContent = b.text; el.appendChild(li);
    } else if (b.type === "box") {
      const info = BOX[b.label] || BOX["핵심"];
      el = document.createElement("div");
      el.className = "box " + info.cls;
      const lab = document.createElement("span"); lab.className = "label";
      lab.textContent = "[" + (b.label || "핵심") + "]";
      const body = document.createElement("div"); body.className = "body"; body.textContent = b.text;
      el.appendChild(lab); el.appendChild(body);
    }
    if (el) root.appendChild(el);
  });
}

// ---------------------------------------------------------------- Word(.docx) 생성
function buildDocx(blocks, title) {
  const D = window.docx;
  if (!D) { setStatus("docx 라이브러리를 불러오지 못했습니다. 새로고침 후 다시 시도하세요.", true); return; }

  const children = [];

  // 제목
  children.push(new D.Paragraph({
    alignment: D.AlignmentType.CENTER,
    spacing: { after: 240 },
    children: [new D.TextRun({ text: title, bold: true, size: 40 })],
  }));

  // 목차 (Word 필드 → 열어서 업데이트하면 페이지 번호 자동 채움)
  children.push(new D.Paragraph({
    spacing: { before: 120, after: 120 },
    children: [new D.TextRun({ text: "목차", bold: true, size: 28 })],
  }));
  children.push(new D.TableOfContents("목차", { hyperlink: true, headingStyleRange: "1-2" }));
  children.push(new D.Paragraph({ children: [new D.PageBreak()] }));

  // 본문
  blocks.forEach((b) => {
    if (b.type === "title") return;
    if (b.type === "h1") {
      children.push(new D.Paragraph({ text: b.text, heading: D.HeadingLevel.HEADING_1, spacing: { before: 280, after: 120 } }));
    } else if (b.type === "h2") {
      children.push(new D.Paragraph({ text: b.text, heading: D.HeadingLevel.HEADING_2, spacing: { before: 200, after: 100 } }));
    } else if (b.type === "para") {
      children.push(new D.Paragraph({ spacing: { after: 120 }, children: [new D.TextRun({ text: b.text, size: 22 })] }));
    } else if (b.type === "bullet") {
      children.push(new D.Paragraph({ text: b.text, bullet: { level: 0 }, spacing: { after: 60 } }));
    } else if (b.type === "box") {
      children.push(boxTable(D, b));
      children.push(new D.Paragraph({ spacing: { after: 60 }, children: [] }));
    }
  });

  const doc = new D.Document({
    styles: { default: { document: { run: { font: "맑은 고딕", size: 22 } } } },
    sections: [{ properties: {}, children }],
  });

  D.Packer.toBlob(doc).then((blob) => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = (title || "강의정리").replace(/[\\/:*?"<>|]/g, "_") + ".docx";
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  });
}

// 단일 셀 표 = 색상 박스
function boxTable(D, b) {
  const c = BOX[b.label] || BOX["핵심"];
  const lines = String(b.text).split("\n");

  const cell = [
    new D.Paragraph({
      spacing: { after: 80 },
      children: [new D.TextRun({ text: "[" + (b.label || "핵심") + "]", bold: true, size: 22, color: c.label })],
    }),
  ];
  lines.forEach((ln) => {
    cell.push(new D.Paragraph({ spacing: { after: 20 }, children: [new D.TextRun({ text: ln, size: 22 })] }));
  });

  const side = { style: D.BorderStyle.SINGLE, size: 6, color: c.bd };
  const left = { style: D.BorderStyle.SINGLE, size: 24, color: c.bd };

  return new D.Table({
    width: { size: 100, type: D.WidthType.PERCENTAGE },
    borders: { top: side, bottom: side, right: side, left: left, insideHorizontal: side, insideVertical: side },
    rows: [
      new D.TableRow({
        children: [
          new D.TableCell({
            shading: { type: D.ShadingType.CLEAR, fill: c.bg, color: "auto" },
            margins: { top: 120, bottom: 120, left: 160, right: 160 },
            children: cell,
          }),
        ],
      }),
    ],
  });
}
