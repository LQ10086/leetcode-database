from __future__ import annotations

import base64
import html
import json
import mimetypes
import re
from dataclasses import dataclass, field
from pathlib import Path
from zipfile import ZipFile

from lxml import etree


ROOT = Path(__file__).resolve().parent
DOCX = ROOT / "leetcode速通.docx"
OUT = ROOT / "leetcode_knowledge_base.html"

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "v": "urn:schemas-microsoft-com:vml",
}
W = NS["w"]
R = NS["r"]


@dataclass
class Block:
    kind: str
    text: str = ""
    src: str = ""


@dataclass
class Question:
    title: str
    blocks: list[Block] = field(default_factory=list)


@dataclass
class Section:
    title: str
    intro: list[Block] = field(default_factory=list)
    questions: list[Question] = field(default_factory=list)


def w_attr(name: str) -> str:
    return f"{{{W}}}{name}"


def r_attr(name: str) -> str:
    return f"{{{R}}}{name}"


def local_name(node: etree._Element) -> str:
    return etree.QName(node).localname


def text_with_breaks(node: etree._Element) -> str:
    parts: list[str] = []
    for el in node.iter():
        name = local_name(el)
        if name == "t":
            parts.append(el.text or "")
        elif name == "tab":
            parts.append("\t")
        elif name in {"br", "cr"}:
            parts.append("\n")
    return "".join(parts).replace("\xa0", " ")


def paragraph_style(paragraph: etree._Element) -> str:
    ppr = paragraph.find("w:pPr", NS)
    if ppr is None:
        return ""
    p_style = ppr.find("w:pStyle", NS)
    if p_style is None:
        return ""
    return p_style.get(w_attr("val"), "")


def is_numbered(paragraph: etree._Element) -> bool:
    ppr = paragraph.find("w:pPr", NS)
    return ppr is not None and ppr.find("w:numPr", NS) is not None


def read_relationships(zf: ZipFile) -> dict[str, str]:
    rels_path = "word/_rels/document.xml.rels"
    if rels_path not in zf.namelist():
        return {}
    rels = etree.fromstring(zf.read(rels_path))
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    out: dict[str, str] = {}
    for rel in rels.findall(f"{{{rel_ns}}}Relationship"):
        rid = rel.get("Id")
        target = rel.get("Target")
        if not rid or not target:
            continue
        if not target.startswith("/"):
            target = "word/" + target
        out[rid] = target.lstrip("/")
    return out


def data_uri(zf: ZipFile, name: str) -> str:
    payload = zf.read(name)
    mime = mimetypes.guess_type(name)[0] or "application/octet-stream"
    return f"data:{mime};base64,{base64.b64encode(payload).decode('ascii')}"


def paragraph_images(paragraph: etree._Element, rels: dict[str, str], images: dict[str, str]) -> list[Block]:
    blocks: list[Block] = []
    for blip in paragraph.findall(".//a:blip", NS):
        rid = blip.get(r_attr("embed"))
        if not rid:
            continue
        target = rels.get(rid)
        if target and target in images:
            blocks.append(Block("image", src=images[target]))
    for image_data in paragraph.findall(".//v:imagedata", NS):
        rid = image_data.get(r_attr("id"))
        if not rid:
            continue
        target = rels.get(rid)
        if target and target in images:
            blocks.append(Block("image", src=images[target]))
    return blocks


def table_text(table: etree._Element) -> str:
    rows: list[str] = []
    for tr in table.findall("w:tr", NS):
        cells: list[str] = []
        for tc in tr.findall("w:tc", NS):
            paragraphs = [text_with_breaks(p) for p in tc.findall("w:p", NS)]
            cells.append("\n".join(paragraphs).rstrip())
        rows.append("\t".join(cells))
    return "\n".join(rows).rstrip()


def append_block(sections: list[Section], current_section: Section | None, current_question: Question | None, block: Block) -> None:
    if current_question is not None:
        current_question.blocks.append(block)
    elif current_section is not None:
        current_section.intro.append(block)
    else:
        sections.append(Section(block.text or "未分类", intro=[]))


def parse_docx(path: Path) -> tuple[str, list[Section]]:
    with ZipFile(path) as zf:
        doc = etree.fromstring(zf.read("word/document.xml"))
        rels = read_relationships(zf)
        images = {name: data_uri(zf, name) for name in zf.namelist() if name.startswith("word/media/")}

        title = path.stem
        sections: list[Section] = []
        current_section: Section | None = None
        current_question: Question | None = None

        body = doc.find("w:body", NS)
        if body is None:
            return title, sections

        for child in body:
            name = local_name(child)
            if name == "p":
                text = text_with_breaks(child).strip("\n")
                style = paragraph_style(child)
                imgs = paragraph_images(child, rels, images)

                if text:
                    if not sections and title == path.stem and style != "1":
                        title = text
                    elif style == "1":
                        current_section = Section(text)
                        sections.append(current_section)
                        current_question = None
                    elif style == "2":
                        if current_section is None:
                            current_section = Section("未分类")
                            sections.append(current_section)
                        current_question = Question(text)
                        current_section.questions.append(current_question)
                    elif style == "3":
                        append_block(sections, current_section, current_question, Block("subheading", text))
                    elif style == "4":
                        append_block(sections, current_section, current_question, Block("label", text))
                    elif is_numbered(child):
                        append_block(sections, current_section, current_question, Block("list", text))
                    else:
                        append_block(sections, current_section, current_question, Block("paragraph", text))

                for img in imgs:
                    append_block(sections, current_section, current_question, img)

            elif name == "tbl":
                text = table_text(child)
                if text:
                    append_block(sections, current_section, current_question, Block("code", text))

    return title, sections


def slug(value: str, fallback: str) -> str:
    value = re.sub(r"\s+", "-", value.strip().lower())
    value = re.sub(r"[^\w\u4e00-\u9fff-]+", "", value)
    return value[:80] or fallback


def block_html(block: Block) -> str:
    if block.kind == "subheading":
        return f"<h4>{html.escape(block.text)}</h4>"
    if block.kind == "label":
        return f"<p class=\"label\">{html.escape(block.text)}</p>"
    if block.kind == "list":
        return f"<li>{html.escape(block.text)}</li>"
    if block.kind == "code":
        return f"<pre><code>{html.escape(block.text)}</code></pre>"
    if block.kind == "image":
        return f"<figure><img alt=\"文档图片\" src=\"{block.src}\"></figure>"
    return f"<p>{html.escape(block.text)}</p>"


def render_blocks(blocks: list[Block]) -> str:
    out: list[str] = []
    in_list = False
    for block in blocks:
        if block.kind == "list":
            if not in_list:
                out.append("<ol>")
                in_list = True
            out.append(block_html(block))
            continue
        if in_list:
            out.append("</ol>")
            in_list = False
        out.append(block_html(block))
    if in_list:
        out.append("</ol>")
    return "\n".join(out)


def render_html(title: str, sections: list[Section]) -> str:
    total_questions = sum(len(section.questions) for section in sections)
    data = {
        "title": title,
        "sections": [
            {
                "title": section.title,
                "questions": [question.title for question in section.questions],
            }
            for section in sections
        ],
    }

    nav_parts: list[str] = []
    main_parts: list[str] = []
    used_ids: set[str] = set()

    def unique_id(base: str, fallback: str) -> str:
        candidate = slug(base, fallback)
        original = candidate
        i = 2
        while candidate in used_ids:
            candidate = f"{original}-{i}"
            i += 1
        used_ids.add(candidate)
        return candidate

    for si, section in enumerate(sections, 1):
        section_id = unique_id(section.title, f"section-{si}")
        nav_items = []
        question_html = []
        for qi, question in enumerate(section.questions, 1):
            qid = unique_id(f"{section.title}-{question.title}", f"q-{si}-{qi}")
            nav_items.append(f"<a href=\"#{qid}\">{html.escape(question.title)}</a>")
            question_html.append(
                f"""
                <details class="question" id="{qid}">
                    <summary>
                        <span>{html.escape(question.title)}</span>
                        <b>展开</b>
                    </summary>
                    <div class="answer">
                        {render_blocks(question.blocks)}
                    </div>
                </details>
                """
            )

        intro = render_blocks(section.intro)
        main_parts.append(
            f"""
            <section class="section" id="{section_id}" data-section="{html.escape(section.title)}">
                <div class="section-head">
                    <h2>{html.escape(section.title)}</h2>
                    <span>{len(section.questions)} 题</span>
                </div>
                {f'<div class="intro">{intro}</div>' if intro else ''}
                <div class="questions">
                    {''.join(question_html)}
                </div>
            </section>
            """
        )
        nav_parts.append(
            f"""
            <details class="nav-section" open>
                <summary><a href="#{section_id}">{html.escape(section.title)}</a><span>{len(section.questions)}</span></summary>
                <div>{''.join(nav_items)}</div>
            </details>
            """
        )

    json_data = json.dumps(data, ensure_ascii=False)

    return f"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{html.escape(title)} - 代码知识库</title>
<style>
:root {{
  color-scheme: light;
  --bg: #f7f8fb;
  --panel: #ffffff;
  --text: #20242d;
  --muted: #687182;
  --line: #dfe3eb;
  --accent: #0b6f63;
  --accent-soft: #e4f4f1;
  --code-bg: #151922;
  --code-text: #edf2f7;
  --shadow: 0 10px 30px rgba(23, 31, 46, .08);
}}
* {{ box-sizing: border-box; }}
html {{ scroll-behavior: smooth; }}
body {{
  margin: 0;
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
  background: var(--bg);
  color: var(--text);
}}
a {{ color: inherit; text-decoration: none; }}
.app {{ display: grid; grid-template-columns: 320px minmax(0, 1fr); min-height: 100vh; }}
.sidebar {{
  position: sticky;
  top: 0;
  height: 100vh;
  overflow: auto;
  border-right: 1px solid var(--line);
  background: #fbfcfe;
  padding: 22px 18px 28px;
}}
.brand h1 {{ margin: 0; font-size: 22px; line-height: 1.25; }}
.brand p {{ margin: 8px 0 18px; color: var(--muted); font-size: 14px; }}
.search {{
  width: 100%;
  border: 1px solid var(--line);
  border-radius: 8px;
  padding: 11px 12px;
  font-size: 15px;
  outline: none;
  background: #fff;
}}
.search:focus {{ border-color: var(--accent); box-shadow: 0 0 0 3px var(--accent-soft); }}
.nav {{ margin-top: 16px; display: grid; gap: 8px; }}
.nav-section {{
  border: 1px solid var(--line);
  border-radius: 8px;
  background: #fff;
  overflow: hidden;
}}
.nav-section summary {{
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 10px;
  padding: 10px 12px;
  cursor: pointer;
  font-weight: 700;
}}
.nav-section summary::-webkit-details-marker {{ display: none; }}
.nav-section summary span {{
  color: var(--accent);
  font-size: 12px;
  background: var(--accent-soft);
  padding: 2px 7px;
  border-radius: 999px;
}}
.nav-section div {{
  border-top: 1px solid var(--line);
  display: grid;
  padding: 6px;
}}
.nav-section div a {{
  color: var(--muted);
  font-size: 13px;
  line-height: 1.45;
  padding: 7px 8px;
  border-radius: 6px;
}}
.nav-section div a:hover {{ background: var(--accent-soft); color: var(--accent); }}
.content {{ min-width: 0; padding: 28px clamp(16px, 4vw, 44px) 60px; }}
.toolbar {{
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  margin-bottom: 20px;
}}
.toolbar h2 {{ margin: 0; font-size: 18px; }}
.actions {{ display: flex; gap: 8px; flex-wrap: wrap; }}
button {{
  border: 1px solid var(--line);
  background: #fff;
  color: var(--text);
  border-radius: 8px;
  min-height: 36px;
  padding: 0 12px;
  cursor: pointer;
  font-weight: 650;
}}
button:hover {{ border-color: var(--accent); color: var(--accent); }}
.section {{ scroll-margin-top: 20px; margin: 0 0 28px; }}
.section-head {{
  display: flex;
  align-items: baseline;
  justify-content: space-between;
  gap: 12px;
  border-bottom: 2px solid var(--text);
  padding-bottom: 10px;
  margin-bottom: 12px;
}}
.section-head h2 {{ margin: 0; font-size: clamp(22px, 3vw, 32px); }}
.section-head span {{ color: var(--muted); white-space: nowrap; }}
.intro {{ margin: 14px 0; }}
.questions {{ display: grid; gap: 10px; }}
.question {{
  border: 1px solid var(--line);
  border-radius: 8px;
  background: var(--panel);
  box-shadow: var(--shadow);
  overflow: hidden;
  scroll-margin-top: 16px;
}}
.question[open] {{ border-color: rgba(11, 111, 99, .42); }}
.question summary {{
  min-height: 54px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 14px 16px;
  cursor: pointer;
  font-weight: 750;
}}
.question summary::-webkit-details-marker {{ display: none; }}
.question summary b {{
  color: var(--accent);
  font-size: 13px;
  white-space: nowrap;
}}
.question[open] summary {{ border-bottom: 1px solid var(--line); background: #fbfefd; }}
.question[open] summary b::before {{ content: "收起"; }}
.question[open] summary b {{ font-size: 0; }}
.question[open] summary b::before {{ font-size: 13px; }}
.answer {{
  padding: 16px;
  font-size: 15px;
  line-height: 1.75;
}}
.answer h4 {{
  margin: 20px 0 8px;
  font-size: 18px;
  border-left: 4px solid var(--accent);
  padding-left: 9px;
}}
.answer h4:first-child {{ margin-top: 0; }}
.answer p, .intro p {{
  margin: 8px 0;
  white-space: pre-wrap;
}}
.label {{
  font-weight: 750;
  color: var(--accent);
}}
ol {{ margin: 8px 0 12px 24px; padding: 0; }}
li {{ margin: 4px 0; }}
pre {{
  margin: 10px 0 14px;
  padding: 14px;
  overflow: auto;
  border-radius: 8px;
  background: var(--code-bg);
  color: var(--code-text);
  line-height: 1.55;
  font-size: 13px;
  -webkit-overflow-scrolling: touch;
}}
code {{
  font-family: "Cascadia Code", Consolas, "SFMono-Regular", Menlo, monospace;
  white-space: pre;
}}
figure {{ margin: 12px 0; }}
img {{ max-width: 100%; height: auto; border: 1px solid var(--line); border-radius: 8px; }}
.hidden {{ display: none !important; }}
.empty {{
  display: none;
  padding: 28px;
  text-align: center;
  color: var(--muted);
  border: 1px dashed var(--line);
  border-radius: 8px;
  background: #fff;
}}
.empty.show {{ display: block; }}
@media (max-width: 860px) {{
  .app {{ display: block; }}
  .sidebar {{
    position: static;
    height: auto;
    max-height: none;
    border-right: 0;
    border-bottom: 1px solid var(--line);
    padding: 16px;
  }}
  .nav {{
    max-height: 42vh;
    overflow: auto;
    padding-right: 2px;
  }}
  .content {{ padding: 18px 14px 42px; }}
  .toolbar {{ align-items: flex-start; flex-direction: column; }}
  .section-head {{ align-items: flex-end; }}
  .question summary {{ padding: 13px 12px; }}
  .answer {{ padding: 13px 12px; font-size: 14px; }}
  pre {{ font-size: 12px; padding: 12px; }}
}}
</style>
</head>
<body>
<div class="app">
  <aside class="sidebar">
    <div class="brand">
      <h1>{html.escape(title)}</h1>
      <p>{len(sections)} 个板块 · {total_questions} 道题</p>
    </div>
    <input class="search" id="search" type="search" placeholder="搜索题目或解析" autocomplete="off">
    <nav class="nav" id="nav">
      {''.join(nav_parts)}
    </nav>
  </aside>
  <main class="content">
    <div class="toolbar">
      <h2>复习知识库</h2>
      <div class="actions">
        <button id="expandAll" type="button">全部展开</button>
        <button id="collapseAll" type="button">全部收起</button>
      </div>
    </div>
    <div class="empty" id="empty">没有匹配的内容</div>
    {''.join(main_parts)}
  </main>
</div>
<script id="kb-data" type="application/json">{html.escape(json_data)}</script>
<script>
const search = document.getElementById('search');
const empty = document.getElementById('empty');
const sections = [...document.querySelectorAll('.section')];
const questions = [...document.querySelectorAll('.question')];

function filterContent() {{
  const query = search.value.trim().toLowerCase();
  let visibleCount = 0;
  questions.forEach((question) => {{
    const text = question.innerText.toLowerCase();
    const hit = !query || text.includes(query);
    question.classList.toggle('hidden', !hit);
    if (hit) {{
      visibleCount += 1;
      if (query) question.open = true;
    }}
  }});
  sections.forEach((section) => {{
    const hasVisible = [...section.querySelectorAll('.question')].some(q => !q.classList.contains('hidden'));
    section.classList.toggle('hidden', !hasVisible);
  }});
  empty.classList.toggle('show', visibleCount === 0);
}}

search.addEventListener('input', filterContent);
document.getElementById('expandAll').addEventListener('click', () => questions.forEach(q => q.open = true));
document.getElementById('collapseAll').addEventListener('click', () => questions.forEach(q => q.open = false));

if (location.hash) {{
  const target = document.querySelector(location.hash);
  if (target && target.tagName.toLowerCase() === 'details') target.open = true;
}}
</script>
</body>
</html>
"""


def main() -> None:
    title, sections = parse_docx(DOCX)
    OUT.write_text(render_html(title, sections), encoding="utf-8")
    questions = sum(len(section.questions) for section in sections)
    blocks = sum(len(question.blocks) for section in sections for question in section.questions)
    print(f"Wrote {OUT}")
    print(f"Sections: {len(sections)}")
    print(f"Questions: {questions}")
    print(f"Question blocks: {blocks}")


if __name__ == "__main__":
    main()
