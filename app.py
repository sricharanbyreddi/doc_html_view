import os
import glob
import json
import re
import base64
import mimetypes
from pathlib import Path

import mammoth
from docx import Document
import markdown as markdown_lib
from bs4 import BeautifulSoup, NavigableString
from flask import Flask, render_template_string, abort

BASE_DIR = Path(__file__).resolve().parent
DOCS_DIR = Path(os.environ.get("DOCS_DIR", BASE_DIR / "documents"))
OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", BASE_DIR / "dist"))
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)

HTML_TEMPLATE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Document Hub</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.10.0/styles/github.min.css">
<style>
:root{--line:#e7eaf0;--muted:#667085;--ink:#182230;--blue:#2563eb;--soft:#eff6ff}
*{box-sizing:border-box}html,body{height:100%;margin:0}
body{font-family:Inter,-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;color:var(--ink);background:#f7f8fa}
.app{height:100vh;display:flex;overflow:hidden}
.sidebar{width:290px;min-width:230px;background:#fff;border-right:1px solid var(--line);display:flex;flex-direction:column}
.sidebar-head{padding:22px 18px 16px;border-bottom:1px solid var(--line)}
.brand{font-size:18px;font-weight:750;letter-spacing:-.3px;margin:0 0 5px}
.subtitle{font-size:12px;color:var(--muted);margin:0 0 15px}
.search{width:100%;border:1px solid #d0d5dd;border-radius:9px;padding:10px 11px;font:inherit;font-size:13px;outline:none}
.search:focus{border-color:#84adff;box-shadow:0 0 0 3px #dbeafe}
.doc-list{list-style:none;margin:0;padding:10px;overflow:auto;flex:1}
.doc-item{padding:11px 12px;margin:3px 0;border-radius:8px;cursor:pointer;font-size:13px;line-height:1.4;color:#344054;overflow-wrap:anywhere}
.doc-item:hover{background:#f4f6f9}.doc-item.active{background:var(--soft);color:var(--blue);font-weight:650}
.sidebar-foot{border-top:1px solid var(--line);padding:12px 18px;font-size:11px;color:var(--muted)}
.main{flex:1;min-width:0;display:flex;flex-direction:column;background:#fff}
.topbar{min-height:72px;padding:18px 32px;border-bottom:1px solid var(--line);display:flex;align-items:center;justify-content:space-between;gap:16px;background:#fff}
.top-title{font-size:17px;font-weight:700;overflow-wrap:anywhere}
.doc-count{font-size:12px;color:var(--muted);white-space:nowrap}
.viewer{overflow:auto;flex:1;padding:36px clamp(20px,5vw,64px);line-height:1.7}
.article{max-width:980px;margin:0 auto;overflow-wrap:anywhere}
.article h1,.article h2,.article h3,.article h4{line-height:1.3;color:#101828;margin:1.7em 0 .65em}
.article h1:first-child,.article h2:first-child,.article h3:first-child{margin-top:0}
.article p{margin:.65em 0;color:#344054}
.article ul,.article ol{padding-left:1.6rem}
.article li{margin:.25em 0}
.article img{max-width:100%;height:auto;display:inline-block;vertical-align:middle;border-radius:5px}
.article figure{margin:1.25rem 0;text-align:center}
.article table{border-collapse:collapse;width:100%;margin:1.3rem 0;font-size:14px;display:block;overflow-x:auto}
.article td,.article th{border:1px solid #d0d5dd;padding:8px 11px;text-align:left;vertical-align:top;min-width:70px}
.article th{background:#f2f4f7;font-weight:650}
.article blockquote{margin:1rem 0;padding:.25rem 1rem;border-left:3px solid #98a2b3;color:#475467;background:#f9fafb}
.article{font-size:15px;color:#344054}.article h1{font-size:30px;border-bottom:1px solid #eaecf0;padding-bottom:10px}.article h2{font-size:24px}.article h3{font-size:19px}.article strong{color:#101828;font-weight:700}.article table strong{font-weight:600}.article hr{border:0;border-top:1px solid #eaecf0;margin:24px 0}.article pre{margin:1rem 0;padding:16px;border:1px solid #e4e7ec;border-radius:10px;background:#f8fafc;overflow:auto;white-space:pre;line-height:1.55;font-size:13px}
.article pre code{font-family:ui-monospace,SFMono-Regular,Consolas,"Liberation Mono",monospace}
.diagram-note{font-size:13px;color:#667085;margin:8px 0}.article code:not(pre code){font-family:ui-monospace,Consolas,monospace;background:#f2f4f7;padding:2px 5px;border-radius:4px;font-size:.9em}
.mermaid{margin:1.2rem 0;padding:18px;overflow:auto;text-align:center;background:#fff;border:1px solid var(--line);border-radius:10px}
.mermaid svg{max-width:100%;height:auto}
.empty{height:100%;display:flex;align-items:center;justify-content:center;text-align:center;color:#98a2b3;padding:30px}
.empty strong{display:block;color:#475467;font-size:17px;margin-bottom:8px}
.notice{max-width:720px;margin:20px auto;padding:16px;border:1px solid #fedf89;background:#fffaeb;border-radius:10px;color:#93370d;font-size:13px}
@media(max-width:700px){.sidebar{width:42%;min-width:145px}.sidebar-head{padding:16px 10px}.brand{font-size:15px}.topbar{padding:15px}.viewer{padding:24px 16px}.doc-item{padding:9px;font-size:12px}}

.article pre, .article pre code { display:block; white-space:pre; overflow-x:auto; max-width:100%; font-family:Consolas,'Cascadia Code',monospace; font-size:13px; line-height:1.55; }
.article pre { background:#f6f8fa; border:1px solid #e4e7ec; border-radius:8px; padding:16px; color:#24292f; margin:16px 0; }
.article .folder-tree code { color:#344054; }
.article table { display:block; max-width:100%; overflow-x:auto; border-collapse:collapse; }
.article th { background:#f2f4f7; }
</style>
</head>
<body>
<div class="app">
  <aside class="sidebar">
    <div class="sidebar-head">
      <h1 class="brand">📚 Document Hub</h1>
      <p class="subtitle">Your documents, in one place</p>
      <input id="search" class="search" type="search" placeholder="Search documents…" aria-label="Search documents">
    </div>
    <ul id="docList" class="doc-list"></ul>
    <div class="sidebar-foot"><span id="visibleCount"></span> documents<br><span style="overflow-wrap:anywhere">Folder: {{ documents_folder }}</span></div>
  </aside>
  <main class="main">
    <header class="topbar">
      <div id="activeTitle" class="top-title">Select a document</div>
      <div id="docCount" class="doc-count"></div>
    </header>
    <section id="viewer" class="viewer">
      <div class="empty"><div><strong>Choose a document</strong>Select a file from the sidebar to read it here.</div></div>
    </section>
  </main>
</div>
<script src="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.10.0/highlight.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.min.js"></script>
<script>
const docs = {{ documents_json|safe }};
const documentsFolder = {{ documents_folder|tojson }};
const list = document.getElementById('docList');
const search = document.getElementById('search');
const viewer = document.getElementById('viewer');
const activeTitle = document.getElementById('activeTitle');
const visibleCount = document.getElementById('visibleCount');
const docCount = document.getElementById('docCount');
let selectedId = null;

if (window.mermaid) mermaid.initialize({startOnLoad:false, securityLevel:'strict', theme:'default'});

function renderList() {
  const query = search.value.trim().toLowerCase();
  const filtered = docs.filter(d => d.title.toLowerCase().includes(query));
  list.replaceChildren();
  visibleCount.textContent = filtered.length;
  filtered.forEach(doc => {
    const li = document.createElement('li');
    li.className = 'doc-item' + (doc.id === selectedId ? ' active' : '');
    li.textContent = doc.title;
    li.title = doc.title;
    li.tabIndex = 0;
    li.setAttribute('role','button');
    li.onclick = () => loadDoc(doc);
    li.onkeydown = e => { if(e.key === 'Enter' || e.key === ' ') { e.preventDefault(); loadDoc(doc); } };
    list.appendChild(li);
  });
  if (!filtered.length) {
    const li = document.createElement('li');
    li.className = 'doc-item';
    li.style.whiteSpace = 'normal';
    li.style.color = '#667085';
    li.style.cursor = 'default';
    li.textContent = docs.length ? 'No matching documents' : 'No .docx files found. Put your Word documents in: ' + documentsFolder;
    list.appendChild(li);
    if (!docs.length) {
      viewer.innerHTML = '<div class="empty"><div><strong>No documents found</strong>Place .docx files in the documents folder shown in the sidebar, then refresh this page.</div></div>';
      activeTitle.textContent = 'Document Hub';
      docCount.textContent = '0 documents';
    }
  }
}
async function loadDoc(doc) {
  selectedId = doc.id;
  activeTitle.textContent = doc.title;
  docCount.textContent = (docs.indexOf(doc)+1) + ' of ' + docs.length;
  viewer.innerHTML = '<article class="article">' + doc.content + '</article>';
  renderList();
  viewer.scrollTop = 0;
  if (window.hljs) viewer.querySelectorAll('pre code').forEach(block => hljs.highlightElement(block));
  if (window.mermaid) {
    const diagrams = viewer.querySelectorAll('.mermaid');
    if (diagrams.length) {
      for (const diagram of diagrams) {
        const source = diagram.textContent || '';
        try {
          await mermaid.parse(source);
          await mermaid.run({nodes: [diagram]});
        } catch (err) {
          console.warn('Mermaid rendering failed:', err);
          const pre = document.createElement('pre');
          const code = document.createElement('code');
          code.textContent = source;
          pre.appendChild(code);
          const note = document.createElement('p');
          note.className = 'diagram-note';
          note.textContent = 'Diagram source shown because Mermaid could not parse this diagram.';
          diagram.replaceWith(note, pre);
        }
      }
    }
  }
}
search.addEventListener('input', renderList);
renderList();
</script>
</body>
</html>
"""

def data_uri_image(image):
    """Return Mammoth-compatible HTML attributes for an embedded image."""
    with image.open() as image_file:
        raw = image_file.read()
    content_type = image.content_type or "image/png"
    encoded = base64.b64encode(raw).decode("ascii")
    # mammoth.images.img_element expects a callback that returns a dict of
    # HTML attributes; img_element itself wraps those attributes in an <img>.
    return {
        "src": f"data:{content_type};base64,{encoded}",
        "alt": image.alt_text or "",
    }

CODE_STYLE_MAP = [
    "p[style-name='Code'] => pre:code",
    "p[style-name='Code Block'] => pre:code",
    "p[style-name='Source Code'] => pre:code",
    "p[style-name='SQL'] => pre:code",
    "p[style-name='Python'] => pre:code",
    "r[style-name='Code'] => code",
]

MERMAID_START = re.compile(
    r"^(?:flowchart\s|graph\s|sequenceDiagram\b|classDiagram\b|stateDiagram(?:-v2)?\b|"
    r"erDiagram\b|journey\b|gantt\b|pie\b|mindmap\b|timeline\b|gitGraph\b|"
    r"quadrantChart\b|xychart\b|block-beta\b|architecture-beta\b)",
    re.IGNORECASE,
)

def normalize_html(html_text):
    """
    Convert Word-generated HTML while handling Markdown-like content conservatively.
    Escaped newlines are decoded only in text blocks that clearly contain code,
    Mermaid, or Markdown table syntax. Ordinary paragraphs, images and tables remain.
    """
    soup = BeautifulSoup(html_text, "html.parser")

    def new_mermaid(source):
        fragment = BeautifulSoup("", "html.parser")
        node = fragment.new_tag("div", attrs={"class": "mermaid"})
        node.string = source.strip()
        return node

    # Work with block elements in document order; do not flatten the whole document.
    for node in list(soup.find_all(["p", "pre"])):
        if node.parent is None:
            continue
        raw = node.get_text("", strip=False).strip()
        if not raw:
            continue

        # Word documents often contain escaped \n sequences copied from Markdown/code files.
        decoded = raw
        if r"\n" in decoded and (
            "```" in decoded or MERMAID_START.match(decoded) or
            re.search(r"(^|\\n)\s*\|.*\|", decoded) or
            re.search(r"(^|\\n)\s*(?:def |class |import |SELECT\b|INSERT\b|CREATE\b|UPDATE\b|DELETE\b)", decoded, re.I)
        ):
            decoded = decoded.replace(r"\r\n", "\n").replace(r"\n", "\n").replace(r"\t", "\t")

        # Mermaid fenced blocks are extracted before Markdown conversion.
        mm = re.search(r"```mermaid\s*(.*?)```", decoded, flags=re.S | re.I)
        if mm:
            node.replace_with(new_mermaid(mm.group(1)))
            continue

        # Mermaid source in a code paragraph should be rendered as a diagram.
        stripped = re.sub(r"^```(?:\w+)?\s*|\s*```$", "", decoded, flags=re.S).strip()
        if MERMAID_START.match(stripped) and (
            "-->" in stripped or "subgraph " in stripped or "participant " in stripped or
            "sequenceDiagram" in stripped or "classDiagram" in stripped
        ):
            node.replace_with(new_mermaid(stripped))
            continue

        # Convert Markdown table exports that use || as cell delimiters to normal pipes.
        # Only do this when multiple table-like lines are present.
        if "|" in decoded and re.search(r"\|\s*:?-{3,}", decoded):
            decoded = decoded.replace("||", "|")
            decoded = re.sub(r"(?m)^\s*\|\s*", "| ", decoded)
            decoded = re.sub(r"(?m)\s*\|\s*$", " |", decoded)

        markdown_markers = (
            decoded.startswith(("# ", "## ", "### ", "#### ", "> ", "- ", "* ", "1. ", "|", "```"))
            or "**" in decoded or "```" in decoded
            or (decoded.count("\n") >= 1 and "|" in decoded)
        )
        if markdown_markers:
            rendered = markdown_lib.markdown(
                decoded,
                extensions=["tables", "fenced_code", "sane_lists"]
            )
            fragment = BeautifulSoup(rendered, "html.parser")
            node.replace_with(fragment)

    # Convert Mermaid fenced code that Markdown generated into a Mermaid container.
    for code in list(soup.select("pre code")):
        language_match = re.search(r"language-([\w+-]+)", " ".join(code.get("class", [])))
        code_text = code.get_text("", strip=False).strip()
        if language_match and language_match.group(1).lower() == "mermaid":
            code.parent.replace_with(new_mermaid(code_text))
        elif MERMAID_START.match(code_text) and ("-->" in code_text or "subgraph " in code_text):
            code.parent.replace_with(new_mermaid(code_text))

    # Responsive images; keep each image at its original location in the document.
    for img in soup.find_all("img"):
        img["loading"] = "lazy"
        img["style"] = (img.get("style", "") + ";max-width:100%;height:auto").strip(";")

    # Detect repository trees, Python, and SQL even when Word marks every line as Normal.
    def clean_text(node):
        return node.get_text("", strip=False).replace("\u00a0", " ").rstrip()

    def is_tree_line(text):
        t = text.rstrip()
        return bool(
            re.match(r"^[│|├└┌┐┬┴┼─┤┘╰╭╴╶╼╾]", t)
            or re.match(r"^(?:├|└|\\|)[─-]{1,}", t)
            or re.match(r"^[\w.@-][\w./@ -]*/$", t)
            or re.match(r"^(?:│|\\|)\s*(?:├|└|[│|])", t)
        )

    def code_kind(text):
        t = text.strip()
        if not t:
            return None
        # SQL statements, DDL column declarations, constraints, and continuation lines.
        if re.match(
            r"^(?:--|SELECT\b|WITH\b|INSERT\b|UPDATE\b|DELETE\b|CREATE\b|ALTER\b|DROP\b|TRUNCATE\b|"
            r"GRANT\b|REVOKE\b|PRIMARY\s+KEY\b|FOREIGN\s+KEY\b|CONSTRAINT\b|REFERENCES\b|"
            r"UNIQUE\b|CHECK\b|JOIN\b|LEFT\s+JOIN\b|RIGHT\s+JOIN\b|INNER\s+JOIN\b|"
            r"WHERE\b|GROUP\s+BY\b|ORDER\s+BY\b|HAVING\b|VALUES\b|SET\b|FROM\b|ON\b|"
            r"[`\"']?\w+[`\"']?\s+(?:UUID|SERIAL|BIGSERIAL|INTEGER|INT|BIGINT|TEXT|VARCHAR|CHAR|"
            r"BOOLEAN|BOOL|DATE|TIMESTAMP\w*|NUMERIC|DECIMAL|JSONB|JSON|TIME)\b|"
            r"\);?$|;\s*$)",
            t, re.I
        ):
            return "sql"
        # Python statements and typical continuation/assignment lines.
        if re.match(
            r"^(?:#|from\s+\S+\s+import\b|import\b|async\s+def\b|def\b|class\b|@\w+|"
            r"if\b|elif\b|else\s*:|for\b|while\b|try\s*:|except\b|finally\s*:|with\b|"
            r"return\b|raise\b|pass\s*$|break\s*$|continue\s*$|yield\b|assert\b|print\s*\(|"
            r"st\.\w+\s*=|[A-Za-z_]\w*(?:\.[A-Za-z_]\w*)*\s*=\s*.+|"
            r"[\)\]\}\],?$|[A-Za-z_]\w+\s*:\s*[\w\[\], .]+$)",
            t
        ):
            return "python"
        return None

    # Work from the top-level document flow, grouping runs before replacing nodes.
    # This is important: classifying a single paragraph cannot identify a multi-line snippet.
    top_blocks = list(soup.find_all(["p", "pre", "table", "ul", "ol", "h1", "h2", "h3", "blockquote"]))
    handled = set()

    # Folder trees: include root folder labels and tree glyphs, even when glyphs were
    # converted to vertical bars by the source document.
    for p in list(soup.find_all("p")):
        if p.parent is None or p in handled:
            continue
        text = clean_text(p).strip()
        if not text:
            continue
        if not is_tree_line(text) and not re.match(r"^[\w.-][\w./ -]*/$", text):
            continue

        run = [p]
        cursor = p.find_next_sibling()
        while cursor is not None and cursor.name in ("p", "pre"):
            ct = clean_text(cursor).strip()
            if not ct:
                # Keep a single blank line if the next paragraph continues the tree.
                nxt = cursor.find_next_sibling()
                if nxt is not None and nxt.name == "p" and is_tree_line(clean_text(nxt).strip()):
                    run.extend([cursor, nxt])
                    handled.update([cursor, nxt])
                    cursor = nxt.find_next_sibling()
                    continue
                break
            if is_tree_line(ct) or re.match(r"^[\w.-][\w./ -]*/$", ct):
                run.append(cursor)
                handled.add(cursor)
                cursor = cursor.find_next_sibling()
                continue
            break

        actual = [x for x in run if clean_text(x).strip()]
        # A root label plus two branches, or three branch lines, is enough to establish a tree.
        if len(actual) >= 3:
            lines = [clean_text(x) for x in actual]
            pre = soup.new_tag("pre", attrs={"class": "folder-tree"})
            code = soup.new_tag("code", attrs={"class": "language-text"})
            code.string = "\n".join(lines)
            pre.append(code)
            actual[0].replace_with(pre)
            for old_node in run:
                if old_node is not actual[0] and old_node.parent is not None:
                    old_node.decompose()
            handled.update(run)

    # Python and SQL: start on a strong signal, then collect continuation lines
    # (including SQL column definitions) until a blank or non-code prose paragraph.
    for p in list(soup.find_all("p")):
        if p.parent is None or p in handled:
            continue
        text = clean_text(p).strip()
        kind = code_kind(text)
        if not kind:
            continue

        run = [p]
        current = p.find_next_sibling()
        run_kind = kind
        blank_count = 0
        while current is not None and current.name in ("p", "pre"):
            ct = clean_text(current).strip()
            if not ct:
                nxt = current.find_next_sibling()
                if nxt is not None and nxt.name == "p" and code_kind(clean_text(nxt).strip()):
                    run.append(current)
                    run.append(nxt)
                    handled.update([current, nxt])
                    current = nxt.find_next_sibling()
                    continue
                break

            ck = code_kind(ct)
            if ck:
                run.append(current)
                handled.add(current)
                if ck == "sql":
                    run_kind = "sql"
                current = current.find_next_sibling()
                continue

            # Continuations: indented expressions, SQL types/constraints, punctuation,
            # function-call arguments, and SQL comma-terminated columns.
            sql_continuation = run_kind == "sql" and bool(re.match(
                r"^(?:[`\w\"']+\s+(?:UUID|SERIAL|BIGSERIAL|INTEGER|INT|BIGINT|TEXT|VARCHAR|CHAR|"
                r"BOOLEAN|BOOL|DATE|TIMESTAMP\w*|NUMERIC|DECIMAL|JSONB|JSON|TIME)\b|"
                r"PRIMARY\s+KEY|FOREIGN\s+KEY|CONSTRAINT|REFERENCES|UNIQUE|CHECK|"
                r"ON\s+DELETE|ON\s+UPDATE|DEFAULT\b|NOT\s+NULL\b|\)|\);|,|"
                r"['\"].*['\"]\s*[,)]?;?$)",
                ct, re.I
            ))
            python_continuation = run_kind == "python" and bool(re.match(
                r"^(?:[)\]}.,]|[A-Za-z_]\w*\s*=\s*.+|[A-Za-z_]\w+\s*:\s*.+|"
                r"['\"].*['\"]\s*[,)]?$|st\.\w+\(.+|requests\.\w+\(.+)",
                ct
            ))
            if sql_continuation or python_continuation:
                run.append(current)
                handled.add(current)
                current = current.find_next_sibling()
                continue
            break

        nonempty = [x for x in run if clean_text(x).strip()]
        # Convert even a single fenced-language-like statement; group when multiple lines.
        if len(nonempty) >= 2 or kind in ("sql", "python"):
            pre = soup.new_tag("pre", attrs={"class": f"detected-code {run_kind}-block"})
            code = soup.new_tag("code", attrs={"class": [f"language-{run_kind}"]})
            code.string = "\n".join(clean_text(x) for x in nonempty)
            pre.append(code)
            nonempty[0].replace_with(pre)
            for old_node in run:
                if old_node is not nonempty[0] and old_node.parent is not None:
                    old_node.decompose()
            handled.update(run)

    # Ensure code blocks have a code element for consistent styling.
    for pre in soup.find_all("pre"):
        if not pre.find("code"):
            code = soup.new_tag("code")
            code.string = pre.get_text("", strip=False)
            pre.clear()
            pre.append(code)

    return str(soup)

def repair_markdown_tables(source):
    """Join soft-wrapped lines inside Markdown pipe tables without changing normal prose."""
    lines = source.splitlines()
    out = []
    in_table = False
    for line in lines:
        stripped = line.strip()
        is_pipe_line = stripped.startswith("|") and "|" in stripped[1:]
        is_separator = bool(re.match(r"^\|?\s*:?-{3,}", stripped)) and "|" in stripped
        if is_pipe_line:
            in_table = True
            out.append(stripped)
        elif in_table and stripped and out and not stripped.startswith(("```", "#", ">")):
            # A continuation line within a wrapped table cell belongs to the previous row.
            out[-1] = out[-1].rstrip()
            if out[-1].endswith("|"):
                out[-1] = out[-1][:-1].rstrip() + " " + stripped + " |"
            else:
                out[-1] += " " + stripped
        else:
            if not stripped:
                in_table = False
            elif stripped.startswith(("```", "#", ">")):
                in_table = False
            out.append(line)
    return "\n".join(out)


def markdown_to_document_html(source):
    """Render complete Markdown documents and promote Mermaid fences to diagram nodes."""
    source = source.replace("\r\n", "\n").replace("\r", "\n")
    source = repair_markdown_tables(source)
    rendered = markdown_lib.markdown(
        source,
        extensions=["tables", "fenced_code", "sane_lists", "nl2br"]
    )
    soup = BeautifulSoup(rendered, "html.parser")

    for code in list(soup.select("pre > code")):
        classes = " ".join(code.get("class", []))
        language_match = re.search(r"language-([\w+-]+)", classes)
        language = language_match.group(1).lower() if language_match else ""
        code_text = code.get_text("", strip=False).strip()
        if language == "mermaid":
            div = soup.new_tag("div", attrs={"class": "mermaid"})
            div.string = code_text
            code.parent.replace_with(div)
        elif not language and MERMAID_START.match(code_text) and (
            "-->" in code_text or "subgraph " in code_text or "participant " in code_text
        ):
            div = soup.new_tag("div", attrs={"class": "mermaid"})
            div.string = code_text
            code.parent.replace_with(div)

    return str(soup)


def convert_docx(path):
    """
    Source-aware DOCX conversion:
    - Markdown-export DOCX files are rendered from the original paragraph text as a whole,
      preserving line breaks, tables, code fences and Mermaid blocks.
    - Ordinary Word documents continue through Mammoth for images, tables and formatting.
    """
    # These files were authored as Markdown text placed inside Word paragraphs. Converting
    # their raw paragraph text avoids Mammoth treating the whole Markdown source as prose.
    try:
        word_doc = Document(str(path))
        paragraphs = [p.text for p in word_doc.paragraphs if p.text.strip()]
        full_text = "\n\n".join(paragraphs)
        inline_images = len(word_doc.inline_shapes) > 0
        markdown_signals = (
            full_text.count("\n") >= 4 and
            (
                bool(re.search(r"(?m)^\s{0,3}#{1,6}\s", full_text)) or
                "```" in full_text or
                bool(re.search(r"(?m)^\s*\|.*\|", full_text))
            )
        )
        if markdown_signals and not inline_images:
            return markdown_to_document_html(full_text)
    except Exception as exc:
        print(f"  Source-aware detection skipped for {path.name}: {exc}")

    with open(path, "rb") as docx_file:
        result = mammoth.convert_to_html(
            docx_file,
            style_map="\n".join(CODE_STYLE_MAP),
            convert_image=mammoth.images.img_element(data_uri_image),
        )
    content = normalize_html(result.value)
    if result.messages:
        print(f"  Conversion notes for {path.name}:")
        for message in result.messages[:5]:
            print(f"   - {message}")
    return content

def build_documents():
    DOCS_DIR.mkdir(parents=True, exist_ok=True)
    files = sorted(
        (p for p in DOCS_DIR.iterdir() if p.is_file() and p.suffix.lower() == ".docx"),
        key=lambda p: p.name.lower()
    ) if DOCS_DIR.exists() else []
    print(f"Scanning documents folder: {DOCS_DIR}")
    print(f"Found {len(files)} .docx file(s).")
    documents = []
    for path in files:
        safe_id = re.sub(r"[^a-zA-Z0-9_-]+", "-", path.stem).strip("-").lower()
        try:
            documents.append({
                "id": safe_id or f"document-{len(documents)+1}",
                "title": path.stem,
                "content": convert_docx(path),
            })
            print(f"Converted: {path.name}")
        except Exception as exc:
            print(f"ERROR converting {path.name}: {exc}")
    return documents

@app.route("/")
def index():
    documents = build_documents()
    return render_template_string(
        HTML_TEMPLATE,
        documents_json=json.dumps(documents, ensure_ascii=False).replace("</", "<\\/"),
        documents_folder=str(DOCS_DIR)
    )

@app.route("/health")
def health():
    return {"status": "ok", "documents_directory": str(DOCS_DIR)}

if __name__ == "__main__":
    DOCS_DIR.mkdir(parents=True, exist_ok=True)
    print(f"Documents folder: {DOCS_DIR}")
    print("Open http://127.0.0.1:5000 in your browser.")
    app.run(host="127.0.0.1", port=int(os.environ.get("PORT", "5000")), debug=True)
