#!/usr/bin/env python3
"""Word mirror of Silva Culta entries: Laura edits .docx files, this script carries
the changes into _pages/*.md and publishes them.

Usage (from anywhere):
  python3 _tools/mirror.py status            what changed where (docx edits, md edits, conflicts)
  python3 _tools/mirror.py export [slug...]  write .docx mirrors from the .md entries
                                             (all, or the given slugs); never overwrites a
                                             .docx with unimported edits unless --force
  python3 _tools/mirror.py import            carry edited/new .docx files into _pages/*.md
  python3 _tools/mirror.py publish           import, run sync.py + --check, commit, push
  python3 _tools/mirror.py roundtrip         self-test: md -> docx -> md for every entry

The .docx folder is ../writing/00 silva culta/ (relative to the treasury repo).
Needs python-docx. State (hashes at last sync) lives in that folder's .mirror-state.json.
"""
import copy
import datetime
import glob
import hashlib
import io
import json
import os
import re
import subprocess
import sys
import tempfile

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DOCS = os.path.normpath(os.path.join(ROOT, "..", "writing", "00 silva culta"))
STATE = os.path.join(DOCS, ".mirror-state.json")
TEMPLATE_NAME = "00 New entry (duplicate me).docx"
LIVE = "https://lauramoorewrites-del.github.io/treasury"
MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]
FLEURON_MD = '<p class="fleuron" aria-hidden="true">❧</p>'
DIVIDER = "———"
META_LABELS = ["Live page", "Topics", "Related", "Figuring out", "Created", "Updated", "Entry id"]


# ---------------------------------------------------------------- utilities

def sha(path):
    with open(path, "rb") as f:
        return hashlib.sha256(f.read()).hexdigest()


def load_state():
    if os.path.exists(STATE):
        return json.load(open(STATE, encoding="utf-8"))
    return {}


def save_state(state):
    with open(STATE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=1, ensure_ascii=False, sort_keys=True)


def today_str():
    d = datetime.date.today()
    return f"{d.day} {MONTHS[d.month - 1]} {d.year}"


def slugify(t):
    t = t.lower().replace("’", "").replace("'", "")
    return re.sub(r"[^a-z0-9]+", "-", t).strip("-")


def filename_for(title):
    name = title.replace("/", "–").replace(":", " –").strip()
    return name[:150] + ".docx"


def locked(path):
    d, b = os.path.split(path)
    return any(os.path.exists(os.path.join(d, p + b[k:])) for p in ("~$",) for k in (0, 1, 2))


def unquote(v):
    v = v.strip()
    if len(v) >= 2 and v[0] == v[-1] == '"':
        return v[1:-1].replace('\\"', '"').replace("\\\\", "\\")
    if len(v) >= 2 and v[0] == v[-1] == "'":
        return v[1:-1].replace("''", "'")
    return v


def yq(s):
    return '"' + s.replace("\\", "\\\\").replace('"', '\\"') + '"'


# ---------------------------------------------------------------- md side

def read_entry(path):
    text = open(path, encoding="utf-8").read()
    m = re.match(r"---\n(.*?)\n---\n", text, re.S)
    fm, body = m.group(1), text[m.end():]
    e = {"fm": fm, "body": body}
    for key in ("id", "title", "created", "updated"):
        mm = re.search(rf"^{key}:\s*(.*)$", fm, re.M)
        e[key] = unquote(mm.group(1)) if mm else ""
    mm = re.search(r"^topics:\s*\[(.*)\]\s*$", fm, re.M)
    e["topics"] = [unquote(t) for t in mm.group(1).split(",") if t.strip()] if mm else []
    e["related"] = [(unquote(t), u.strip("/"))
                    for t, u in re.findall(r'^  - title:\s*(.*)\n    url:\s*"?([^"\n]*)"?', fm, re.M)]
    mm = re.search(r"^figuring_out:\n((?:  - .*\n?)*)", fm + "\n", re.M)
    e["figuring"] = [unquote(x) for x in re.findall(r"^  - (.*)$", mm.group(1), re.M)] if mm else []
    return e


def all_entries():
    out = {}
    for f in sorted(glob.glob(os.path.join(ROOT, "_pages", "*.md"))):
        out[os.path.basename(f)[:-3]] = read_entry(f)
    return out


def taxonomy():
    fm = open(os.path.join(ROOT, "index.md"), encoding="utf-8").read().split("\n---", 1)[0]
    block = re.search(r"^topic_groups:\n((?:(?:[ \t].*)?\n)*)", fm + "\n", re.M).group(1)
    return [unquote(x) for x in re.findall(r"^\s{6,}- (.*)$", block, re.M)]


def parse_blocks(body):
    """Split an entry body into typed blocks."""
    blocks = []
    chunks = re.split(r"\n[ \t]*\n", body.strip("\n"))
    prev_list_end = False
    for ch in chunks:
        lines = ch.split("\n")
        if ch.strip() == FLEURON_MD:
            blocks.append(("fleuron", None))
        elif ch.strip() == "---":
            blocks.append(("divider", None))
        elif re.match(r"^#{1,3} ", ch):
            level = len(ch.split(" ", 1)[0])
            blocks.append((f"h{level}", ch.split(" ", 1)[1]))
        elif lines[0].startswith(">"):
            paras, cur = [], []
            for l in lines:
                s = re.sub(r"^> ?", "", l)
                if s.strip() == "" or re.match(r"^(\d+\. |- )", s):
                    if cur:
                        paras.append(" ".join(cur))
                    cur = [s] if s.strip() else []
                else:
                    cur.append(s)
            if cur:
                paras.append(" ".join(cur))
            paras = [re.sub(r"\\\\ ", "\n", p) for p in paras]
            for p in paras:
                if re.match(r"^\d+\. ", p):
                    blocks.append(("qnum", p))
                elif re.match(r"^- ", p):
                    blocks.append(("qbullet", p[2:]))
                else:
                    blocks.append(("quote", p))
            blocks.append(("endquote", None))
        elif re.match(r"^- ", lines[0]):
            loose = prev_list_end
            items = []
            for l in lines:
                if l.startswith("- "):
                    items.append(l[2:])
                else:
                    items[-1] += " " + l.strip()
            for it in items:
                blocks.append(("bullet_loose" if loose else "bullet", it))
            # mark the previous list's items as loose too when separated by a blank line
            if loose:
                for i in range(len(blocks) - len(items) - 1, -1, -1):
                    if blocks[i][0] == "bullet":
                        blocks[i] = ("bullet_loose", blocks[i][1])
                    elif blocks[i][0] != "bullet_loose":
                        break
            prev_list_end = True
            continue
        elif re.match(r"^\d+\. ", lines[0]):
            for l in lines:
                blocks.append(("num", l))
        elif re.match(r"^ {4}!\[", ch) or re.match(r"^!\[", ch):
            indent = ch.startswith("    ")
            m = re.match(r'^\s*!\[(.*?)\]\((.*)\)\s*$', ch)
            blocks.append(("image_cont" if indent else "image", (m.group(1), m.group(2))))
            # an indented image keeps the list going
            if indent:
                continue
        else:
            blocks.append(("para", re.sub(r"\\\\ ", "\n", " ".join(l.strip() for l in lines))))
        prev_list_end = False
    return blocks


INLINE_LINK = re.compile(r"\[((?:[^\[\]]|\[[^\[\]]*\])*)\]\(((?:[^()\s]|\([^()\s]*\))+)\)")


def parse_inline(s, b=False, i=False, url=None):
    """Markdown inline -> list of (text, bold, italic, url)."""
    out, buf, k = [], [], 0

    def flush():
        if buf:
            out.append(("".join(buf), b, i, url))
            buf.clear()

    def find_close(tok, start):
        j = start
        while j < len(s):
            if s[j] == "\\":
                j += 2
                continue
            if s.startswith(tok, j):
                if tok == "*" and s.startswith("**", j):
                    j += 2
                    continue
                return j
            j += 1
        return -1

    while k < len(s):
        c = s[k]
        if c == "\\" and k + 1 < len(s) and s[k + 1] in "\\*_[]()#>-`!":
            buf.append(s[k + 1])
            k += 2
            continue
        if c == "[" and url is None:
            m = INLINE_LINK.match(s, k)
            if m:
                flush()
                out.extend(parse_inline(m.group(1), b, i, m.group(2)))
                k = m.end()
                continue
        if s.startswith("<i>", k) or s.startswith("<em>", k):
            tag = "i" if s.startswith("<i>", k) else "em"
            j = s.find(f"</{tag}>", k)
            if j > 0:
                flush()
                out.extend(parse_inline(s[k + len(tag) + 2:j], b, True, url))
                k = j + len(tag) + 3
                continue
        for tok, nb, ni in (("***", True, True), ("**", True, False), ("*", False, True)):
            if s.startswith(tok, k):
                j = find_close(tok, k + len(tok))
                if j > k + len(tok):
                    flush()
                    out.extend(parse_inline(s[k + len(tok):j], b or nb, i or ni, url))
                    k = j + len(tok)
                    break
        else:
            if c == "_" and (k == 0 or not s[k - 1].isalnum()):
                m = re.match(r"_([^_\s][^_]*?)_(?![\w])", s[k:])
                if m:
                    flush()
                    out.extend(parse_inline(m.group(1), b, True, url))
                    k += m.end()
                    continue
            buf.append(c)
            k += 1
            continue
        continue
    flush()
    return out


# ---------------------------------------------------------------- docx writing

def build_styles(d):
    st = d.styles
    normal = st["Normal"]
    normal.font.name = "Georgia"
    normal.font.size = Pt(11.5)
    normal.paragraph_format.space_after = Pt(8)
    normal.paragraph_format.line_spacing = 1.2
    for name in ("Title", "Heading 1", "Heading 2", "Heading 3"):
        s = st[name]
        s.font.name = "Georgia"
        s.font.color.rgb = RGBColor(0x2E, 0x3B, 0x34)
    st["Title"].font.size = Pt(24)
    q = st["Quote"]
    q.font.italic = False
    q.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    q.paragraph_format.left_indent = Inches(0.45)
    q.paragraph_format.right_indent = Inches(0.3)
    q.paragraph_format.space_after = Pt(6)

    def add(name, base, **kw):
        if name in [s.name for s in st]:
            return st[name]
        s = st.add_style(name, 1)
        s.base_style = st[base]
        s.quick_style = True
        for k, v in kw.items():
            if k == "align":
                s.paragraph_format.alignment = v
            elif k == "size":
                s.font.size = v
            elif k == "color":
                s.font.color.rgb = v
            elif k == "left":
                s.paragraph_format.left_indent = v
            elif k == "first":
                s.paragraph_format.first_line_indent = v
            elif k == "after":
                s.paragraph_format.space_after = v
            elif k == "before":
                s.paragraph_format.space_before = v
        return s

    green = RGBColor(0x5B, 0x6E, 0x5F)
    add("Attribution", "Normal", size=Pt(10.5), color=RGBColor(0x55, 0x55, 0x55), after=Pt(14))
    add("Fleuron", "Normal", align=WD_ALIGN_PARAGRAPH.CENTER, color=green, before=Pt(6), after=Pt(10))
    add("Divider", "Normal", align=WD_ALIGN_PARAGRAPH.CENTER, color=green)
    add("Quote Numbered", "Quote", left=Inches(0.75), first=Inches(-0.25), after=Pt(3))
    add("Quote Bullet", "Quote", left=Inches(0.75), first=Inches(-0.25), after=Pt(3))
    add("Numbered Item", "Normal", left=Inches(0.3), first=Inches(-0.25), after=Pt(3))
    add("List Bullet Spaced", "List Bullet", after=Pt(10))
    add("Image Under Item", "List Continue")
    add("Meta", "Normal", size=Pt(9.5), after=Pt(0))


def add_hyperlink(p, url):
    rid = p.part.relate_to(url, RT.HYPERLINK, is_external=True)
    h = OxmlElement("w:hyperlink")
    h.set(qn("r:id"), rid)
    p._p.append(h)
    return h


def add_run(container_el, p, text, b, i, link):
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    if b:
        rpr.append(OxmlElement("w:b"))
    if i:
        rpr.append(OxmlElement("w:i"))
    if link:
        c = OxmlElement("w:color")
        c.set(qn("w:val"), "3B6B8C")
        u = OxmlElement("w:u")
        u.set(qn("w:val"), "single")
        rpr.append(c)
        rpr.append(u)
    if len(rpr):
        r.append(rpr)
    for n, piece in enumerate(text.split("\n")):
        if n:
            r.append(OxmlElement("w:br"))
        t = OxmlElement("w:t")
        t.text = piece
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        r.append(t)
    container_el.append(r)


def to_word_url(u):
    if u.startswith("/treasury/"):
        return LIVE + u[len("/treasury"):]
    return u


def write_inline(p, md):
    segs = parse_inline(md)
    k = 0
    while k < len(segs):
        url = segs[k][3]
        if url:
            h = add_hyperlink(p, to_word_url(url))
            while k < len(segs) and segs[k][3] == url:
                add_run(h, p, segs[k][0], segs[k][1], segs[k][2], True)
                k += 1
        else:
            add_run(p._p, p, segs[k][0], segs[k][1], segs[k][2], False)
            k += 1


def image_path(src):
    m = re.search(r'"(/assets/[^"]+)"', src)
    rel = m.group(1) if m else src.replace("/treasury", "", 1)
    return rel, os.path.join(ROOT, rel.lstrip("/"))


def add_image(d, style, alt, src):
    rel, path = image_path(src)
    p = d.add_paragraph(style=style)
    run = p.add_run()
    pic = run.add_picture(path, width=Inches(4.5))
    if pic.height > Inches(4.5):  # keep portrait photos to about half a page
        ratio = Inches(4.5) / pic.height
        pic.height, pic.width = Inches(4.5), int(pic.width * ratio)
    docpr = p._p.findall(".//" + qn("wp:docPr"))[-1]
    docpr.set("descr", alt)
    docpr.set("name", rel if src.startswith("{{") else "raw:" + src)


def meta_cell_paras(cell, items, linkify=None):
    cell.paragraphs[0].style = "Meta"
    first = True
    for it in items:
        p = cell.paragraphs[0] if first else cell.add_paragraph(style="Meta")
        first = False
        if linkify:
            h = add_hyperlink(p, linkify[1](it))
            add_run(h, p, linkify[0](it), False, False, True)
        else:
            add_run(p._p, p, it, False, False, False)


def export_one(slug, e, path):
    d = docx.Document()
    build_styles(d)
    sec = d.sections[0]
    sec.left_margin = sec.right_margin = Inches(1.1)
    d.add_paragraph(e["title"], style="Title")
    t = d.add_table(rows=len(META_LABELS), cols=2)
    t.style = "Table Grid"
    vals = {
        "Live page": [slug] if slug else [],
        "Topics": [", ".join(e["topics"])],
        "Related": e["related"],
        "Figuring out": e["figuring"],
        "Created": [e["created"]],
        "Updated": [e["updated"]],
        "Entry id": [e["id"]],
    }
    for row, label in zip(t.rows, META_LABELS):
        row.cells[0].width = Inches(1.2)
        row.cells[1].width = Inches(5.1)
        row.cells[0].paragraphs[0].style = "Meta"
        add_run(row.cells[0].paragraphs[0]._p, None, label, True, False, False)
        if label == "Live page":
            meta_cell_paras(row.cells[1], vals[label], (lambda s: f"{LIVE}/{s}/", lambda s: f"{LIVE}/{s}/"))
        elif label == "Related":
            meta_cell_paras(row.cells[1], vals[label], (lambda r: r[0], lambda r: f"{LIVE}/{r[1]}/"))
        else:
            meta_cell_paras(row.cells[1], vals[label])
    d.add_paragraph("")
    for kind, val in parse_blocks(e["body"]):
        if kind == "fleuron":
            d.add_paragraph("❧", style="Fleuron")
        elif kind == "divider":
            d.add_paragraph(DIVIDER, style="Divider")
        elif kind in ("h1", "h2", "h3"):
            write_inline(d.add_paragraph(style=f"Heading {kind[1]}"), val)
        elif kind == "quote":
            write_inline(d.add_paragraph(style="Quote"), val)
        elif kind == "qnum":
            write_inline(d.add_paragraph(style="Quote Numbered"), val)
        elif kind == "qbullet":
            write_inline(d.add_paragraph(style="Quote Bullet"), val)
        elif kind == "endquote":
            continue
        elif kind == "bullet":
            write_inline(d.add_paragraph(style="List Bullet"), val)
        elif kind == "bullet_loose":
            write_inline(d.add_paragraph(style="List Bullet Spaced"), val)
        elif kind == "num":
            write_inline(d.add_paragraph(style="Numbered Item"), val)
        elif kind in ("image", "image_cont"):
            add_image(d, "Image Under Item" if kind == "image_cont" else "Normal", *val)
        else:
            style = "Attribution" if val.startswith("— ") else "Normal"
            write_inline(d.add_paragraph(style=style), val)
    d.core_properties.title = e["title"]
    d.core_properties.subject = slug or ""
    d.save(path)


# ---------------------------------------------------------------- docx reading

def run_flags(r):
    rpr = r.find(qn("w:rPr"))
    b = i = False
    if rpr is not None:
        def on(tag):
            el = rpr.find(qn(tag))
            return el is not None and el.get(qn("w:val")) not in ("0", "false", "none")
        b, i = on("w:b"), on("w:i")
        rs = rpr.find(qn("w:rStyle"))
        if rs is not None:
            v = rs.get(qn("w:val")).lower()
            i = i or v in ("emphasis", "subtleemphasis", "intenseemphasis")
            b = b or v in ("strong",)
    return b, i


def read_paragraph(p_el, part):
    """Return (segments, images); segments are (text, b, i, url); images are (name, alt, blob, ext)."""
    segs, images = [], []
    field = {"url": None, "in_instr": False}

    def walk(el, url):
        for ch in el:
            tag = ch.tag
            if tag == qn("w:r"):
                fc = ch.find(qn("w:fldChar"))
                if fc is not None:
                    typ = fc.get(qn("w:fldCharType"))
                    if typ == "begin":
                        field["in_instr"], field["instr"] = True, ""
                    elif typ == "separate":
                        field["in_instr"] = False
                        m = re.search(r'HYPERLINK\s+"([^"]+)"', field.get("instr", ""))
                        field["url"] = m.group(1) if m else None
                    elif typ == "end":
                        field["url"], field["in_instr"] = None, False
                    continue
                it = ch.find(qn("w:instrText"))
                if it is not None:
                    field["instr"] = field.get("instr", "") + (it.text or "")
                    continue
                if field["in_instr"]:
                    continue
                b, i = run_flags(ch)
                for sub in ch:
                    if sub.tag == qn("w:t"):
                        segs.append((sub.text or "", b, i, url or field["url"]))
                    elif sub.tag == qn("w:tab"):
                        segs.append((" ", b, i, url or field["url"]))
                    elif sub.tag in (qn("w:br"), qn("w:cr")):
                        if sub.get(qn("w:type")) in (None, "textWrapping"):
                            segs.append(("\n", b, i, url or field["url"]))
                    elif sub.tag == qn("w:noBreakHyphen"):
                        segs.append(("‑", b, i, url))
                    elif sub.tag == qn("w:drawing"):
                        dp = sub.find(".//" + qn("wp:docPr"))
                        blip = sub.find(".//" + qn("a:blip"))
                        rid = blip.get(qn("r:embed")) if blip is not None else None
                        ip = part.related_parts[rid] if rid else None
                        images.append((dp.get("name", ""), dp.get("descr", "") or dp.get("title", ""),
                                       ip.blob if ip else None,
                                       os.path.splitext(ip.partname)[1] if ip else ".png"))
            elif tag == qn("w:hyperlink"):
                rid = ch.get(qn("r:id"))
                u = part.rels[rid].target_ref if rid and rid in part.rels else None
                if u is None and ch.get(qn("w:anchor")):
                    u = "#" + ch.get(qn("w:anchor"))
                walk(ch, u)
            elif tag in (qn("w:ins"), qn("w:smartTag"), qn("w:customXml"), qn("w:fldSimple"),
                         qn("w:sdt"), qn("w:sdtContent"), qn("w:moveTo")):
                u = url
                if tag == qn("w:fldSimple"):
                    m = re.search(r'HYPERLINK\s+"([^"]+)"', ch.get(qn("w:instr")) or "")
                    u = m.group(1) if m else url
                walk(ch, u)
            # w:del, w:moveFrom, w:pPr, bookmarks, comments: skipped
    walk(p_el, None)
    return segs, images


def to_md_url(u):
    if u and u.startswith(LIVE + "/"):
        return "/treasury" + u[len(LIVE):]
    return u


def emit_inline(segs):
    # merge, normalise links, push edge spaces outside formatting
    merged = []
    for t, b, i, u in segs:
        if not t:
            continue
        u = to_md_url(u)
        if t.strip(" ") == "":
            b = i = False
        if merged and merged[-1][1:] == (b, i, u):
            merged[-1] = (merged[-1][0] + t, b, i, u)
        else:
            merged.append((t, b, i, u))

    def esc(t):
        return t.replace("\\", "\\\\").replace("*", "\\*")

    def fmt(group):
        out, stack = [], []
        pieces = []
        for t, b, i, _ in group:
            lead = t[: len(t) - len(t.lstrip(" "))]
            trail = t[len(t.rstrip(" ")):] if t.strip() else ""
            core = t.strip(" ") if t.strip() else t
            if lead and t.strip():
                pieces.append((lead, None, None))
            pieces.append((core, b, i))
            if trail:
                pieces.append((trail, None, None))
        for t, b, i in pieces:
            if b is None:  # bare whitespace keeps the current formatting open only if it continues
                out.append(("ws", t))
                continue
            want = []
            if b:
                want.append("**")
            if i:
                want.append("*")
            # close markers not wanted
            while any(m not in want for m in stack):
                m = stack.pop()
                out.append(("close", m))
            for m in want:
                if m not in stack:
                    stack.append(m)
                    out.append(("open", m))
            out.append(("text", esc(t)))
        while stack:
            out.append(("close", stack.pop()))
        # move whitespace that sits just before a close marker to after it
        s, k = [], 0
        res = []
        for kind, v in out:
            res.append((kind, v))
        text = ""
        pending_ws = ""
        for kind, v in res:
            if kind == "ws":
                pending_ws += v
            elif kind == "close":
                text += v
            elif kind == "open":
                text += pending_ws + v
                pending_ws = ""
            else:
                text += pending_ws + v
                pending_ws = ""
        return text + pending_ws

    out, k = "", 0
    while k < len(merged):
        u = merged[k][3]
        j = k
        while j < len(merged) and merged[j][3] == u:
            j += 1
        inner = fmt(merged[k:j])
        if u:
            lead = inner[: len(inner) - len(inner.lstrip(" "))]
            trail = inner[len(inner.rstrip(" ")):]
            out += f"{lead}[{inner.strip(' ')}]({u}){trail}"
        else:
            out += inner
        k = j
    out = re.sub(r" *\n *", "\n", re.sub(r" {2,}", " ", out)).strip()
    return out


def style_of(p_el, d):
    ppr = p_el.find(qn("w:pPr"))
    sid = None
    if ppr is not None and ppr.find(qn("w:pStyle")) is not None:
        sid = ppr.find(qn("w:pStyle")).get(qn("w:val"))
    if sid is None:
        return "Normal"
    try:
        return d.styles.get_by_id(sid, 1).name
    except Exception:
        return sid


def read_docx(path, slug_hint=None, new_image_dir=None):
    """docx -> dict(title, meta..., body markdown, new_images[(relpath, blob)])"""
    d = docx.Document(path)
    part = d.part
    body_el = d.element.body
    title, meta, blocks = None, {}, []
    new_images = []
    img_counter = [0]
    for el in body_el:
        if el.tag == qn("w:tbl") and not meta:
            tbl = docx.table.Table(el, d)
            for row in tbl.rows:
                label = row.cells[0].text.strip()
                paras = []
                for p in row.cells[1].paragraphs:
                    segs, _ = read_paragraph(p._p, part)
                    text = "".join(s[0] for s in segs).strip()
                    urls = [s[3] for s in segs if s[3]]
                    if text:
                        paras.append((text, urls[0] if urls else None))
                meta[label] = paras
            continue
        if el.tag != qn("w:p"):
            continue
        style = style_of(el, d)
        segs, images = read_paragraph(el, part)
        if title is None and style == "Title":
            title = "".join(s[0] for s in segs).strip()
            continue
        for name, alt, blob, ext in images:
            if name.startswith("/assets/"):
                src = '{{ "' + name + '" | relative_url }}'
                if not os.path.exists(os.path.join(ROOT, name.lstrip("/"))) and blob:
                    new_images.append((name, blob))
            elif name.startswith("raw:"):
                src = name[4:]
            else:
                img_counter[0] += 1
                h = hashlib.sha1(blob or b"").hexdigest()[:8]
                rel = f"/assets/images/{slug_hint or 'new'}-{h}{ext.lower()}"
                new_images.append((rel, blob))
                src = '{{ "' + rel + '" | relative_url }}'
            blocks.append(("image_cont" if style in ("Image Under Item", "List Continue") else "image",
                           f"![{alt}]({src})"))
        md = emit_inline(segs)
        if not md:
            continue
        if style == "Fleuron" or md in ("❧",):
            blocks.append(("fleuron", FLEURON_MD))
        elif style == "Divider" or md in (DIVIDER, "---", "***", "* * *"):
            blocks.append(("divider", "---"))
        elif style.startswith("Heading") and style[-1] in "123":
            blocks.append(("h", "#" * int(style[-1]) + " " + md))
        elif style == "Quote Numbered":
            blocks.append(("qlist", md))
        elif style == "Quote Bullet":
            blocks.append(("qlist", "- " + md))
        elif style in ("Quote", "Intense Quote"):
            blocks.append(("quote", md))
        elif style in ("List Bullet", "List Bullet 2", "List Paragraph"):
            blocks.append(("bullet", "- " + md))
        elif style == "List Bullet Spaced":
            blocks.append(("bullet_loose", "- " + md))
        elif style in ("Numbered Item", "List Number"):
            blocks.append(("num", md))
        else:
            blocks.append(("para", md))
    # join blocks
    out = []
    for k, (kind, text) in enumerate(blocks):
        prev = blocks[k - 1][0] if k else None
        text = text.replace("\n", "\\\\\n" + ("> " if kind in ("quote", "qlist") else ""))
        if kind in ("quote", "qlist"):
            text = "> " + text
            sep = ("\n" if kind == prev == "qlist" else "\n>\n") if prev in ("quote", "qlist") else "\n\n"
        elif kind in ("bullet", "num") and prev == kind:
            sep = "\n"
        elif kind == "image_cont":
            text = "    " + text
            sep = "\n\n"
        else:
            sep = "\n\n"
        out.append((sep if out else "") + text)
    body = "".join(out) + "\n"
    return {"title": title or "", "meta": meta, "body": body, "new_images": new_images}


# ---------------------------------------------------------------- front matter merge

def meta_values(meta):
    def texts(label):
        return [t for t, _ in meta.get(label, [])]
    topics = []
    for t in texts("Topics"):
        topics += [x.strip().replace("'", "’") for x in t.split(",") if x.strip()]
    related = []
    for t, u in meta.get("Related", []):
        slug = None
        if u:
            m = re.search(r"/treasury/([^/]+)/?$", u) or re.search(r"/([^/]+)/?$", u)
            slug = m.group(1) if m else None
        related.append((t, slug))
    live = meta.get("Live page", [])
    slug = None
    if live:
        m = re.search(r"/treasury/([^/]+)/?", live[0][1] or live[0][0])
        slug = m.group(1) if m else None
    return {"topics": topics, "related": related, "figuring": texts("Figuring out"), "slug": slug}


def build_fm(old_fm, e):
    """Rewrite only the fields that changed, keeping the rest of the front matter verbatim."""
    fm = old_fm
    old = {}
    for key in ("title", "updated"):
        m = re.search(rf"^{key}:\s*(.*)$", fm, re.M)
        old[key] = unquote(m.group(1)) if m else None
    if e["title"] != old["title"]:
        fm = re.sub(r"^title:.*$", lambda m: "title: " + yq(e["title"]), fm, count=1, flags=re.M)
    if e["updated"] != old["updated"]:
        fm = re.sub(r"^updated:.*$", lambda m: "updated: " + yq(e["updated"]), fm, count=1, flags=re.M)
    cur_topics = re.search(r"^topics:\s*\[(.*)\]\s*$", fm, re.M)
    if [unquote(t) for t in cur_topics.group(1).split(",") if t.strip()] != e["topics"]:
        fm = re.sub(r"^topics:.*$", lambda m: "topics: [" + ", ".join(e["topics"]) + "]", fm, count=1, flags=re.M)
    rel_now = [(unquote(t), u.strip("/")) for t, u in
               re.findall(r'^  - title:\s*(.*)\n    url:\s*"?([^"\n]*)"?', fm, re.M)]
    if rel_now != e["related"]:
        block = "related:\n" + "".join(f'  - title: {yq(t)}\n    url: "/{u}/"\n' for t, u in e["related"])
        fm = re.sub(r"^related:\n(?:  .*\n?)*", lambda m: block, fm + "\n", count=1, flags=re.M).rstrip("\n")
    fig_m = re.search(r"^figuring_out:\n((?:  - .*\n?)*)", fm + "\n", re.M)
    fig_now = [unquote(x) for x in re.findall(r"^  - (.*)$", fig_m.group(1), re.M)] if fig_m else []
    if fig_now != e["figuring"]:
        block = ("figuring_out:\n" + "".join(f"  - {yq(x)}\n" for x in e["figuring"])) if e["figuring"] else ""
        if fig_m:
            fm = re.sub(r"^figuring_out:\n(?:  - .*\n?)*", lambda m: block, fm + "\n", count=1, flags=re.M).rstrip("\n")
        else:
            fm = fm.rstrip("\n") + "\n" + block.rstrip("\n")
    return fm


# ---------------------------------------------------------------- commands

def docx_files():
    out = []
    for f in sorted(glob.glob(os.path.join(DOCS, "*.docx"))):
        b = os.path.basename(f)
        if b.startswith("~$") or b == TEMPLATE_NAME:
            continue
        out.append(f)
    return out


def docx_slug(path):
    try:
        d = docx.Document(path)
    except Exception:
        return None
    s = d.core_properties.subject
    if s:
        return s
    info = read_docx(path)
    return meta_values(info["meta"])["slug"]


def cmd_export(slugs, force=False):
    os.makedirs(DOCS, exist_ok=True)
    state = load_state()
    entries = all_entries()
    targets = slugs or list(entries)
    by_slug = {}
    for f in docx_files():
        s = docx_slug(f)
        if s:
            by_slug[s] = f
    done, skipped = [], []
    for slug in targets:
        e = entries[slug]
        md_path = os.path.join(ROOT, "_pages", slug + ".md")
        existing = by_slug.get(slug)
        if existing:
            st = state.get(slug, {})
            if not force and st.get("docx_sha") and sha(existing) != st["docx_sha"]:
                skipped.append(f"{slug}: Word file has edits not yet imported")
                continue
            if locked(existing):
                skipped.append(f"{slug}: Word file is open (lock file present)")
                continue
        target = os.path.join(DOCS, filename_for(e["title"]))
        if existing and existing != target:
            os.remove(existing)
        export_one(slug, e, target)
        state[slug] = {"docx": os.path.basename(target), "docx_sha": sha(target), "md_sha": sha(md_path)}
        done.append(slug)
    tpl = os.path.join(DOCS, TEMPLATE_NAME)
    if not os.path.exists(tpl):
        export_one("", {"title": "Title of the new entry", "topics": [], "related": [], "figuring": [],
                        "created": "", "updated": "", "id": "", "body":
                        "Body text goes here.\n\n> A quotation goes in the Quote style.\n\n"
                        "— Author, *Title* (Publisher, year).\n"}, tpl)
    save_state(state)
    print(f"exported {len(done)}" + ("".join("\n  skipped " + s for s in skipped)))
    return done, skipped


def cmd_status(verbose=True):
    state = load_state()
    entries = all_entries()
    docx_edits, md_edits, conflicts, new_docs, missing = [], [], [], [], []
    seen = set()
    for f in docx_files():
        s = docx_slug(f)
        if not s:
            new_docs.append(f)
            continue
        seen.add(s)
        st = state.get(s, {})
        md_path = os.path.join(ROOT, "_pages", s + ".md")
        dchg = sha(f) != st.get("docx_sha")
        mchg = os.path.exists(md_path) and sha(md_path) != st.get("md_sha")
        if dchg and mchg:
            conflicts.append(s)
        elif dchg:
            docx_edits.append((s, f))
        elif mchg:
            md_edits.append(s)
    for s in entries:
        if s not in seen:
            missing.append(s)
    if verbose:
        for label, xs in (("Word edits to import", [s for s, _ in docx_edits]), ("new Word entries", [os.path.basename(f) for f in new_docs]),
                          (".md changed since export (re-export)", md_edits), ("CONFLICT (both changed)", conflicts),
                          ("entries with no Word file", missing)):
            if xs:
                print(f"{label}: " + ", ".join(xs))
        if not any((docx_edits, new_docs, md_edits, conflicts, missing)):
            print("in sync")
    return docx_edits, new_docs, md_edits, conflicts, missing


def cmd_import():
    state = load_state()
    entries = all_entries()
    tax = set(taxonomy())
    titles = {e["title"].lower(): s for s, e in entries.items()}
    docx_edits, new_docs, _, conflicts, _ = cmd_status(verbose=False)
    changed, problems = [], []
    for s in conflicts:
        problems.append(f"{s}: both the Word file and the .md changed since the last sync; resolve by hand")
    jobs = [(s, f, False) for s, f in docx_edits] + [(None, f, True) for f in new_docs]
    for slug, f, is_new in jobs:
        if locked(f):
            problems.append(f"{os.path.basename(f)}: open in Word with unsaved state? (lock file present) — save and close, then rerun")
        info = read_docx(f, slug_hint=slug)
        mv = meta_values(info["meta"])
        title = info["title"].strip()
        if not title:
            problems.append(f"{os.path.basename(f)}: no Title paragraph")
            continue
        if is_new:
            slug = slugify(title)
            if slug in entries:
                problems.append(f"{os.path.basename(f)}: an entry with slug '{slug}' already exists")
                continue
            info = read_docx(f, slug_hint=slug)
        bad = [t for t in mv["topics"] if t not in tax]
        if bad or not mv["topics"]:
            problems.append(f"{slug}: unknown or missing topics {bad or ''} — use the taxonomy names")
            continue
        related = []
        for t, rs in mv["related"]:
            rs = rs or titles.get(t.lower())
            if not rs or rs not in entries:
                problems.append(f"{slug}: related entry '{t}' not found")
                related = None
                break
            related.append((t, rs))
        if related is None:
            continue
        for rel, blob in info["new_images"]:
            p = os.path.join(ROOT, rel.lstrip("/"))
            os.makedirs(os.path.dirname(p), exist_ok=True)
            with open(p, "wb") as fh:
                fh.write(blob)
        md_path = os.path.join(ROOT, "_pages", slug + ".md")
        if is_new:
            nid = max(int(e["id"]) for e in entries.values() if str(e["id"]).isdigit()) + 1
            fm = (f"layout: default\nid: {nid}\ntitle: {yq(title)}\ncreated: {yq(today_str())}\n"
                  f"updated: {yq(today_str())}\ntopics: [{', '.join(mv['topics'])}]\nrelated:\n" +
                  "".join(f'  - title: {yq(t)}\n    url: "/{u}/"\n' for t, u in related) +
                  ("figuring_out:\n" + "".join(f"  - {yq(x)}\n" for x in mv["figuring"]) if mv["figuring"] else ""))
            fm = fm.rstrip("\n")
            entries[slug] = {"id": str(nid)}
        else:
            old = entries[slug]
            content_changed = info["body"].strip() != old["body"].strip() or title != old["title"]
            e = {"title": title, "topics": mv["topics"], "related": related, "figuring": mv["figuring"],
                 "updated": today_str() if content_changed else old["updated"]}
            fm = build_fm(old["fm"], e)
        text = "---\n" + fm + "\n---\n\n" + info["body"].lstrip("\n")
        if not is_new and text == open(md_path, encoding="utf-8").read():
            state.setdefault(slug, {})["docx_sha"] = sha(f)
            continue
        with open(md_path, "w", encoding="utf-8") as fh:
            fh.write(text)
        state[slug] = {"docx": os.path.basename(f), "docx_sha": sha(f), "md_sha": sha(md_path)}
        changed.append(slug)
        if is_new and not locked(f):
            # give the new Word file its entry link and canonical name
            os.remove(f)
            save_state(state)
            cmd_export([slug], force=True)
            state = load_state()
    save_state(state)
    print("imported: " + (", ".join(changed) if changed else "nothing"))
    for p in problems:
        print("PROBLEM " + p)
    return changed, problems


def cmd_publish():
    changed, problems = cmd_import()
    if problems:
        print("Not publishing until the problems above are fixed.")
        sys.exit(1)
    # also publish .md changes already imported (e.g. `import` was run first, then checked)
    pending = subprocess.run(["git", "status", "--porcelain", "_pages"], cwd=ROOT,
                             capture_output=True, text=True).stdout.split()
    changed = changed or [os.path.basename(p)[:-3] for p in pending if p.endswith(".md")]
    if not changed:
        print("nothing to publish")
        return
    subprocess.run([sys.executable, "_tools/sync.py"], cwd=ROOT, check=True)
    subprocess.run([sys.executable, "_tools/sync.py", "--check"], cwd=ROOT, check=True)
    # sync.py touched only index/topics/entries; record md hashes again in case
    state = load_state()
    for s in changed:
        if s in state:
            state[s]["md_sha"] = sha(os.path.join(ROOT, "_pages", s + ".md"))
    save_state(state)
    subprocess.run(["git", "add", "-A"], cwd=ROOT, check=True)
    msg = "Word mirror: update " + ", ".join(changed) + "\n\nCo-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
    subprocess.run(["git", "commit", "-qm", msg], cwd=ROOT, check=True)
    subprocess.run(["gh", "auth", "switch", "-u", "lauramoorewrites-del"], cwd=ROOT, capture_output=True)
    subprocess.run(["git", "push", "-q", "origin", "main"], cwd=ROOT, check=True)
    print("pushed: " + ", ".join(changed))


def norm(body):
    """Differences that render identically: '> ' vs '>' blank quote lines, trailing/double spaces,
    _x_ / <i>x</i> vs *x*, *[x](u)* vs [*x*](u), soft-wrapped lines, adjacent blockquotes."""
    b = body.strip()
    b = re.sub(r"<i>(.*?)</i>", r"*\1*", b)
    b = re.sub(r"(?<![\w*])_([^_\s][^_]*?)_(?!\w)", r"*\1*", b)
    b = re.sub(r"\*\[([^\]]*)\]\(([^)]*)\)\*", r"[*\1*](\2)", b)
    b = re.sub(r"[ \t]+$", "", b, flags=re.M)
    b = re.sub(r"(?<=\S) {2,}", " ", b)
    b = re.sub(r"^(>[^\n]*)\n\n(?=>)", r"\1\n>\n", b, flags=re.M)
    b = re.sub(r"\n(?=[^\n>\-#\d ])(?<=[^\n]\n)", " ", b)
    b = re.sub(r"\n> (?=[^\n\d-])(?<=[^>\n]\n> )", " ", b)
    return b


def cmd_roundtrip():
    entries = all_entries()
    bad = 0
    with tempfile.TemporaryDirectory() as tmp:
        for slug, e in entries.items():
            p = os.path.join(tmp, slug + ".docx")
            export_one(slug, e, p)
            info = read_docx(p, slug_hint=slug)
            mv = meta_values(info["meta"])
            problems = []
            if norm(info["body"]) != norm(e["body"]):
                problems.append("body")
            if info["title"] != e["title"]:
                problems.append("title")
            if mv["topics"] != e["topics"]:
                problems.append(f"topics {mv['topics']} != {e['topics']}")
            if [r for _, r in mv["related"]] != [u for _, u in e["related"]] or \
               [t for t, _ in mv["related"]] != [t for t, _ in e["related"]]:
                problems.append("related")
            if mv["figuring"] != e["figuring"]:
                problems.append("figuring")
            if mv["slug"] != slug:
                problems.append("slug")
            if info["new_images"]:
                problems.append("images")
            if problems:
                bad += 1
                print(f"== {slug}: {problems}")
                if "body" in problems:
                    a = e["body"].strip().split("\n")
                    b = info["body"].strip().split("\n")
                    import difflib
                    for l in list(difflib.unified_diff(a, b, lineterm="", n=0))[2:12]:
                        print("   " + l[:160])
    print(f"{len(entries) - bad}/{len(entries)} round-trip cleanly")


if __name__ == "__main__":
    args = sys.argv[1:]
    cmd = args[0] if args else "status"
    if cmd == "export":
        cmd_export([a for a in args[1:] if not a.startswith("--")], force="--force" in args)
    elif cmd == "import":
        cmd_import()
    elif cmd == "publish":
        cmd_publish()
    elif cmd == "roundtrip":
        cmd_roundtrip()
    else:
        cmd_status()
