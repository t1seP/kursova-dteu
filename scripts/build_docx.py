"""Збірка курсової ДТЕУ: kursova.md -> .docx за ДСТУ 3008:2015 і методичками ДТЕУ.

Використання:
    py build_docx.py kursova.md                      # -> build/Курсова_<Прізвище>.docx
    py build_docx.py kursova.md -o out.docx --update --pdf

--update   відкрити у Word: оновити зміст, порахувати сторінки, підставити {{N_STOR}}
--pdf      додатково зберегти PDF (потрібен Word)

Розмітка описана в shablony.md. Шапка між рядками '---' задає титулку і варіант оформлення.
"""
import argparse
import os
import re
import sys

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Length, Mm, Pt, RGBColor

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import word_com  # noqa: E402

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

FONT = "Times New Roman"
BODY_PT = 14
TABLE_PT = 12
LINE = 1.5
INDENT = Cm(1.25)
TEXT_WIDTH = Mm(175)  # 210 - 25 - 10
BLANK_LINE = Pt(21)   # один вільний рядок при 14 пт і інтервалі 1,5

STRUCTURAL = {"ЗМІСТ", "ПЕРЕЛІК УМОВНИХ ПОЗНАЧЕНЬ", "ВСТУП", "ВИСНОВКИ ТА ПРОПОЗИЦІЇ",
              "ВИСНОВКИ І ПРОПОЗИЦІЇ", "СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ", "ДОДАТКИ"}
UA_ORDER = "АБВГҐДЕЄЖЗИІЇЙКЛМНОПРСТУФХЦЧШЩЬЮЯ"


# ---------------------------------------------------------------- шапка і розбір

def parse_front_matter(text):
    meta = {}
    if not text.startswith("---"):
        return meta, text
    end = text.find("\n---", 3)
    if end == -1:
        return meta, text
    block = text[3:end]
    body = text[end + 4:].lstrip("\n")
    for line in block.splitlines():
        if line.lstrip().startswith("#"):
            continue
        line = re.sub(r"\s+#.*$", "", line).rstrip()
        if not line.strip() or ":" not in line:
            continue
        k, v = line.split(":", 1)
        meta[k.strip()] = v.strip().strip('"')
    return meta, body


def sort_key_ua(s):
    s = re.sub(r"^[\W\d_]+", "", s).upper()
    key = []
    for ch in s:
        if ch in UA_ORDER:
            key.append((0, UA_ORDER.index(ch)))
        elif "A" <= ch <= "Z":
            key.append((1, ord(ch)))
        elif ch.isalpha():
            key.append((0, 100 + ord(ch)))
        else:
            key.append((-1, 0))
    return key


def split_references(body):
    """Повертає (текст до списку, список записів [(ключ|None, текст)], текст після списку)."""
    m = re.search(r"^#\s+СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ\s*$", body, re.M)
    if not m:
        return body, None, ""
    rest = body[m.end():]
    m2 = re.search(r"^#\s+\S", rest, re.M)
    refs_block = rest[:m2.start()] if m2 else rest
    after = rest[m2.start():] if m2 else ""
    entries = []
    for line in refs_block.splitlines():
        s = line.strip()
        if not s or s == "[[PAGEBREAK]]":
            continue
        mk = re.match(r"^\[@([^\]]+)\]\s*(.+)$", s)
        mn = re.match(r"^\d+[.)]\s*(.+)$", s)
        if mk:
            entries.append([mk.group(1).strip(), mk.group(2).strip()])
        elif mn:
            entries.append([None, mn.group(1).strip()])
        elif entries:
            entries[-1][1] += " " + s
        else:
            entries.append([None, s])
    return body[:m.start()], entries, after


def resolve_citations(before, entries, after, order, sep):
    """Замінює [@ключ] на номери. order: alpha | mention."""
    keyed = [e for e in entries if e[0]]
    if not keyed:
        return before, entries, after, []
    warnings = []
    if order == "mention":
        seen = []
        for m in re.finditer(r"\[@([^\]]+)\]", before + after):
            for part in m.group(1).split(";"):
                k = part.split(",")[0].strip().lstrip("@")
                if k not in seen:
                    seen.append(k)
        pos = {k: i for i, k in enumerate(seen)}
        entries = sorted(entries, key=lambda e: pos.get(e[0], 10_000))
    else:
        entries = sorted(entries, key=lambda e: sort_key_ua(e[1]))
    num = {e[0]: i + 1 for i, e in enumerate(entries) if e[0]}

    def repl(m):
        parts = []
        for part in m.group(1).split(";"):
            bits = [b.strip() for b in part.split(",", 1)]
            k = bits[0].lstrip("@")
            if k not in num:
                warnings.append(f"невідомий ключ джерела: @{k}")
                parts.append((10_000, f"@{k}?"))
                continue
            label = str(num[k]) + (", " + bits[1] if len(bits) > 1 else "")
            parts.append((num[k], label))
        parts.sort()
        return "[" + (sep + " ").join(p[1] for p in parts) + "]"

    cited = set(re.findall(r"@([\w\-]+)", " ".join(re.findall(r"\[@[^\]]+\]", before + after))))
    for e in entries:
        if e[0] and e[0] not in cited:
            warnings.append(f"джерело без посилання в тексті: @{e[0]}")
    before = re.sub(r"\[(@[^\]]+)\]", repl, before)
    after = re.sub(r"\[(@[^\]]+)\]", repl, after)
    return before, entries, after, warnings


# ---------------------------------------------------------------- форматування docx

def set_fonts(el_with_rpr, name=FONT):
    rpr = el_with_rpr.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(a), name)
    for a in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rfonts.get(qn(a)) is not None:
            del rfonts.attrib[qn(a)]
    lang = rpr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rpr.append(lang)
    lang.set(qn("w:val"), "uk-UA")
    lang.set(qn("w:eastAsia"), "uk-UA")


def style_paragraph(pf, align=None, indent=True, line=LINE, before=0, after=0):
    if align is not None:
        pf.alignment = align
    pf.first_line_indent = INDENT if indent else Cm(0)
    pf.left_indent = Cm(0)
    pf.space_before = before if isinstance(before, Length) else Pt(before)
    pf.space_after = after if isinstance(after, Length) else Pt(after)
    if line == 1:
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    else:
        pf.line_spacing = line


def setup_styles(doc):
    normal = doc.styles["Normal"]
    normal.font.size = Pt(BODY_PT)
    normal.font.color.rgb = RGBColor(0, 0, 0)
    set_fonts(normal.element)
    style_paragraph(normal.paragraph_format, WD_ALIGN_PARAGRAPH.JUSTIFY)
    normal.paragraph_format.widow_control = True

    for name, level in (("Heading 1", 1), ("Heading 2", 2), ("Heading 3", 3)):
        st = doc.styles[name]
        st.font.size = Pt(BODY_PT)
        st.font.bold = True
        st.font.italic = False
        st.font.color.rgb = RGBColor(0, 0, 0)
        set_fonts(st.element)
        pf = st.paragraph_format
        pf.keep_with_next = True
        pf.line_spacing = LINE
        pf.first_line_indent = Cm(0)
        pf.space_before = Pt(0)
        pf.space_after = BLANK_LINE
    doc.styles["Heading 1"].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h2 = doc.styles["Heading 2"].paragraph_format
    h2.alignment = WD_ALIGN_PARAGRAPH.LEFT
    h2.first_line_indent = INDENT
    h2.space_before = BLANK_LINE

    format_toc_styles(doc)


def format_toc_styles(doc):
    """Стилі рядків змісту. Word під час оновлення поля підміняє їх вбудованими,
    тому функцію викликаємо і до, і після оновлення у Word."""
    for i, name in enumerate(("toc 1", "toc 2")):
        try:
            st = doc.styles[name]
        except KeyError:
            st = doc.styles.add_style(name, 1)
            st.element.styleId = f"TOC{i + 1}"
        if st.element.get(qn("w:customStyle")) is not None:
            del st.element.attrib[qn("w:customStyle")]
        st.base_style = doc.styles["Normal"]
        st.hidden = False
        st.unhide_when_used = True
        pf = st.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.first_line_indent = Cm(0)
        pf.left_indent = Cm(0.5 * i)
        pf.right_indent = Cm(1)
        pf.space_after = Pt(0)
        pf.line_spacing = LINE
        pf.tab_stops.clear_all()
        pf.tab_stops.add_tab_stop(TEXT_WIDTH, WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)


def setup_section(doc):
    s = doc.sections[0]
    s.orientation = WD_ORIENT.PORTRAIT
    s.page_width, s.page_height = Mm(210), Mm(297)
    s.left_margin, s.right_margin = Mm(25), Mm(10)
    s.top_margin, s.bottom_margin = Mm(20), Mm(20)
    s.header_distance = Mm(10)
    s.footer_distance = Mm(10)
    s.different_first_page_header_footer = True
    p = s.header.paragraphs[0]
    style_paragraph(p.paragraph_format, WD_ALIGN_PARAGRAPH.RIGHT, indent=False, line=1)
    add_field(p, "PAGE", "2")


def add_field(paragraph, instr, placeholder=""):
    def fld(t):
        r = OxmlElement("w:r")
        c = OxmlElement("w:fldChar")
        c.set(qn("w:fldCharType"), t)
        if t == "begin":
            c.set(qn("w:dirty"), "true")
        r.append(c)
        return r
    run_el = paragraph.add_run()._r
    run_el.addprevious(fld("begin"))
    r = OxmlElement("w:r")
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = f" {instr} "
    r.append(it)
    run_el.addprevious(r)
    run_el.addprevious(fld("separate"))
    t = OxmlElement("w:t")
    t.text = placeholder
    run_el.append(t)
    run_el.addnext(fld("end"))


INLINE = re.compile(r"(\*\*.+?\*\*|\*[^*\s][^*]*?\*)")


def add_runs(p, text, size=None, bold=None):
    for chunk in INLINE.split(text):
        if not chunk:
            continue
        b, i = bold, None
        if chunk.startswith("**") and chunk.endswith("**"):
            chunk, b = chunk[2:-2], True
        elif chunk.startswith("*") and chunk.endswith("*") and len(chunk) > 2:
            chunk, i = chunk[1:-1], True
        r = p.add_run(chunk)
        if b:
            r.bold = True
        if i:
            r.italic = True
        if size:
            r.font.size = Pt(size)


def para(doc, text="", align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=True, size=None, bold=None,
         line=LINE, before=0, after=0, keep=False, page_break=False):
    p = doc.add_paragraph()
    style_paragraph(p.paragraph_format, align, indent, line, before, after)
    p.paragraph_format.keep_with_next = keep
    p.paragraph_format.page_break_before = page_break
    if text:
        add_runs(p, text, size, bold)
    return p


# ---------------------------------------------------------------- титулка і зміст

def blank(v, n=30):
    return v if v else "_" * n


def title_page(doc, m):
    C, L = WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT
    para(doc, "Державний торговельно-економічний університет", C, False, line=1)
    para(doc, blank(m.get("kafedra"), 40), C, False, line=1, after=96)
    para(doc, "КУРСОВА РОБОТА", C, False, bold=True, line=1, after=6)
    para(doc, f"з дисципліни «{blank(m.get('dyscyplina'), 25)}»", C, False, line=1, after=6)
    para(doc, "на тему:", C, False, line=1, after=6)
    para(doc, f"«{blank(m.get('tema'), 50)}»", C, False, bold=True, line=1, after=6)
    if m.get("baza") is not None:
        para(doc, f"(за матеріалами {blank(m.get('baza'), 40)})", C, False, line=1, after=60)
    block = Cm(8)
    stat = m.get("stat", "студента")
    lines = [
        f"{stat.capitalize()} {blank(m.get('kurs'), 3)} курсу {blank(m.get('hrupa'), 4)} групи",
        f"факультету {blank(m.get('fakultet'), 20)}",
        f"спеціальності {blank(m.get('specialnist'), 20)}",
        f"освітньої програми «{blank(m.get('prohrama'), 20)}»",
        blank(m.get("student"), 35),
        "",
        "Науковий керівник:",
        blank(m.get("kerivnyk"), 35),
    ]
    for ln in lines:
        p = para(doc, ln, L, False, line=1)
        p.paragraph_format.left_indent = block
    p = para(doc, "Кількість балів: ________   Оцінка за шкалою ЄКТС: _____", L, False, line=1, before=36)
    p = para(doc, "Члени комісії:", L, False, line=1, before=12)
    for _ in range(3):
        para(doc, "______________          ______________________________", L, False, line=1, before=10)
        p = para(doc, "(підпис)                              (прізвище та ініціали)", L, False, size=10, line=1)
    # місто і рік у колонтитулі першої сторінки: завжди внизу, незалежно від довжини теми
    fp = doc.sections[0].first_page_footer.paragraphs[0]
    style_paragraph(fp.paragraph_format, C, indent=False, line=1)
    fp.add_run(f"{m.get('misto', 'Київ')} {blank(m.get('rik'), 4)}")
    p.add_run().add_break(WD_BREAK.PAGE)


def toc(doc):
    para(doc, "ЗМІСТ", WD_ALIGN_PARAGRAPH.CENTER, False, bold=True, after=BLANK_LINE)
    p = para(doc, "", WD_ALIGN_PARAGRAPH.LEFT, False)
    add_field(p, r'TOC \o "1-2" \h \z \u',
              "Зміст з'явиться після оновлення: правою кнопкою по цьому тексту, «Оновити поле» (або build_docx.py --update).")


# ---------------------------------------------------------------- тіло

def add_table(doc, label, title, rows, variant):
    lab = para(doc, label, WD_ALIGN_PARAGRAPH.RIGHT if variant == "A" else WD_ALIGN_PARAGRAPH.LEFT,
               False, before=BLANK_LINE, keep=True)
    para(doc, title, WD_ALIGN_PARAGRAPH.CENTER, False, keep=True, line=LINE)
    ncols = max(len(r) for r in rows)
    t = doc.add_table(rows=len(rows), cols=ncols)
    t.style = doc.styles["Table Grid"]
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for ri, row in enumerate(rows):
        for ci in range(ncols):
            cell = t.cell(ri, ci)
            text = row[ci] if ci < len(row) else ""
            p = cell.paragraphs[0]
            style_paragraph(p.paragraph_format,
                            WD_ALIGN_PARAGRAPH.CENTER if ri == 0 or ci > 0 else WD_ALIGN_PARAGRAPH.LEFT,
                            False, 1)
            add_runs(p, text, TABLE_PT, True if ri == 0 else None)
    tr_pr = t.rows[0]._tr.get_or_add_trPr()
    hdr = OxmlElement("w:tblHeader")
    hdr.set(qn("w:val"), "true")
    tr_pr.append(hdr)
    return lab


def note(doc, text):
    return para(doc, text, WD_ALIGN_PARAGRAPH.LEFT, False, size=TABLE_PT, line=1, after=BLANK_LINE)


def add_figure(doc, path, caption, width_cm, base_dir):
    full = path if os.path.isabs(path) else os.path.join(base_dir, path)
    p = para(doc, "", WD_ALIGN_PARAGRAPH.CENTER, False, line=1, before=BLANK_LINE, keep=True)
    if os.path.exists(full):
        shape = p.add_run().add_picture(full, width=Cm(width_cm) if width_cm else None)
        maxw = Cm(16.5)
        if shape.width > maxw:
            ratio = maxw / shape.width
            shape.width = int(shape.width * ratio)
            shape.height = int(shape.height * ratio)
    else:
        add_runs(p, f"[РИСУНОК НЕ ЗНАЙДЕНО: {path}]")
    para(doc, caption, WD_ALIGN_PARAGRAPH.CENTER, False)


def add_formula(doc, formula, number):
    p = para(doc, "", WD_ALIGN_PARAGRAPH.LEFT, False, before=Pt(14), after=Pt(14))
    ts = p.paragraph_format.tab_stops
    ts.add_tab_stop(Mm(87.5), WD_TAB_ALIGNMENT.CENTER)
    ts.add_tab_stop(TEXT_WIDTH, WD_TAB_ALIGNMENT.RIGHT)
    add_runs(p, f"\t{formula}\t{number}")


def render(doc, body, meta, entries, base_dir):
    variant = meta.get("variant", "A").upper()
    lines = body.splitlines()
    i = 0
    buf = []
    section = None
    after_formula = False
    stats = {"tables": 0, "figures": 0, "appendices": 0, "main_chars": 0}
    in_main = False

    def flush():
        nonlocal buf
        if buf:
            text = " ".join(s.strip() for s in buf)
            para(doc, text)
            if in_main:
                stats["main_chars"] += len(text)
            buf = []

    while i < len(lines):
        raw = lines[i]
        s = raw.strip()
        if not s:
            flush()
            after_formula = False
            i += 1
            continue
        if s == "[[PAGEBREAK]]":
            flush()
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
            i += 1
            continue
        m = re.match(r"^(#{1,3})\s+(.+?)\s*$", s)
        if m:
            flush()
            level, text = len(m.group(1)), m.group(2).strip().rstrip(".")
            if level == 1:
                up = text.upper()
                section = up
                if up == "ЗМІСТ":
                    i += 1
                    continue
                if up == "ВСТУП":
                    in_main = True
                if up.startswith("СПИСОК"):
                    in_main = False
                h = doc.add_heading(up, level=1)
                structural = up in STRUCTURAL
                h.paragraph_format.page_break_before = structural or variant == "A"
                if up == "СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ" and entries is not None:
                    for n, (_, t) in enumerate(entries, 1):
                        para(doc, f"{n}. {t}", WD_ALIGN_PARAGRAPH.LEFT)  # по ширині довгі URL дають дірки
            elif level == 2 and section == "ДОДАТКИ":
                stats["appendices"] += 1
                para(doc, text, WD_ALIGN_PARAGRAPH.CENTER, False, page_break=True, keep=True)
            elif level == 3 and section == "ДОДАТКИ":
                para(doc, text, WD_ALIGN_PARAGRAPH.CENTER, False, bold=True, keep=True, after=BLANK_LINE)
            else:
                doc.add_heading(text, level=2)
            i += 1
            continue
        mt = re.match(r"^(Таблиця\s+[\wА-Яа-яІіЇїЄєҐґ.]+)\s*[:.]\s*(.+)$", s)
        if mt and i + 1 < len(lines) and lines[i + 1].strip().startswith("|"):
            flush()
            rows = []
            i += 1
            while i < len(lines) and lines[i].strip().startswith("|"):
                row = lines[i].strip().strip("|")
                if not re.match(r"^\s*:?-{2,}", row):
                    rows.append([c.strip() for c in row.split("|")])
                i += 1
            add_table(doc, mt.group(1), mt.group(2), rows, variant)
            stats["tables"] += 1
            if i < len(lines) and re.match(r"^(Джерело|Примітка)", lines[i].strip()):
                note(doc, lines[i].strip())
                i += 1
            else:
                note(doc, "")
            continue
        mf = re.match(r"^!\[(.+?)\]\((.+?)\)(?:\{w=(\d+(?:\.\d+)?)\})?$", s)
        if mf:
            flush()
            add_figure(doc, mf.group(2), mf.group(1), float(mf.group(3)) if mf.group(3) else None, base_dir)
            stats["figures"] += 1
            i += 1
            if i < len(lines) and re.match(r"^(Джерело|Примітка)", lines[i].strip()):
                note(doc, lines[i].strip()).alignment = WD_ALIGN_PARAGRAPH.CENTER
                i += 1
            else:
                note(doc, "")
            continue
        mq = re.match(r"^\$\$\s*(.+?)\s*\$\$\s*(\(.+\))?\s*,?$", s)
        if mq:
            flush()
            add_formula(doc, mq.group(1), mq.group(2) or "")
            after_formula = True
            i += 1
            continue
        if after_formula:
            p = para(doc, s, WD_ALIGN_PARAGRAPH.LEFT, False)
            if not s.startswith("де "):
                p.paragraph_format.left_indent = Cm(0.6)
            i += 1
            continue
        if re.match(r"^(- |\d+\) |\d+\. |[а-яі]\) )", s):
            flush()
            para(doc, s if not s.startswith("- ") else "- " + s[2:])
            if in_main:
                stats["main_chars"] += len(s)
            i += 1
            continue
        if re.match(r"^(Джерело|Примітка):", s):
            flush()
            note(doc, s)
            i += 1
            continue
        buf.append(s)
        i += 1
    flush()
    return stats


def fill_placeholders(docx_path, values):
    """Підставляє числа у {{...}} і відновлює стилі змісту після Word. Повертає True, якщо був запис."""
    doc = Document(docx_path)
    format_toc_styles(doc)
    changed = True
    for p in doc.paragraphs:
        for r in p.runs:
            for k, v in values.items():
                if k in r.text:
                    r.text = r.text.replace(k, str(v))
    doc.save(docx_path)
    return changed


def main():
    ap = argparse.ArgumentParser(description="kursova.md -> docx за вимогами ДТЕУ")
    ap.add_argument("src")
    ap.add_argument("-o", "--out")
    ap.add_argument("--update", action="store_true", help="оновити зміст і сторінки через Word")
    ap.add_argument("--pdf", action="store_true", help="зберегти PDF через Word")
    a = ap.parse_args()

    with open(a.src, encoding="utf-8-sig") as f:
        meta, body = parse_front_matter(f.read())
    base_dir = os.path.dirname(os.path.abspath(a.src))
    before, entries, after = split_references(body)
    warnings = []
    if entries is not None:
        before, entries, after, warnings = resolve_citations(
            before, entries, after, meta.get("sources_order", "alpha"), meta.get("cite_sep", ";"))
        body = before + "# СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ\n\n" + after

    out = a.out
    if not out:
        surname = (meta.get("student") or "студент").split()[0]
        out = os.path.join(base_dir, "build", f"Курсова_{surname}.docx")
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)

    doc = Document()
    setup_styles(doc)
    setup_section(doc)
    title_page(doc, meta)
    toc(doc)
    stats = render(doc, body, meta, entries, base_dir)
    doc.save(out)

    n_src = len(entries) if entries else 0
    est = round(stats["main_chars"] / 1800 + stats["tables"] * 0.4 + stats["figures"] * 0.5)
    values = {"{{N_DZHEREL}}": n_src, "{{N_TABL}}": stats["tables"], "{{N_RYS}}": stats["figures"],
              "{{N_DOD}}": stats["appendices"], "{{N_STOR}}": est}

    pages_info = ""
    if a.update or a.pdf:
        res = word_com.run_word(out, update=True)
        if res and res.get("ok"):
            main_pages = word_com.main_text_pages(res.get("headings"))
            if main_pages:
                values["{{N_STOR}}"] = main_pages
            pages_info = f"Word: усього {res.get('pages')} стор., основний текст {main_pages} стор."
            if fill_placeholders(out, values):
                res = word_com.run_word(out, update=True)
            if a.pdf:
                word_com.run_word(out, update=False, pdf_path=os.path.splitext(out)[0] + ".pdf")
        else:
            pages_info = "Word недоступний або помилка: " + str(res.get("error") if res else "немає Word")
            fill_placeholders(out, values)
    else:
        fill_placeholders(out, values)

    print(f"✅ Збережено: {os.path.abspath(out)}")
    print(f"   варіант {meta.get('variant', 'A')}, джерел {n_src}, таблиць {stats['tables']}, "
          f"рисунків {stats['figures']}, додатків {stats['appendices']}")
    print(f"   оцінка обсягу основного тексту: ~{est} стор. (1800 знаків/стор. + таблиці і рисунки)")
    if pages_info:
        print("   " + pages_info)
    else:
        print("   Зміст оновиться у Word (ПКМ по змісту, «Оновити поле») або запусти з --update")
    for w in warnings:
        print("⚠ " + w)


if __name__ == "__main__":
    main()
