"""Перевірка курсової ДТЕУ: оформлення, посилання, список ДСТУ 8302, вступ, обсяг.

Використання:
    py check_kursova.py Курсова.docx              # повна перевірка docx
    py check_kursova.py kursova.md                # текстові перевірки чернетки
    py check_kursova.py Курсова.docx --online     # + перевірити, що URL і DOI відповідають
    py check_kursova.py Курсова.docx --word       # + точні сторінки через Word
    --min-pages 25 --max-pages 35 --min-src 20 --max-src 40 --apa

Код виходу 1, якщо є хоча б одна помилка (❌).
"""
import argparse
import datetime as dt
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import word_com  # noqa: E402

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

REQUIRED = ["ЗМІСТ", "ВСТУП", "СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ"]
CONCL = ("ВИСНОВКИ ТА ПРОПОЗИЦІЇ", "ВИСНОВКИ І ПРОПОЗИЦІЇ")
STRUCT = set(REQUIRED) | set(CONCL) | {"ДОДАТКИ", "ПЕРЕЛІК УМОВНИХ ПОЗНАЧЕНЬ"}
CITIES = ("Київ|Харків|Львів|Одеса|Дніпро|Суми|Тернопіль|Житомир|Полтава|Чернівці|Ужгород|Вінниця|"
          "Запоріжжя|Луцьк|Рівне|Херсон|Миколаїв|Черкаси|Чернігів|Кропивницький|Ірпінь|Мукачево|"
          "Івано-Франківськ|Хмельницький|Біла Церква|Острог")
CLICHES = ["у сучасному світі", "в сучасному світі", "не секрет", "варто зазначити", "важливо підкреслити",
           "відіграє ключову роль", "невід'ємною частиною", "невід’ємною частиною", "в епоху",
           "всеосяжн", "синергі", "у сучасних реаліях", "в умовах динамічного розвитку",
           "слід зазначити", "необхідно відзначити"]
FIRST_PERSON = re.compile(r"(?<![\w'’])(я|мені|мною|мене|мій|моя|моє|мої|на мою думку)(?![\w'’])", re.I)
BAD_SOURCES = re.compile(r"wikipedia|wiki\.|referat|bestreferat|ukrreferat|studfile|studopedia|"
                         r"pidruchniki|allbest|kursovi|diplomy|essay", re.I)
CONCL_START = re.compile(r"^(Отже|Таким чином|Узагальнюючи|Підсумовуючи|Проведений|Проведене|"
                         r"Результати|Виходячи з|Загалом|Отож|У підсумку|Резюмуючи)", re.I)
UA_ORDER = "АБВГҐДЕЄЖЗИІЇЙКЛМНОПРСТУФХЦЧШЩЬЮЯ"


class Report:
    def __init__(self):
        self.items = []

    def ok(self, msg):
        self.items.append(("✅", msg))

    def warn(self, msg):
        self.items.append(("⚠", msg))

    def err(self, msg):
        self.items.append(("❌", msg))

    def print(self):
        for group in ("❌", "⚠", "✅"):
            for mark, msg in self.items:
                if mark == group:
                    print(f"{mark} {msg}")
        e = sum(1 for m, _ in self.items if m == "❌")
        w = sum(1 for m, _ in self.items if m == "⚠")
        print(f"\nПідсумок: помилок {e}, попереджень {w}, пройдено {sum(1 for m, _ in self.items if m == '✅')}")
        return e


# ---------------------------------------------------------------- читання

def load_md(path):
    with open(path, encoding="utf-8-sig") as f:
        text = f.read()
    if text.startswith("---"):
        end = text.find("\n---", 3)
        if end != -1:
            text = text[end + 4:]
    paras = []
    buf = []
    for line in text.splitlines():
        s = line.strip()
        if not s or s.startswith("|") or s == "[[PAGEBREAK]]":
            if buf:
                paras.append({"text": " ".join(buf), "kind": "p"})
                buf = []
            if s.startswith("|"):
                paras.append({"text": s, "kind": "table", "w": 0.05})
            continue
        m = re.match(r"^(#{1,3})\s+(.+)$", s)
        if m:
            if buf:
                paras.append({"text": " ".join(buf), "kind": "p"})
                buf = []
            paras.append({"text": m.group(2).strip(), "kind": f"h{len(m.group(1))}"})
            continue
        mi = re.match(r"^!\[(.+?)\]", s)
        if mi:
            if buf:
                paras.append({"text": " ".join(buf), "kind": "p"})
                buf = []
            paras.append({"text": mi.group(1), "kind": "caption"})
            continue
        if re.match(r"^(- |\d+[.)] |\[@)", s) or re.match(r"^Таблиця\s", s):
            if buf:
                paras.append({"text": " ".join(buf), "kind": "p"})
                buf = []
            paras.append({"text": s, "kind": "p"})
            continue
        buf.append(s)
    if buf:
        paras.append({"text": " ".join(buf), "kind": "p"})
    return paras


def load_docx(path):
    from docx import Document
    doc = Document(path)
    paras = []
    body = doc.element.body
    for child in body.iterchildren():
        tag = child.tag.split("}")[1]
        if tag == "p":
            from docx.text.paragraph import Paragraph
            p = Paragraph(child, doc)
            style = p.style.name if p.style is not None else ""
            kind = "h1" if style == "Heading 1" else "h2" if style == "Heading 2" else "p"
            has_img = bool(child.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip"))
            paras.append({"text": p.text.strip(), "kind": kind, "obj": p, "img": has_img})
        elif tag == "tbl":
            paras.append({"text": "", "kind": "table", "w": 0.4})
    return doc, paras


# ---------------------------------------------------------------- розбиття на частини

def sections(paras):
    """Ділить на частини за заголовками. Заголовком вважається h1 або абзац, що дорівнює назві
    структурної частини чи починається з «N. » великими літерами."""
    out = []
    cur = {"title": "_ТИТУЛКА", "paras": []}
    for p in paras:
        t = p["text"].strip().rstrip(".")
        is_head = p["kind"] == "h1" or t in STRUCT or (
            re.match(r"^\d+\.\s+[А-ЯІЇЄҐ]", t) and t == t.upper() and len(t) < 200)
        if is_head and "\t" not in p["text"]:
            out.append(cur)
            cur = {"title": t.upper(), "paras": [], "head": p}
        else:
            cur["paras"].append(p)
    out.append(cur)
    return out


# ---------------------------------------------------------------- перевірки

def check_page(doc, r):
    from docx.shared import Mm
    for i, s in enumerate(doc.sections):
        w, h = s.page_width, s.page_height
        portrait = (abs(w - Mm(210)) < Mm(2) and abs(h - Mm(297)) < Mm(2))
        landscape = (abs(w - Mm(297)) < Mm(2) and abs(h - Mm(210)) < Mm(2))
        if not (portrait or landscape):
            r.err(f"Розділ {i + 1}: формат сторінки не А4")
        elif landscape:
            r.warn(f"Розділ {i + 1}: альбомна орієнтація (допустимо лише для великих таблиць)")
        lm, rm, tm, bm = (x / Mm(1) for x in (s.left_margin, s.right_margin, s.top_margin, s.bottom_margin))
        if portrait:
            bad = []
            if lm < 24.5:
                bad.append(f"ліве {lm:.0f} мм < 25")
            if rm < 9.5:
                bad.append(f"праве {rm:.0f} мм < 10")
            if tm < 19.5:
                bad.append(f"верхнє {tm:.0f} мм < 20")
            if bm < 19.5:
                bad.append(f"нижнє {bm:.0f} мм < 20")
            if bad:
                r.err(f"Поля (розділ {i + 1}): " + ", ".join(bad))
            else:
                r.ok(f"Поля: ліве {lm:.0f}, праве {rm:.0f}, верхнє {tm:.0f}, нижнє {bm:.0f} мм")
    s = doc.sections[0]
    hdr = " ".join(p._p.xml for p in s.header.paragraphs)
    if "PAGE" not in hdr:
        r.warn("Нумерація сторінок: не знайдено поля PAGE у верхньому колонтитулі (номер праворуч угорі)")
    elif not s.different_first_page_header_footer:
        r.warn("Номер сторінки стоїть і на титулці: увімкни «Особливий колонтитул для першої сторінки»")
    else:
        r.ok("Нумерація сторінок у верхньому колонтитулі, титулка без номера")


def eff_font(run, para):
    name = run.font.name
    size = run.font.size
    st = para.style
    while st is not None and (name is None or size is None):
        name = name or st.font.name
        size = size or st.font.size
        st = st.base_style
    return name, (size.pt if size else None)


def eff_spacing(para):
    ls = para.paragraph_format.line_spacing
    st = para.style
    while ls is None and st is not None:
        ls = st.paragraph_format.line_spacing
        st = st.base_style
    return ls


def check_typography(secs, r):
    bad_font, bad_size, bad_line, colored, total = [], [], [], 0, 0
    for sec in secs:
        if sec["title"] in ("_ТИТУЛКА", "ЗМІСТ"):
            continue
        for p in sec["paras"]:
            obj = p.get("obj")
            if obj is None or not p["text"]:
                continue
            if re.match(r"^(Джерело|Примітка)", p["text"]) or p.get("img"):
                continue
            total += 1
            for run in obj.runs:
                if not run.text.strip():
                    continue
                name, size = eff_font(run, obj)
                if name and name != "Times New Roman":
                    bad_font.append(p["text"][:50])
                    break
                if size and abs(size - 14) > 0.1:
                    bad_size.append(f"{size:g} пт: {p['text'][:50]}")
                    break
                if run.font.color is not None and run.font.color.rgb not in (None,) and \
                        str(run.font.color.rgb) not in ("000000",):
                    colored += 1
                    break
            ls = eff_spacing(obj)
            if isinstance(ls, float) and abs(ls - 1.5) > 0.05:
                bad_line.append(f"{ls:g}: {p['text'][:50]}")
    if bad_font:
        r.err(f"Шрифт не Times New Roman у {len(bad_font)} абз., напр.: «{bad_font[0]}»")
    if bad_size:
        r.err(f"Кегль не 14 у {len(bad_size)} абз. основного тексту, напр.: {bad_size[0]}")
    if bad_line:
        r.warn(f"Інтервал не 1,5 у {len(bad_line)} абз., напр.: {bad_line[0]}")
    if colored:
        r.warn(f"Кольоровий текст у {colored} абз.: має бути лише чорний")
    if total and not (bad_font or bad_size or bad_line):
        r.ok(f"Шрифт TNR 14, інтервал 1,5 у всіх {total} абзацах основного тексту")


def check_structure(secs, r, is_docx=True):
    titles = [s["title"] for s in secs]
    for req in REQUIRED if is_docx else [x for x in REQUIRED if x != "ЗМІСТ"]:
        if req in titles:
            r.ok(f"Є «{req}»")
        else:
            r.err(f"Немає структурної частини «{req}»")
    if any(t in titles for t in CONCL):
        r.ok("Є «ВИСНОВКИ ТА ПРОПОЗИЦІЇ»")
    else:
        r.err("Немає «ВИСНОВКИ ТА ПРОПОЗИЦІЇ»")
    if "ДОДАТКИ" not in titles:
        r.warn("Немає «ДОДАТКИ» (ЕФП-2024 вимагає фінзвітність підприємства в додатках)")
    punkty = [t for t in titles if re.match(r"^\d+\.\s", t)]
    nums = [int(re.match(r"^(\d+)", t).group(1)) for t in punkty]
    if not punkty:
        r.err("Не знайдено пунктів основної частини «1. НАЗВА»")
    else:
        if nums != list(range(1, len(nums) + 1)):
            r.err(f"Нумерація пунктів збита: {nums}")
        if not 3 <= len(punkty) <= 5:
            r.warn(f"Пунктів {len(punkty)}: простий план ДТЕУ має 3-5")
        else:
            r.ok(f"Пунктів основної частини: {len(punkty)}")
    for s in secs:
        h = s.get("head")
        if h and h["text"].rstrip() != h["text"].rstrip().rstrip("."):
            r.warn(f"Крапка в кінці заголовка: «{h['text'][:60]}»")
        if h and s["title"] in STRUCT and h["text"].strip() != h["text"].strip().upper():
            r.warn(f"Заголовок не великими літерами: «{h['text'][:60]}»")


def section_text(sec):
    return " ".join(p["text"] for p in sec["paras"] if p["kind"] != "table")


def check_text(secs, r, is_docx=True):
    main = [s for s in secs if s["title"] not in ("_ТИТУЛКА", "ЗМІСТ", "СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ", "ДОДАТКИ")
            and not s["title"].startswith("СПИСОК")]
    text = " ".join(section_text(s) for s in main)
    allt = " ".join(section_text(s) for s in secs if s["title"] not in ("_ТИТУЛКА", "ЗМІСТ"))
    em, en = allt.count("—"), allt.count("–")
    if em or en:
        ex = re.search(r".{0,30}[—–].{0,30}", allt)
        r.err(f"Довгі/середні тире: — {em}, – {en} (правило: без тире). Напр.: «{ex.group(0) if ex else ''}»")
    else:
        r.ok("Довгих і середніх тире немає")
    fp = [m.group(0) for m in FIRST_PERSON.finditer(text)]
    if fp:
        ctx = FIRST_PERSON.search(text)
        r.warn(f"Перша особа однини ({len(fp)}): «{text[max(0, ctx.start() - 30):ctx.end() + 30]}»")
    else:
        r.ok("Перша особа однини не вживається")
    low = text.lower()
    found = [(c, low.count(c)) for c in CLICHES if c in low]
    if found:
        r.warn("Кліше: " + ", ".join(f"«{c}» ×{n}" for c, n in found))
    vstup = next((s for s in secs if s["title"] == "ВСТУП"), None)
    if vstup:
        v = section_text(vstup).lower().replace("’", "'").replace("ʼ", "'")
        need = {"актуальність": "актуальн", "мета": "мет", "завдання": "завдан", "об'єкт": "об'єкт",
                "предмет": "предмет", "методи": "метод", "інформаційна база": "інформаційн",
                "практичне значення": "практичн", "база дослідження (підприємство)": "баз"}
        miss = [k for k, stem in need.items() if stem not in v]
        if miss:
            r.warn("У вступі не знайдено: " + ", ".join(miss))
        else:
            r.ok("Вступ містить усі обов'язкові елементи (актуальність, мета, завдання, об'єкт, предмет, методи, база, практичне значення)")
        if (is_docx and "{{" in v) or re.search(r"\bnn\b", v):
            r.err("У вступі лишилися заглушки ({{...}} або NN)")
    for s in secs:
        if re.match(r"^\d+\.\s", s["title"]):
            ps = [p for p in s["paras"] if p["kind"] == "p" and p["text"] and not p["text"].startswith(("Джерело", "Примітка"))]
            if ps and not CONCL_START.match(ps[-1]["text"]):
                r.warn(f"Пункт «{s['title'][:45]}» не закінчується висновком («Отже, ...», 2-3 речення)")
    return text


CITE = re.compile(r"\[([^\[\]]*\d[^\[\]]*)\]")


def parse_cites(text):
    nums, seps = [], set()
    for m in CITE.finditer(text):
        inner = m.group(1)
        if re.search(r"[A-Za-zА-Яа-яІіЇїЄє]{3,}", re.sub(r"\bс\.|\bp\.|\bpp\.", "", inner)):
            continue
        for part in re.split(r";", inner):
            part = re.sub(r",\s*(с|p|pp)\.\s*[\d\-–]+", "", part.strip())
            if "," in part:
                seps.add(",")
            if ";" in inner:
                seps.add(";")
            for n in re.findall(r"\d+", part):
                nums.append(int(n))
    return nums, seps


def check_refs(secs, body_text, r, a):
    sec = next((s for s in secs if s["title"].startswith("СПИСОК")), None)
    if not sec:
        return []
    entries, keys = [], []
    for p in sec["paras"]:
        t = p["text"].strip()
        if not t:
            continue
        t = re.sub(r"^\d+[.)]\s*", "", t)
        mk = re.match(r"^\[@([^\]]+)\]\s*", t)
        if mk:
            keys.append(mk.group(1).strip())
            t = t[mk.end():]
        entries.append(t)
    n = len(entries)
    order = "alpha"
    if keys:
        # чернетка з ключами [@ключ]: номери і порядок призначить build_docx.py
        cited = set(k.strip().lstrip("@").split(",")[0].strip()
                    for m in re.finditer(r"\[(@[^\]]+)\]", body_text) for k in m.group(1).split(";"))
        unknown = sorted(cited - set(keys))
        unused = [k for k in keys if k not in cited]
        if unknown:
            r.err(f"Посилання на ключі, яких немає в списку: {unknown}")
        if unused:
            r.err(f"Джерела без посилань у тексті: {unused}")
        if not unknown and not unused:
            r.ok("Кожне джерело зі списку згадано в тексті (ключі [@...])")
        a = argparse.Namespace(**{**vars(a), "apa": True, "_keys": True})
    if n < a.min_src:
        r.err(f"Джерел {n}: менше норми {a.min_src}")
    elif n > a.max_src:
        r.warn(f"Джерел {n}: більше {a.max_src}, перевір, чи всі використані")
    else:
        r.ok(f"Джерел у списку: {n}")

    if not a.apa:
        nums, seps = parse_cites(body_text)
        if not nums:
            r.err("У тексті немає посилань [n] на джерела")
        else:
            over = sorted({x for x in nums if x > n or x < 1})
            if over:
                r.err(f"Посилання на неіснуючі номери: {over}")
            unused = [i for i in range(1, n + 1) if i not in set(nums)]
            if unused:
                r.err(f"Джерела без посилань у тексті: {unused}")
            else:
                r.ok("Кожне джерело зі списку згадано в тексті")
            if len(seps) > 1:
                r.warn("Роздільник у посиланнях змішаний: і «;», і «,» між номерами")
            first = []
            for x in nums:
                if x not in first:
                    first.append(x)
            if first[:min(10, n)] == list(range(1, min(10, n) + 1)):
                order = "mention"

    year_now = dt.date.today().year
    old, problems = 0, []
    for i, e in enumerate(entries, 1):
        years = [int(y) for y in re.findall(r"\b(19[5-9]\d|20[0-4]\d)\b", e) if int(y) <= year_now]
        is_law = re.search(r"Закон|Постанова|Кодекс|Указ|Наказ|ДСТУ|Конституц|Розпорядження", e)
        if not years:
            problems.append((i, "немає року"))
        elif not is_law and max(years) < year_now - 5:
            old += 1
        if re.search(r"https?://|www\.", e) and not re.search(r"дата звернення|accessed", e, re.I):
            problems.append((i, "URL без «(дата звернення: дд.мм.рррр)»"))
        if re.search(r"\.\s[–—-]\s", e):
            problems.append((i, "старе тире-роздільник «. -» (ГОСТ 7.1), за ДСТУ 8302 лише крапка"))
        if re.search(r"(^|\s)К\.\s*:", e) or re.search(r"(^|\s)Х\.\s*:", e):
            problems.append((i, "скорочене місто «К.»/«Х.», пиши повністю"))
        if re.search(rf"({CITIES})\s?:(?=\S)", e) or re.search(rf"({CITIES}):", e):
            problems.append((i, "двокрапка перед видавництвом без пробілів («Київ : Видавництво»)"))
        if re.match(r"^[А-ЯІЇЄҐA-Z]\.\s?[А-ЯІЇЄҐA-Z]\.\s+[А-ЯІЇЄҐA-Z][а-яіїєґa-z'’]+", e):
            problems.append((i, "ініціали перед прізвищем на початку запису"))
        if re.search(rf"({CITIES})\s:\s", e) and not re.search(r"\b\d+\s*(с|p)\.", e) and not re.search(r"\b(С|P)\.\s*\d", e):
            problems.append((i, "книга без кількості сторінок («256 с.»)"))
        if BAD_SOURCES.search(e):
            problems.append((i, "неакадемічне джерело (вікі, реферати)"))
        if "[ПЕРЕВІРИТИ]" in e or "ПЕРЕВІРИТИ" in e:
            problems.append((i, "джерело не перевірене"))
    if problems:
        for i, msg in problems[:25]:
            r.err(f"Джерело {i}: {msg}  «{entries[i - 1][:70]}»")
        if len(problems) > 25:
            r.err(f"... ще {len(problems) - 25} проблем у списку")
    else:
        r.ok("Записи списку пройшли перевірку формату ДСТУ 8302")
    non_law = sum(1 for e in entries if not re.search(r"Закон|Постанова|Кодекс|Указ|ДСТУ", e))
    if non_law and old / non_law > 1 / 3:
        r.warn(f"Застарілі джерела: {old} з {non_law} старші за 5 років (бажано не більше третини)")
    else:
        r.ok(f"Свіжість: {old} джерел старші за 5 років")
    latin = sum(1 for e in entries if re.match(r"^[A-Za-z]", e))
    if latin == 0:
        r.warn("Немає іноземних джерел (комісія оцінює «вітчизняні та зарубіжні»)")
    if not any(re.search(r"Закон|Постанова|Кодекс", e) for e in entries):
        r.warn("Немає законів і нормативних актів (ЕФП-2024: висвітлення законодавчої бази обов'язкове)")
    if order == "alpha" and not a.apa and not getattr(a, "_keys", False):
        def key(s):
            s = re.sub(r"^[«\"'(\[]+", "", s).upper()
            return [(0, UA_ORDER.index(c)) if c in UA_ORDER else (1, ord(c)) if "A" <= c <= "Z"
                    else (0, 50) if c.isalpha() else (-1, 0) for c in s]
        for i in range(1, n):
            if key(entries[i]) < key(entries[i - 1]):
                r.warn(f"Порушено алфавітний порядок: {i} «{entries[i - 1][:30]}» → {i + 1} «{entries[i][:30]}» "
                       f"(або список за порядком згадки, але тоді перші посилання мають іти 1, 2, 3...)")
                break
    return entries


def check_tables_figs(paras, body_text, r):
    tabs = []
    figs = []
    for idx, p in enumerate(paras):
        t = p["text"]
        mt = re.match(r"^Таблиця\s+([А-Я]?\.?\d+(?:\.\d+)?)\s*[:.]?\s*(.*)$", t)
        if mt and len(t) < 250:
            tabs.append(mt.group(1))
        mf = re.match(r"^Рис\.\s*([А-Я]?\.?\d+(?:\.\d+)?)\.?\s+\S", t)
        if mf and (p["kind"] == "caption" or (idx > 0 and paras[idx - 1].get("img")) or
                   (p.get("obj") is not None and p["obj"].alignment == 1)):
            figs.append(mf.group(1))
    low = body_text
    miss_t = [n for n in tabs if not re.search(rf"табл\.\s*{re.escape(n)}(?![\d])", low)]
    miss_f = [n for n in figs if not re.search(rf"рис\.\s*{re.escape(n)}(?![\d])", low)]
    if tabs:
        if miss_t:
            r.err(f"Таблиці без посилання в тексті «(табл. N)»: {miss_t}")
        else:
            r.ok(f"Таблиць {len(tabs)}, усі згадані в тексті")
    else:
        r.warn("Немає жодної таблиці (аналітична частина зазвичай має 5+ таблиць)")
    if figs:
        if miss_f:
            r.err(f"Рисунки без посилання в тексті «(рис. N)»: {miss_f}")
        else:
            r.ok(f"Рисунків {len(figs)}, усі згадані в тексті")
    else:
        r.warn("Немає рисунків з підписом «Рис. N. Назва» (схеми, діаграми додають бали за наочність)")
    for label, items in (("таблиць", tabs), ("рисунків", figs)):
        main = [x for x in items if not re.match(r"^[А-Я]", x)]
        if main and all("." not in x for x in main):
            nums = [int(x) for x in main]
            if nums != list(range(1, len(nums) + 1)):
                r.warn(f"Наскрізна нумерація {label} збита: {main}")
    return tabs, figs


def check_volume(secs, r, a, docx_path):
    main = [s for s in secs if s["title"] == "ВСТУП" or re.match(r"^\d+\.\s", s["title"]) or s["title"] in CONCL]
    chars = sum(len(section_text(s)) for s in main)
    est = chars / 1800 + sum(p.get("w", 0) for s in main for p in s["paras"] if p["kind"] == "table")
    pages = None
    if a.word and docx_path:
        res = word_com.run_word(docx_path, update=False)
        if res and res.get("ok"):
            pages = word_com.main_text_pages(res.get("headings"))
            if pages:
                r.ok(f"Word: усього {res.get('pages')} стор., основний текст (вступ...висновки) {pages} стор.")
        else:
            r.warn("Word недоступний, обсяг оцінено приблизно")
    v = pages or est
    label = f"{pages} стор." if pages else f"~{est:.0f} стор. (оцінка по знаках)"
    if v < a.min_pages - 0.5:
        r.err(f"Обсяг основного тексту {label} < {a.min_pages}")
    elif v > a.max_pages + 0.5:
        r.warn(f"Обсяг основного тексту {label} > {a.max_pages}")
    else:
        r.ok(f"Обсяг основного тексту {label} у межах {a.min_pages}-{a.max_pages}")
    vstup = next((s for s in secs if s["title"] == "ВСТУП"), None)
    if vstup and len(section_text(vstup)) / 1800 > 3.3:
        r.warn(f"Вступ ~{len(section_text(vstup)) / 1800:.1f} стор.: норма 2-3")
    concl = next((s for s in secs if s["title"] in CONCL), None)
    if concl and not 1.5 <= len(section_text(concl)) / 1800 <= 3.5:
        r.warn(f"Висновки ~{len(section_text(concl)) / 1800:.1f} стор.: норма 2-3")
    p1 = next((s for s in secs if s["title"].startswith("1. ")), None)
    body = sum(len(section_text(s)) for s in secs if re.match(r"^\d+\.\s", s["title"]))
    if p1 and body and len(section_text(p1)) / body > 0.3:
        r.warn(f"Теоретичний пункт {len(section_text(p1)) / body:.0%} основної частини: норма до 25 %")


def check_online(entries, r):
    urls = []
    for i, e in enumerate(entries, 1):
        for u in re.findall(r"https?://[^\s)]+", e):
            urls.append((i, u.rstrip(".,;")))
        for d in re.findall(r"\b10\.\d{4,9}/[^\s)]+", e):
            if "doi.org" not in e:
                urls.append((i, "https://doi.org/" + d.rstrip(".,;")))
    bad = 0
    for i, u in urls:
        u = urllib.parse.quote(u, safe=":/?&=#%+@,;~")
        req = urllib.request.Request(u, headers={"User-Agent": "Mozilla/5.0 (kursova-check)"})
        try:
            with urllib.request.urlopen(req, timeout=12) as resp:
                code = resp.status
        except urllib.error.HTTPError as ex:
            code = ex.code
        except Exception as ex:  # noqa: BLE001
            code = type(ex).__name__
        if code not in (200, 301, 302, 403):
            bad += 1
            r.err(f"Джерело {i}: {u} не відповідає ({code})")
        elif code == 403:
            r.warn(f"Джерело {i}: {u} відповідає 403 (сайт блокує ботів), відкрий вручну")
    if urls and not bad:
        r.ok(f"Онлайн: усі {len(urls)} URL/DOI відповідають")


def main():
    ap = argparse.ArgumentParser(description="Перевірка курсової ДТЕУ")
    ap.add_argument("path")
    ap.add_argument("--online", action="store_true")
    ap.add_argument("--word", action="store_true")
    ap.add_argument("--apa", action="store_true")
    ap.add_argument("--min-pages", type=int, default=25)
    ap.add_argument("--max-pages", type=int, default=35)
    ap.add_argument("--min-src", type=int, default=20)
    ap.add_argument("--max-src", type=int, default=40)
    a = ap.parse_args()

    r = Report()
    is_docx = a.path.lower().endswith(".docx")
    if is_docx:
        doc, paras = load_docx(a.path)
        check_page(doc, r)
    else:
        paras = load_md(a.path)
    secs = sections(paras)
    if is_docx:
        check_typography(secs, r)
    check_structure(secs, r, is_docx)
    body_text = check_text(secs, r, is_docx)
    entries = check_refs(secs, body_text, r, a)
    check_tables_figs(paras, body_text, r)
    check_volume(secs, r, a, a.path if is_docx else None)
    if a.online and entries:
        check_online(entries, r)
    print(f"Перевірка: {os.path.abspath(a.path)}\n")
    sys.exit(1 if r.print() else 0)


if __name__ == "__main__":
    main()
