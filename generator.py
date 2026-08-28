# -*- coding: utf-8 -*-
"""
Модуль генерации документов пропусков по промту (п. 14.3).
Заявление 14.3 + индивидуальные и транспортные пропуска.
Работает полностью локально, без интернета и LLM.
"""
import os
import re
import sys
from datetime import datetime
from copy import deepcopy

from docx import Document
from docx.oxml.ns import qn

# Цель въезда (Placeholder_6) — по промту: кнопка или «свой вариант».
GOAL_BLAGOUSTROISTVO = "благоустройство места захоронения"
GOAL_CUSTOM = "свой вариант"

# Районы (Placeholder_4) — по промту, множественный выбор из 8.
DISTRICTS = [
    "Брагинский", "Буда-Кошелевский", "Ветковский", "Добрушский",
    "Кормянский", "Наровлянский", "Хойникский", "Чечерский",
]

# Кому на подписание (Placeholder_20) — полный список.
SIGNERS = [
    "Заместитель начальника отдела Путькова Т.М.",
    "Начальник отдела Соломейчук А.В.",
    "Заместитель начальника главного управления В.О.Шабловский",
    "Главный специалист Гвоздарев А.А.",
    "Главный специалист Геращенко Г.Н.",
    "Главный специалист Колесан А.И.",
    "Главный специалист Курило А.В.",
    "Главный специалист Новик П.Н.",
    "Главный специалист Одиноченко И.В.",
    "Главный специалист Першко А.С.",
]

# Регулярка для плейсхолдеров: Placeholder_1.1, Placeholder_4.2, **Placeholder_20**
_PH_RE = re.compile(r"\*{0,2}(Placeholder_\d+(?:\.\d+)?)\*{0,2}")


def resource_path():
    """Путь к ресурсам: совместим с PyInstaller (из .exe берём из _MEIPASS)."""
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return base
    return os.path.dirname(os.path.abspath(__file__))


def runtime_base():
    """Постоянное место рядом с программой (для данных и вывода).

    В .exe это папка самого exe, чтобы база знаний и результат сохранялись
    между запусками (в отличие от временного _MEIPASS)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.argv[0]))
    return os.path.dirname(os.path.abspath(__file__))


def template_dir():
    return os.path.join(resource_path(), "templates")


def load_reference():
    """Справочник: список {district, object}. Объект = 'кладбище о.н.п. <название>'."""
    path = os.path.join(template_dir(), "справочник.docx")
    entries = []
    if not os.path.exists(path):
        return entries
    doc = Document(path)
    for p in doc.paragraphs:
        parts = [t for t in p.text.split("\t") if t.strip()]
        if len(parts) >= 3:
            district = parts[0].strip()
            obj = "кладбище о.н.п. " + parts[2].strip()
            if district and obj:
                entries.append({"district": district, "object": obj})
    return entries


def filter_objects_by_districts(reference, districts):
    """Только кладбища, относящиеся к выбранным районам."""
    allowed = set(districts)
    return [r for r in reference if r["district"] in allowed]


def join_districts(districts):
    return ", ".join(districts)


def join_objects(objects, custom=""):
    lst = [o.get("object", "") for o in objects if o.get("object")]
    if custom and custom.strip():
        lst.append(custom.strip())
    return "; ".join(lst)


def _inline_replace(paragraph, mapping):
    """Заменяет плейсхолдеры в абзаце regex-ом (учитывая разбивку по runs).
    Все заменённые значения (плейсхолдеры) — подчёркнутые."""
    full = "".join(r.text for r in paragraph.runs)
    if "Placeholder_" not in full:
        return

    def repl(m):
        tok = m.group(1)
        return str(mapping.get(tok, m.group(0)))

    new = _PH_RE.sub(repl, full)
    if new == full:
        return
    if paragraph.runs:
        paragraph.runs[0].text = new
        for r in paragraph.runs:
            r.font.underline = True
        for r in paragraph.runs[1:]:
            r.text = ""


def fill_doc(doc, mapping):
    """Заменяет плейсхолдеры во всех абзацах и таблицах документа."""
    for para in doc.paragraphs:
        _inline_replace(para, mapping)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _inline_replace(para, mapping)
    return doc


def _iter_all_paragraphs(doc):
    """Все абзацы документа, включая ячейки таблиц."""
    for para in doc.paragraphs:
        yield para
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    yield para


def _set_para_text(p, text):
    """Записывает текст в абзац, сохраняя форматирование первого run."""
    runs = p.findall(qn('w:r'))
    if not runs:
        return
    r0 = runs[0]
    for t in r0.findall(qn('w:t')):
        r0.remove(t)
    t = r0.makeelement(qn('w:t'), {})
    t.text = str(text)
    r0.append(t)
    for r in runs[1:]:
        for t in r.findall(qn('w:t')):
            t.text = ""


def _build_persons_block(doc, persons):
    """Перестраивает блок «для меня лично и следующего(их) со мной».

    Заявитель в этом месте НЕ показывается (пропуск выдается отдельно).
    Здесь перечисляются только сопровождающие, каждый по маске
    «Фамилия Имя Отчество, дата рождения» — одной строкой, без абзацев.
    Строка-шаблон (номер + данные) клонируется на каждого сопровождающего.
    """
    W = qn('w:p')
    WTR = qn('w:tr')
    WTC = qn('w:tc')
    TBL = qn('w:tbl')
    anchor_text = "для граждан Республики Беларусь"

    # находим таблицу заявления (плейсхолдеры разбиты на несколько w:t — склеиваем)
    tbl = None
    for t in doc.element.body.iter(TBL):
        full = "".join(x.text or "" for x in t.iter(qn('w:t')))
        if "Placeholder_10.1" in full:
            tbl = t
            break
    if tbl is None:
        return
    rows = [tr for tr in tbl if tr.tag == WTR]

    # строка-шаблон (с данными Placeholder_10.x) и якорная строка
    tpl_row = None
    anchor_row = None
    for tr in rows:
        txt = "".join(c.text for c in tr.iter(qn('w:t')))
        if tpl_row is None and "Placeholder_10.1" in txt:
            tpl_row = tr
        if anchor_row is None and anchor_text in txt:
            anchor_row = tr
        if tpl_row is not None and anchor_row is not None:
            break
    if tpl_row is None or anchor_row is None:
        return

    # клоны-шаблоны (до удаления из дерева)
    num_tpl = deepcopy(tpl_row)
    data_tpl = deepcopy(tpl_row)

    # удаляем исходную строку-шаблон и следующую за ней «пустую» строку-«2»
    try:
        i = rows.index(tpl_row)
        for tr in rows[i:i + 2]:
            if tr is not None and tr in tbl:
                tbl.remove(tr)
    except Exception:
        pass

    # вставляем строки сопровождающих перед якорной строкой
    for i, pers in enumerate(persons):
        name = " ".join(x for x in [pers.get("last_name", ""), pers.get("first_name", ""), pers.get("middle_name", "")] if x)
        birth = pers.get("birth_date", "")
        line = (name + (", " if name and birth else "") + birth).strip()
        new_row = deepcopy(data_tpl)
        cells = [c for c in new_row if c.tag == WTC]
        if len(cells) >= 2:
            k1 = [p for p in cells[1] if p.tag == W]
            if k1:
                _set_para_text(k1[0], line)
                for pp in k1[1:]:
                    for rr in pp.findall(qn('w:r')):
                        for tt in rr.findall(qn('w:t')):
                            tt.text = ""
            k0 = [p for p in cells[0] if p.tag == W]
            if k0:
                _set_para_text(k0[0], str(i + 1))
                for pp in k0[1:]:
                    for rr in pp.findall(qn('w:r')):
                        for tt in rr.findall(qn('w:t')):
                            tt.text = ""
        anchor_row.addprevious(new_row)


def _today_dmy():
    return datetime.now().strftime("%d.%m.%Y")


def make_application(data):
    """Заявление 14.3."""
    doc = Document(os.path.join(template_dir(), "Заявление_14.3.docx"))
    districts = data.get("districts", [])
    m = {
        "Placeholder_1.1": data["last_name"],
        "Placeholder_1.2": data["first_name"],
        "Placeholder_1.3": data["middle_name"],
        "Placeholder_2": data.get("birth_date", ""),
        "Placeholder_3": data.get("id_number", ""),
        "Placeholder_4": join_districts(districts),
        "Placeholder_5": data.get("objects", ""),
        "Placeholder_6": data.get("goal", ""),
        "Placeholder_7": data.get("date_from", ""),
        "Placeholder_8": data.get("date_to", ""),
        "Placeholder_12": data.get("car_make", ""),
        "Placeholder_13": data.get("car_number", ""),
        "Placeholder_9": data.get("app_date") or _today_dmy(),
    }
    # распределение районов по ячейкам Placeholder_4.1/4.2/4.3 (если есть в шаблоне)
    for i in range(3):
        m["Placeholder_4.%d" % (i + 1)] = districts[i] if i < len(districts) else ""

    fill_doc(doc, m)

    # блок людей: только сопровождающие (заявитель здесь не показывается)
    _build_persons_block(doc, data.get("persons", []))
    return doc


def make_individual(data, person=None):
    """Индивидуальный пропуск. person = None => заявитель, иначе пассажир."""
    person = person if person is not None else data
    doc = Document(os.path.join(template_dir(), "Пропуск_индивидуальный.docx"))
    m = {
        "Placeholder_1.1": person["last_name"],
        "Placeholder_1.2": person["first_name"],
        "Placeholder_1.3": person["middle_name"],
        "Placeholder_4": join_districts(data.get("districts", [])),
        "Placeholder_5": data.get("objects", ""),
        "Placeholder_6": data.get("goal", ""),
        "Placeholder_7": data.get("date_from", ""),
        "Placeholder_8": data.get("date_to", ""),
        "Placeholder_20": data.get("issued_by", ""),
    }
    fill_doc(doc, m)
    return doc


def make_transport(data):
    """Транспортный пропуск (заинтересованное лицо — заявитель)."""
    doc = Document(os.path.join(template_dir(), "Пропуск_транспортный.docx"))
    m = {
        "Placeholder_1.1": data["last_name"],
        "Placeholder_1.2": data["first_name"],
        "Placeholder_1.3": data["middle_name"],
        "Placeholder_4": join_districts(data.get("districts", [])),
        "Placeholder_5": data.get("objects", ""),
        "Placeholder_6": data.get("goal", ""),
        "Placeholder_7": data.get("date_from", ""),
        "Placeholder_8": data.get("date_to", ""),
        "Placeholder_12": data.get("car_make", ""),
        "Placeholder_13": data.get("car_number", ""),
        "Placeholder_20": data.get("issued_by", ""),
    }
    fill_doc(doc, m)
    return doc


def sanitize(s, default="файл"):
    illegal = '<>:"/\\|?*'
    return ("".join(c for c in s if c not in illegal).strip() or default)


def generate_all(data, output_dir=None):
    """Полный комплект документов по промту. Возвращает список файлов.

    Создаёт подпапку ДД.ММ.ГГГГ.ЧЧ.ММ и имена файлов по промту."""
    if output_dir is None:
        output_dir = os.path.join(runtime_base(), "output")
    stamp = datetime.now().strftime("%d.%m.%Y.%H.%M")
    out = os.path.join(output_dir, stamp)
    os.makedirs(out, exist_ok=True)

    created = []
    lname = sanitize(data["last_name"], "Заявитель")

    # 1. Заявление
    d = make_application(data)
    p = os.path.join(out, "Заявление_%s.docx" % lname)
    d.save(p)
    created.append(p)

    # 2. Индивидуальный пропуск заявителя
    d = make_individual(data)
    p = os.path.join(out, "Пропуск_Заявитель_%s.docx" % lname)
    d.save(p)
    created.append(p)

    # 3. Индивидуальные пропуска пассажиров (отдельный файл на каждого)
    for person in data.get("persons", []):
        plname = sanitize(person["last_name"], "Пассажир")
        d = make_individual(data, person)
        p = os.path.join(out, "Пропуск_Пассажир_%s.docx" % plname)
        d.save(p)
        created.append(p)

    # 4. Транспортный пропуск
    if data.get("car_make", "").strip() or data.get("car_number", "").strip():
        car = sanitize(data.get("car_number", "") or "Авто", "Авто")
        d = make_transport(data)
        p = os.path.join(out, "Пропуск_Транспорт_%s.docx" % car)
        d.save(p)
        created.append(p)

    # запись в БД истории (для авто-подстановки в будущих сессиях)
    _db_remember(data)

    return created


# ----------------------------------------------------------------------------
# Лёгкая «База знаний» для авто-подстановки (упрощённый вариант памяти из промта)
# ----------------------------------------------------------------------------
def _db_path():
    d = os.path.join(runtime_base(), "data")
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, "permit_history.json")


def _db_load():
    import json
    try:
        with open(_db_path(), "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def db_find(fio):
    """Ищет прошлую запись по фамилии-имени-отчеству. Возвращает dict или None."""
    try:
        key = (fio or "").strip().lower()
        if not key:
            return None
        for rec in _db_load():
            rk = " ".join(x for x in [rec.get("last_name", ""), rec.get("first_name", ""), rec.get("middle_name", "")] if x).strip().lower()
            if rk == key:
                return rec
        return None
    except Exception:
        return None


def _db_remember(data):
    try:
        import json
        rec = {
            "last_name": data["last_name"], "first_name": data["first_name"],
            "middle_name": data["middle_name"], "birth_date": data.get("birth_date", ""),
            "id_number": data.get("id_number", ""), "goal": data.get("goal", ""),
            "districts": data.get("districts", []),
            "objects": data.get("objects", ""),
            "car_make": data.get("car_make", ""), "car_number": data.get("car_number", ""),
        }
        recs = [r for r in _db_load() if not (
            r.get("last_name") == rec["last_name"]
            and r.get("first_name") == rec["first_name"]
            and r.get("middle_name") == rec["middle_name"])]
        recs.insert(0, rec)
        with open(_db_path(), "w", encoding="utf-8") as f:
            json.dump(recs[:200], f, ensure_ascii=False, indent=1)
    except Exception:
        pass