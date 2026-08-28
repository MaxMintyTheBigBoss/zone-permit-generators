# -*- coding: utf-8 -*-
"""
Модуль генерации документов пропусков по п. 14.5 (вывоз имущества).
Заявление 14.5 + пропуск на вывоз имущества + транспортный пропуск.
Работает полностью локально, без интернета и LLM.
"""
import os
import re
import sys
from datetime import datetime
from copy import deepcopy

from docx import Document
from docx.oxml.ns import qn

# Цель въезда (Placeholder_6) — по условию всегда "для вывоза имущества"
GOAL_145 = "для вывоза имущества"

# Районы (Placeholder_4) — те же 8, что и для 14.3
DISTRICTS = [
    "Брагинский", "Буда-Кошелевский", "Ветковский", "Добрушский",
    "Кормянский", "Наровлянский", "Хойникский", "Чечерский",
]

# Кому на подписание (Placeholder_20) — полный список
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

# Регулярка для плейсхолдеров
_PH_RE = re.compile(r"\*{0,2}(Placeholder_\d+(?:\.\d+)?)\*{0,2}")


def resource_path():
    """Путь к ресурсам: совместим с PyInstaller."""
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return base
    return os.path.dirname(os.path.abspath(__file__))


def runtime_base():
    """Постоянное место рядом с программой (для данных и вывода)."""
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
    Заменённый текст делается подчёркнутым, как в бланках."""
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
        # подчёркивание для вставленного текста
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


def _today_dmy():
    return datetime.now().strftime("%d.%m.%Y")


def make_application_145(data):
    """Заявление 14.5 (вывоз имущества)."""
    doc = Document(os.path.join(template_dir(), "Заявление_14.5.docx"))
    districts = data.get("districts", [])
    m = {
        "Placeholder_1.1": data["last_name"],
        "Placeholder_1.2": data["first_name"],
        "Placeholder_1.3": data["middle_name"],
        "Placeholder_2": data.get("birth_date", ""),
        "Placeholder_3": data.get("id_number", ""),
        "Placeholder_4": join_districts(districts),
        "Placeholder_5": data.get("objects", ""),
        "Placeholder_7": data.get("date_from", ""),
        "Placeholder_8": data.get("date_to", ""),
        "Placeholder_12": data.get("car_make", ""),
        "Placeholder_13": data.get("car_number", ""),
        "Placeholder_21": data.get("cargo", ""),  # вид и количество имущества
    }
    # распределение районов по ячейкам Placeholder_4.1/4.2/4.3 (если есть в шаблоне)
    for i in range(3):
        m["Placeholder_4.%d" % (i + 1)] = districts[i] if i < len(districts) else ""

    fill_doc(doc, m)
    return doc


def make_permit_cargo(data):
    """Пропуск на вывоз имущества."""
    doc = Document(os.path.join(template_dir(), "Пропуск_вывоз_имущества.docx"))
    districts = data.get("districts", [])
    m = {
        "Placeholder_1.1": data["last_name"],
        "Placeholder_1.2": data["first_name"],
        "Placeholder_1.3": data["middle_name"],
        "Placeholder_4": join_districts(districts),
        "Placeholder_5": data.get("objects", ""),
        "Placeholder_7": data.get("date_from", ""),
        "Placeholder_8": data.get("date_to", ""),
        "Placeholder_20": data.get("issued_by", ""),
        "Placeholder_21": data.get("cargo", ""),  # вид и количество имущества
    }
    fill_doc(doc, m)
    return doc


def make_transport_145(data):
    """Транспортный пропуск для 14.5 (Placeholder_6 всегда "для вывоза имущества")."""
    doc = Document(os.path.join(template_dir(), "Пропуск_транспортный.docx"))
    districts = data.get("districts", [])
    m = {
        "Placeholder_1.1": data["last_name"],
        "Placeholder_1.2": data["first_name"],
        "Placeholder_1.3": data["middle_name"],
        "Placeholder_4": join_districts(districts),
        "Placeholder_5": data.get("objects", ""),
        "Placeholder_6": GOAL_145,  # фиксированная цель
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


def generate_all_145(data, output_dir=None):
    """Полный комплект документов по п. 14.5. Возвращает список файлов."""
    if output_dir is None:
        output_dir = os.path.join(runtime_base(), "output")
    stamp = datetime.now().strftime("%d.%m.%Y.%H.%M")
    out = os.path.join(output_dir, stamp)
    os.makedirs(out, exist_ok=True)

    created = []
    lname = sanitize(data["last_name"], "Заявитель")

    # 1. Заявление 14.5
    d = make_application_145(data)
    p = os.path.join(out, "Заявление_14.5_%s.docx" % lname)
    d.save(p)
    created.append(p)

    # 2. Пропуск на вывоз имущества
    d = make_permit_cargo(data)
    p = os.path.join(out, "Пропуск_ВывозИмущества_%s.docx" % lname)
    d.save(p)
    created.append(p)

    # 3. Транспортный пропуск (если заполнено авто)
    if data.get("car_make", "").strip() or data.get("car_number", "").strip():
        car = sanitize(data.get("car_number", "") or "Авто", "Авто")
        d = make_transport_145(data)
        p = os.path.join(out, "Пропуск_Транспорт_%s.docx" % car)
        d.save(p)
        created.append(p)

    # запись в БД истории
    _db_remember_145(data)

    return created


# ---------------------------------------------------------------------------
# Лёгкая «База знаний» для автоподстановки
# ---------------------------------------------------------------------------
def _db_path_145():
    d = os.path.join(runtime_base(), "data")
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, "permit_history_145.json")


def _db_load_145():
    import json
    try:
        with open(_db_path_145(), "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def db_find_145(fio):
    """Ищет прошлую запись по фамилии-имени-отчеству. Возвращает dict или None."""
    try:
        key = (fio or "").strip().lower()
        if not key:
            return None
        for rec in _db_load_145():
            rk = " ".join(x for x in [rec.get("last_name", ""), rec.get("first_name", ""), rec.get("middle_name", "")] if x).strip().lower()
            if rk == key:
                return rec
        return None
    except Exception:
        return None


def _db_remember_145(data):
    try:
        import json
        rec = {
            "last_name": data["last_name"], "first_name": data["first_name"],
            "middle_name": data["middle_name"], "birth_date": data.get("birth_date", ""),
            "id_number": data.get("id_number", ""), "cargo": data.get("cargo", ""),
            "districts": data.get("districts", []),
            "objects": data.get("objects", ""),
            "car_make": data.get("car_make", ""), "car_number": data.get("car_number", ""),
        }
        recs = [r for r in _db_load_145() if not (
            r.get("last_name") == rec["last_name"]
            and r.get("first_name") == rec["first_name"]
            and r.get("middle_name") == rec["middle_name"])]
        recs.insert(0, rec)
        with open(_db_path_145(), "w", encoding="utf-8") as f:
            json.dump(recs[:200], f, ensure_ascii=False, indent=1)
    except Exception:
        pass