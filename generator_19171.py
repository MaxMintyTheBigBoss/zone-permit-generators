# -*- coding: utf-8 -*-
"""
Модуль генерации документов пропусков по п. 19.17.1.
Индивидуальные пропуска + транспортные пропуска (без заявления).
Работает полностью локально, без интернета и LLM.
"""
import os
import re
import sys
from datetime import datetime
from copy import deepcopy

from docx import Document
from docx.oxml.ns import qn

# Районы (Placeholder_4) — те же 8
DISTRICTS = [
    "Брагинский", "Буда-Кошелевский", "Ветковский", "Добрушский",
    "Кормянский", "Наровлянский", "Хойникский", "Чечерский",
]

# Подписанты (Placeholder_20) — "Кому на подписание"
SIGNERS_19171 = [
    "Заместитель начальника отдела Путькова Т.М.",
    "Начальник отдела Соломейчук А.В.",
    "Заместитель начальника главного управления В.О.Шабловский",
    "Главный специалист Гвоздарев А.А.",
    "Главный специалист Геращенко Г.Н.",
    "Главный специалист Колесан А.И.",
    "Главный специалист Курило И.В.",
    "Главный специалист Новик П.Н.",
    "Главный специалист Одиноченко И.В.",
    "Главный специалист Першко А.С.",
]

# Постоянная часть Placeholder_5
PGREZ_TEXT = 'ГПНИУ "ПГРЭЗ"'

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


def join_districts(districts):
    return ", ".join(districts)


def join_objects(objects_list, custom=""):
    """objects_list — список строк (объектов)"""
    lst = [o for o in objects_list if o]
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


def _today_dmy():
    return datetime.now().strftime("%d.%m.%Y")


def make_individual_19171(data, person):
    """Индивидуальный пропуск для одного лица — по образцу indiv_19.17.1.docx.
    Плейсхолдеры:
    - Placeholder_10.1 — Фамилия
    - Placeholder_10.2 — Имя
    - Placeholder_10.3 — Отчество
    - Placeholder_22 — Наименование организации (место работы)
    - Placeholder_25 — Должность
    - Placeholder_4 — Районы
    - Placeholder_5 — Объект
    - Placeholder_6 — Цель въезда
    - Placeholder_7 — Дата с
    - Placeholder_8 — Дата по
    - Placeholder_20 — Кому на подписание
    """
    doc = Document(os.path.join(template_dir(), "Пропуск_индивидуальный_19.17.1.docx"))
    districts = data.get("districts", [])
    m = {
        "Placeholder_10.1": person.get("last_name", ""),
        "Placeholder_10.2": person.get("first_name", ""),
        "Placeholder_10.3": person.get("middle_name", ""),
        "Placeholder_22": data.get("org_info", ""),
        "Placeholder_25": person.get("position", ""),
        "Placeholder_4": join_districts(districts),
        "Placeholder_5": data.get("objects", ""),
        "Placeholder_6": data.get("goal", ""),
        "Placeholder_7": data.get("date_from", ""),
        "Placeholder_8": data.get("date_to", ""),
        "Placeholder_20": data.get("issued_by", ""),
    }
    fill_doc(doc, m)
    return doc


def make_transport_19171(data, vehicle):
    """Транспортный пропуск для одного авто — по образцу транспортный_19.17.1.docx.
    Плейсхолдеры:
    - Placeholder_1.1, 1.2, 1.3 — Заинтересованное лицо (ФИО представителя организации)
    - Placeholder_4 — Районы
    - Placeholder_5 — Объект
    - Placeholder_6 — Цель въезда
    - Placeholder_7 — Дата с
    - Placeholder_8 — Дата по
    - Placeholder_12 — Марка-модель
    - Placeholder_13 — Регистрационный знак
    - Placeholder_20 — Кому на подписание
    """
    doc = Document(os.path.join(template_dir(), "Пропуск_транспортный_19.17.1.docx"))
    districts = data.get("districts", [])
    # Заинтересованное лицо = представитель организации (если есть) или организация
    org_rep_last = data.get("org_rep_last", "")
    org_rep_first = data.get("org_rep_first", "")
    org_rep_middle = data.get("org_rep_middle", "")
    # Если ФИО представителя не указано, используем название организации
    if not org_rep_last and not org_rep_first and not org_rep_middle:
        org_rep_last = data.get("org_info", "")
    m = {
        "Placeholder_1.1": org_rep_last,
        "Placeholder_1.2": org_rep_first,
        "Placeholder_1.3": org_rep_middle,
        "Placeholder_4": join_districts(districts),
        "Placeholder_5": data.get("objects", ""),
        "Placeholder_6": data.get("goal", ""),
        "Placeholder_7": data.get("date_from", ""),
        "Placeholder_8": data.get("date_to", ""),
        "Placeholder_12": vehicle.get("make", ""),
        "Placeholder_13": vehicle.get("number", ""),
        "Placeholder_20": data.get("issued_by", ""),
    }
    fill_doc(doc, m)
    return doc


def sanitize(s, default="файл"):
    illegal = '<>:"/\\|?*'
    return ("".join(c for c in s if c not in illegal).strip() or default)


def generate_all_19171(data, output_dir=None):
    """Полный комплект документов по п. 19.17.1 (только пропуска, без заявления).
    Возвращает список файлов.
    """
    if output_dir is None:
        output_dir = os.path.join(runtime_base(), "output")
    stamp = datetime.now().strftime("%d.%m.%Y.%H.%M")
    out = os.path.join(output_dir, stamp)
    os.makedirs(out, exist_ok=True)

    created = []
    org = sanitize(data.get("org_short", "Организация"), "Организация")

    # 1. Индивидуальные пропуска для каждого лица
    for i, person in enumerate(data.get("persons", [])):
        d = make_individual_19171(data, person)
        lname = sanitize(person["last_name"], f"Лицо{i+1}")
        p = os.path.join(out, f"Пропуск_Индивидуальный_{lname}.docx")
        d.save(p)
        created.append(p)

    # 2. Транспортные пропуска для каждого авто
    for i, vehicle in enumerate(data.get("vehicles", [])):
        d = make_transport_19171(data, vehicle)
        car = sanitize(vehicle.get("number", "") or f"Авто{i+1}", f"Авто{i+1}")
        p = os.path.join(out, f"Пропуск_Транспорт_{car}.docx")
        d.save(p)
        created.append(p)

    # запись в БД истории
    _db_remember_19171(data)

    return created


# ---------------------------------------------------------------------------
# База знаний
# ---------------------------------------------------------------------------
def _db_path_19171():
    d = os.path.join(runtime_base(), "data")
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, "permit_history_19171.json")


def _db_load_19171():
    import json
    try:
        with open(_db_path_19171(), "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def db_find_19171(org_key):
    """Ищет прошлую запись по ключу организации."""
    try:
        key = (org_key or "").strip().lower()
        if not key:
            return None
        for rec in _db_load_19171():
            rk = str(rec.get("org_key", "")).strip().lower()
            if rk == key:
                return rec
        return None
    except Exception:
        return None


def _db_remember_19171(data):
    try:
        import json
        org_key = (data.get("org_short", "") or data.get("org_info", "")).strip().lower()
        rec = {
            "org_key": org_key,
            "org_info": data.get("org_info", ""),
            "org_short": data.get("org_short", ""),
            "org_rep_last": data.get("org_rep_last", ""),
            "org_rep_first": data.get("org_rep_first", ""),
            "org_rep_middle": data.get("org_rep_middle", ""),
            "goal": data.get("goal", ""),
            "districts": data.get("districts", []),
            "objects": data.get("objects", ""),
            "include_pgrez": data.get("include_pgrez", False),
            "custom_object": data.get("custom_object", ""),
            "date_from": data.get("date_from", ""),
            "date_to": data.get("date_to", ""),
            "issued_by": data.get("issued_by", ""),
            "persons": data.get("persons", []),
            "vehicles": data.get("vehicles", []),
        }
        recs = [r for r in _db_load_19171() if r.get("org_key") != org_key]
        recs.insert(0, rec)
        with open(_db_path_19171(), "w", encoding="utf-8") as f:
            json.dump(recs[:200], f, ensure_ascii=False, indent=1)
    except Exception:
        pass