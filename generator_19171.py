# -*- coding: utf-8 -*-
"""
Генерация документов пропусков по п. 19.17.1.
Только индивидуальные и транспортные пропуска (без заявления).
Использует общие модули permit_common и permit_db.
"""
import os
import sys
from datetime import datetime
from copy import deepcopy

from docx import Document
from docx.oxml.ns import qn

from permit_common import (
    resource_path, runtime_base, template_dir,
    load_reference, filter_objects_by_districts,
    join_districts, join_objects, sanitize, today_dmy,
    fill_doc, _PH_RE,
)
import permit_db as dbm

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


def _today_dmy():
    return today_dmy()


def make_individual_19171(data, person):
    """Индивидуальный пропуск для одного лица (образец indiv_19.17.1.docx)."""
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
    """Транспортный пропуск для одного авто (образец транспортный_19.17.1.docx)."""
    doc = Document(os.path.join(template_dir(), "Пропуск_транспортный_19.17.1.docx"))
    districts = data.get("districts", [])
    org_rep_last = data.get("org_rep_last", "")
    org_rep_first = data.get("org_rep_first", "")
    org_rep_middle = data.get("org_rep_middle", "")
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


def generate_all_19171(data, output_dir=None):
    """Полный комплект документов по п. 19.17.1 (только пропуска)."""
    if output_dir is None:
        output_dir = os.path.join(runtime_base(), "output")
    stamp = datetime.now().strftime("%d.%m.%Y.%H.%M")
    out = os.path.join(output_dir, stamp)
    os.makedirs(out, exist_ok=True)

    created = []
    org = sanitize(data.get("org_short", "Организация"), "Организация")

    for i, person in enumerate(data.get("persons", [])):
        d = make_individual_19171(data, person)
        lname = sanitize(person["last_name"], f"Лицо{i+1}")
        p = os.path.join(out, f"Пропуск_Индивидуальный_{lname}.docx")
        d.save(p)
        created.append(p)

    for i, vehicle in enumerate(data.get("vehicles", [])):
        d = make_transport_19171(data, vehicle)
        car = sanitize(vehicle.get("number", "") or f"Авто{i+1}", f"Авто{i+1}")
        p = os.path.join(out, f"Пропуск_Транспорт_{car}.docx")
        d.save(p)
        created.append(p)

    _db_remember_19171(data)
    return created


# ---------------------------------------------------------------------------
# База знаний — SQLite
# ---------------------------------------------------------------------------
def _db_path_19171():
    d = os.path.join(runtime_base(), "data")
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, dbm.db_filename("19171"))


def _db_conn_19171():
    return dbm.PermitDB(_db_path_19171())


def _db_load_19171():
    try:
        return _db_conn_19171().load()
    except Exception:
        return []


def _db_load_org_19171():
    """Все записи (по организациям). Для UI выпадающего списка."""
    try:
        return _db_conn_19171().load("org")
    except Exception:
        return []


def db_find_19171(org_key):
    return dbm.PermitDB(_db_path_19171()).find(org_key, "org") if (org_key or "").strip() else None


def _db_remember_19171(data):
    db = _db_conn_19171()
    try:
        import json as _json
        legacy = os.path.join(runtime_base(), "data", "permit_history_19171.json")
        if os.path.exists(legacy):
            recs = _json.load(open(legacy, encoding="utf-8"))
            if recs:
                dbm.migrate_from_json(db, legacy, "org")
                os.rename(legacy, legacy + ".migrated")
    except Exception:
        pass
    org_key = (data.get("org_short", "") or data.get("org_info", "")).strip()
    rec = {
        "key": org_key, "org_key": org_key,
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
    db.upsert(rec, "org")
    db.close()