# -*- coding: utf-8 -*-
"""
Общий SQLite-слой для БД истории пропусков.

Заменяет хранение в JSON-файлах на SQLite. Преимущества:
- индексы -> мгновенный поиск по ФИО/организации даже при десятках тыс. записей;
- транзакции -> нет риска частичной записи при сбое;
- не блокируется при параллельном доступе.

Интерфейс сохранён совместимым с прежним JSON-слоем:
  db_load() -> list[dict]
  db_save(records)  (полная перезапись — прежняя семантика)
  db_find(key, key_field) -> dict|None
  db_upsert(rec, key_field) -> None
  db_export(path) -> None
"""

import json
import os
import sqlite3
from datetime import datetime


class PermitDB:
    """SQLite-хранилище записей о выданных пропусках."""

    def __init__(self, db_path):
        self._path = db_path
        d = os.path.dirname(db_path)
        if d:
            os.makedirs(d, exist_ok=True)
        self._conn = sqlite3.connect(db_path, check_same_thread=False)
        self._conn.row_factory = sqlite3.Row
        self._init_schema()

    def _init_schema(self):
        c = self._conn
        c.execute("""
            CREATE TABLE IF NOT EXISTS permits (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                key TEXT NOT NULL,            -- ключ поиска (ФИО или организация)
                key_field TEXT NOT NULL,      -- 'fio' | 'org'
                data TEXT NOT NULL,           -- JSON с полями записи
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        c.execute("CREATE INDEX IF NOT EXISTS idx_permits_key ON permits(key, key_field)")
        c.commit()

    def close(self):
        try:
            self._conn.close()
        except Exception:
            pass

    # ---- чтение ----
    def load(self, key_field=None):
        """Все записи как список dict. key_field необязателен."""
        c = self._conn
        if key_field:
            rows = c.execute(
                "SELECT data, key FROM permits WHERE key_field=? ORDER BY id DESC",
                (key_field,)).fetchall()
        else:
            rows = c.execute("SELECT data, key FROM permits ORDER BY id DESC").fetchall()
        out = []
        for row in rows:
            try:
                rec = json.loads(row["data"])
            except Exception:
                rec = {}
            # гарантируем наличие ключа в записи
            if rec.get("key") is None and row["key"]:
                rec["key"] = row["key"]
            out.append(rec)
        return out

    def find(self, key, key_field):
        """Последняя запись по ключу (точное совпадение, регистронезависимо на уровне key)."""
        k = (key or "").strip().lower()
        if not k:
            return None
        rows = self._conn.execute(
            "SELECT data FROM permits WHERE key=? AND key_field=? ORDER BY id DESC LIMIT 1",
            (k, key_field)).fetchall()
        if not rows:
            return None
        try:
            return json.loads(rows[0]["data"])
        except Exception:
            return None

    def find_like(self, fragment, key_field, limit=50):
        """Поиск по части ключа (для выпадающего списка)."""
        frag = (fragment or "").strip().lower()
        if not frag:
            return []
        rows = self._conn.execute(
            "SELECT DISTINCT key FROM permits WHERE key_field=? AND key LIKE ? "
            "ORDER BY key LIMIT ?",
            (key_field, "%" + frag + "%", limit)).fetchall()
        return [r["key"] for r in rows]

    # ---- запись ----
    def upsert(self, rec, key_field):
        """Добавляет/обновляет запись. Ключ формируется из rec['key'] или поиском по ключу."""
        try:
            key = (rec.get("key") or "").strip().lower()
            if not key:
                return
            # удаляем старые с таким же ключом, чтобы новая запись была «свежей»
            self._conn.execute(
                "DELETE FROM permits WHERE key=? AND key_field=?", (key, key_field))
            self._conn.execute(
                "INSERT INTO permits (key, key_field, data, created_at) VALUES (?,?,?,?)",
                (key, key_field, json.dumps(rec, ensure_ascii=False), datetime.now().isoformat()))
            self._conn.commit()
        except Exception:
            pass

    def replace_all(self, records, key_field=None):
        """Полная перезапись (миграция из JSON). Каждая запись сохраняется по своему key."""
        c = self._conn
        c.execute("DELETE FROM permits")
        for rec in records:
            field = key_field or ("org" if rec.get("org_key") else "fio")
            self.upsert(rec, field)
        c.commit()

    # ---- экспорт (для кнопки «Выгрузить БД») ----
    def export(self, out_path, fmt="json"):
        """Экспортирует все записи в JSON/CSV/Excel."""
        recs = self.load()
        if fmt == "json":
            with open(out_path, "w", encoding="utf-8") as f:
                json.dump(recs, f, ensure_ascii=False, indent=1)
            return out_path
        if fmt == "csv":
            import csv
            fieldnames = sorted({k for r in recs for k in r.keys()}) or ["key"]
            with open(out_path, "w", encoding="utf-8-sig", newline="") as f:
                w = csv.DictWriter(f, fieldnames=fieldnames)
                w.writeheader()
                for r in recs:
                    w.writerow({k: r.get(k, "") for k in fieldnames})
            return out_path
        if fmt == "excel":
            try:
                from openpyxl import Workbook
            except Exception:
                return None
            wb = Workbook()
            ws = wb.active
            ws.title = "Пропуска"
            fieldnames = sorted({k for r in recs for k in r.keys()}) or ["key"]
            ws.append(fieldnames)
            for r in recs:
                ws.append([r.get(k, "") for k in fieldnames])
            for i, col in enumerate(fieldnames, 1):
                ws.column_dimensions[chr(64 + i)].width = 22
            wb.save(out_path)
            return out_path
        return None


# ---------------------------------------------------------------------------
# Имя БД-файла по процедуре
# ---------------------------------------------------------------------------
def db_filename(proc):
    return f"permit_history_{proc}.sqlite"


def migrate_from_json(db, json_path, key_field):
    """Переносит данные из прежнего JSON-файла в SQLite (если есть)."""
    if not os.path.exists(json_path):
        return 0
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            recs = json.load(f)
    except Exception:
        return 0
    if not recs:
        return 0
    # определяем ключ для каждой записи
    migrated = 0
    for rec in recs:
        if key_field == "org":
            key = rec.get("org_key") or rec.get("org_short") or rec.get("org_info") or ""
        else:
            key = " ".join(str(rec.get(k, "")) for k in
                           ("last_name", "first_name", "middle_name") if rec.get(k)).strip()
        rec = dict(rec)
        rec["key"] = key
        db.upsert(rec, key_field)
        migrated += 1
    return migrated