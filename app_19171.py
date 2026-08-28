# -*- coding: utf-8 -*-
"""
Графический интерфейс «Генератор пропусков 19.17.1».
Версия 0.3.8
Создатель: Соломейчук Алексей / Salamiaichuk Aliaksei
Email: al.vl.solo@yandex.by
© 2025 Соломейчук Алексей / Salamiaichuk Aliaksei
Полностью локальное приложение (без интернета и LLM).
Запуск: двойной клик по exe / python app_19171.py
"""
import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from tkcalendar import DateEntry
    HAS_TKCAL = True
except ImportError:
    HAS_TKCAL = False

import generator_19171 as generator

APP_NAME = "Генератор пропусков 19.17.1"
APP_VERSION = "0.3.9"
BG_COLOR = "#E6EBE0"  # светло-серый/голубоватый
BTN_BG = "#CAD4CC"    # чуть темнее для кнопок
BTN_ACTIVE = "#B3C3B8"
STATUS_BG = "#CAD4CC"

# Постоянная часть объекта
PGREZ_TEXT = 'ГПНИУ "ПГРЭЗ"'


class PersonRow(ttk.Frame):
    """Строка одного лица (для вкладки 2)."""
    def __init__(self, master, remove_callback, **kw):
        super().__init__(master, **kw)
        self.remove_callback = remove_callback
        self.var_position = tk.StringVar()
        self.var_last = tk.StringVar()
        self.var_first = tk.StringVar()
        self.var_middle = tk.StringVar()
        
        ttk.Entry(self, textvariable=self.var_position, width=18).grid(row=0, column=0, padx=2)
        ttk.Entry(self, textvariable=self.var_last, width=18).grid(row=0, column=1, padx=2)
        ttk.Entry(self, textvariable=self.var_first, width=16).grid(row=0, column=2, padx=2)
        ttk.Entry(self, textvariable=self.var_middle, width=16).grid(row=0, column=3, padx=2)
        ttk.Button(self, text="×", width=3, command=self._remove).grid(row=0, column=4, padx=2)

    def _remove(self):
        if self.remove_callback:
            self.remove_callback(self)

    def get_data(self):
        return {
            "position": self.var_position.get().strip(),
            "last_name": self.var_last.get().strip(),
            "first_name": self.var_first.get().strip(),
            "middle_name": self.var_middle.get().strip(),
        }

    def set_data(self, data):
        self.var_position.set(data.get("position", ""))
        self.var_last.set(data.get("last_name", ""))
        self.var_first.set(data.get("first_name", ""))
        self.var_middle.set(data.get("middle_name", ""))

    def clear(self):
        self.var_position.set("")
        self.var_last.set("")
        self.var_first.set("")
        self.var_middle.set("")


class VehicleRow(ttk.Frame):
    """Строка одного авто (для вкладки 3)."""
    def __init__(self, master, remove_callback, **kw):
        super().__init__(master, **kw)
        self.remove_callback = remove_callback
        self.var_make = tk.StringVar()
        self.var_number = tk.StringVar()
        
        ttk.Entry(self, textvariable=self.var_make, width=25).grid(row=0, column=0, padx=2)
        ttk.Entry(self, textvariable=self.var_number, width=18).grid(row=0, column=1, padx=2)
        ttk.Button(self, text="×", width=3, command=self._remove).grid(row=0, column=2, padx=2)

    def _remove(self):
        if self.remove_callback:
            self.remove_callback(self)

    def get_data(self):
        return {
            "make": self.var_make.get().strip(),
            "number": self.var_number.get().strip(),
        }

    def set_data(self, data):
        self.var_make.set(data.get("make", ""))
        self.var_number.set(data.get("number", ""))

    def clear(self):
        self.var_make.set("")
        self.var_number.set("")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} v{APP_VERSION}")
        self.geometry("850x800")
        self.minsize(800, 720)
        self.configure(bg=BG_COLOR)

        # --- переменные данных ---
        self.var_org_info = tk.StringVar()         # Организация-заявитель
        self.var_org_short = tk.StringVar()        # краткое название для имен файлов
        
        self.var_goal = tk.StringVar()             # Цель въезда (свободное поле)
        self.var_date_from = tk.StringVar()
        self.var_date_to = tk.StringVar()
        self.var_issued = tk.StringVar(value=generator.SIGNERS_19171[0])
        self.var_custom_object = tk.StringVar()
        self.var_include_pgrez = tk.BooleanVar(value=False)  # чекбокс ГПНИУ "ПГРЭЗ" - по умолчанию ВЫКЛЮЧЕН
        self.var_output = tk.StringVar(value=self._default_output())

        # списки строк
        self.person_rows = []
        self.vehicle_rows = []

        self.district_vars = {}
        for d in generator.DISTRICTS:
            self.district_vars[d] = tk.BooleanVar(value=False)

        self._build_ui()
        self._refresh_district_checks()
        self._refresh_org_db_list()

        # стиль для ttk
        style = ttk.Style(self)
        style.theme_use("default")
        for s in ("TFrame", "TLabel", "TCheckbutton", "TRadiobutton", "TNotebook", "TNotebook.Tab",
                  "TLabelFrame", "TLabelframe.Label", "TSeparator"):
            try:
                style.configure(s, background=BG_COLOR)
            except Exception:
                pass
        style.configure("TButton", background=BTN_BG, foreground="black")
        style.map("TButton", background=[("active", BTN_ACTIVE)])

        self.status_bar = tk.Label(self, text=f"Готов к работе.  Шаблоны: {self._tpl_count()}",
                                   anchor="w", bg=STATUS_BG, fg="black")
        self.status_bar.pack(side="bottom", fill="x")

    # ------------------------------------------------------------ UI-сборка
    def _default_output(self):
        return os.path.join(self._app_base(), "output")

    def _app_base(self):
        if getattr(sys, "frozen", False):
            return os.path.dirname(os.path.abspath(sys.argv[0]))
        return generator.resource_path()

    def _tpl_count(self):
        try:
            d = generator.template_dir()
            return str(len([f for f in os.listdir(d) if f.endswith(".docx")]))
        except Exception:
            return "?"

    def _build_ui(self):
        pad = dict(padx=8, pady=3)
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=6)

        tab1 = ttk.Frame(nb); nb.add(tab1, text="1. Заявитель и цель")
        self._build_tab1(tab1, pad)

        tab2 = ttk.Frame(nb); nb.add(tab2, text="2. Направляемые лица")
        self._build_tab2(tab2, pad)

        tab3 = ttk.Frame(nb); nb.add(tab3, text="3. Автомобили")
        self._build_tab3(tab3, pad)

        tab4 = ttk.Frame(nb); nb.add(tab4, text="4. Районы и объекты")
        self._build_tab4(tab4, pad)

        # вкладка "О программе"
        tab_about = ttk.Frame(nb); nb.add(tab_about, text="О программе")
        self._build_about_tab(tab_about)

        # кнопки — grid с четырьмя равными колонками
        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=8, pady=(0, 8))
        for i in range(4):
            btns.columnconfigure(i, weight=1)
        ttk.Button(btns, text="Сгенерировать пропуска", command=self.generate).grid(row=0, column=0, sticky="ew", padx=2)
        ttk.Button(btns, text="Очистить форму", command=self.clear_form).grid(row=0, column=1, sticky="ew", padx=2)
        ttk.Button(btns, text="Куда сохранять…", command=self.choose_output).grid(row=0, column=2, sticky="ew", padx=2)
        ttk.Button(btns, text="Открыть папку", command=self.open_output).grid(row=0, column=3, sticky="ew", padx=2)

    def _entry(self, parent, label, var, row, col, width=30, **kw):
        ttk.Label(parent, text=label).grid(row=row, column=col, sticky="w", **kw)
        ent = ttk.Entry(parent, textvariable=var, width=width)
        ent.grid(row=row, column=col + 1, sticky="ew", **kw)
        return ent

    def _build_tab1(self, parent, pad):
        parent.columnconfigure(1, weight=1)
        parent.columnconfigure(3, weight=1)
        r = 0

        ttk.Label(parent, text="Организация-заявитель", font=("", 10, "bold")).grid(row=r, column=0, columnspan=4, sticky="w", **pad); r += 1

        # Поиск в базе по наименованию организации
        ttk.Label(parent, text="Поиск в базе (введите или выберите организацию):").grid(row=r, column=0, columnspan=2, sticky="w", **pad)
        self.db_org_cb = ttk.Combobox(parent, width=50, state="normal")
        self.db_org_cb.grid(row=r, column=2, columnspan=2, sticky="ew", **pad)
        self.db_org_cb.bind("<KeyRelease>", self._db_org_search_typed)
        self.db_org_cb.bind("<<ComboboxSelected>>", self._on_pick_org)
        r += 1

        self._entry(parent, "Организация-заявитель:", self.var_org_info, r, 0, width=60, **pad); r += 1

        ttk.Separator(parent, orient="horizontal").grid(row=r, column=0, columnspan=4, sticky="ew", **pad); r += 1

        ttk.Label(parent, text="Цель въезда:", font=("", 10, "bold")).grid(row=r, column=0, sticky="w", **pad)
        # просто пустое поле для свободного заполнения
        ttk.Entry(parent, textvariable=self.var_goal, width=55).grid(row=r, column=1, columnspan=3, sticky="ew", **pad)
        r += 1

        ttk.Separator(parent, orient="horizontal").grid(row=r, column=0, columnspan=4, sticky="ew", **pad); r += 1
        self._date_row(parent, "Срок действия — с:", self.var_date_from, r, **pad); r += 1
        self._date_row(parent, "Срок действия — по:", self.var_date_to, r, **pad); r += 1
        # "Дата заявления" убрана

    def _build_tab2(self, parent, pad):
        parent.columnconfigure(0, weight=1)
        ttk.Label(parent, text="Направляемые лица (индивидуальные пропуска). Минимум 1.", font=("", 10, "bold")).grid(row=0, column=0, sticky="w", **pad)
        
        # заголовки - без placeholder_25
        hdr = ttk.Frame(parent)
        hdr.grid(row=1, column=0, sticky="ew", **pad)
        for c, t in enumerate(["Должность", "Фамилия", "Имя", "Отчество"]):
            ttk.Label(hdr, text=t, font=("", 9, "bold")).grid(row=0, column=c, sticky="w", padx=2)
        
        self.persons_frame = ttk.Frame(parent)
        self.persons_frame.grid(row=2, column=0, sticky="nsew", **pad)
        parent.rowconfigure(2, weight=1)
        
        btns = ttk.Frame(parent)
        btns.grid(row=3, column=0, sticky="w", **pad)
        ttk.Button(btns, text="+ Добавить лицо", command=self.add_person).pack(side="left")
        ttk.Button(btns, text="Очистить всех", command=self.clear_persons).pack(side="left", padx=6)
        
        # добавляем одну пустую строку по умолчанию
        self.add_person()

    def _build_tab3(self, parent, pad):
        parent.columnconfigure(0, weight=1)
        ttk.Label(parent, text="Транспортные средства (транспортные пропуска). Минимум 1.", font=("", 10, "bold")).grid(row=0, column=0, sticky="w", **pad)
        
        # заголовки
        hdr = ttk.Frame(parent)
        hdr.grid(row=1, column=0, sticky="ew", **pad)
        for c, t in enumerate(["Марка-модель", "Регистрационный знак"]):
            ttk.Label(hdr, text=t, font=("", 9, "bold")).grid(row=0, column=c, sticky="w", padx=2)
        
        self.vehicles_frame = ttk.Frame(parent)
        self.vehicles_frame.grid(row=2, column=0, sticky="nsew", **pad)
        parent.rowconfigure(2, weight=1)
        
        btns = ttk.Frame(parent)
        btns.grid(row=3, column=0, sticky="w", **pad)
        ttk.Button(btns, text="+ Добавить авто", command=self.add_vehicle).pack(side="left")
        ttk.Button(btns, text="Очистить все", command=self.clear_vehicles).pack(side="left", padx=6)
        
        # добавляем одну пустую строку по умолчанию
        self.add_vehicle()

    def _build_tab4(self, parent, pad):
        parent.columnconfigure(1, weight=1)
        r = 0

        ttk.Label(parent, text="Район(ы):", font=("", 10, "bold")).grid(row=r, column=0, sticky="nw", **pad)
        rf = ttk.Frame(parent); rf.grid(row=r, column=1, columnspan=3, sticky="w", **pad)
        col = 0; row2 = 0
        for d in generator.DISTRICTS:
            cb = ttk.Checkbutton(rf, text=d, variable=self.district_vars[d],
                                 command=self._refresh_district_checks)
            cb.grid(row=row2, column=col, sticky="w", padx=2, pady=1)
            col += 1
            if col > 3:
                col = 0; row2 += 1
        r += 1

        ttk.Separator(parent, orient="horizontal").grid(row=r, column=0, columnspan=4, sticky="ew", **pad); r += 1

        ttk.Label(parent, text="Объект(ы):", font=("", 10, "bold")).grid(row=r, column=0, columnspan=4, sticky="w", **pad); r += 1
        
        # чекбокс ГПНИУ "ПГРЭЗ"
        cb_pgrez = ttk.Checkbutton(parent, text=PGREZ_TEXT, variable=self.var_include_pgrez)
        cb_pgrez.grid(row=r, column=0, columnspan=4, sticky="w", **pad); r += 1
        
        # свободное поле (без placeholder_5)
        self._entry(parent, "Произвольный объект:", self.var_custom_object, r, 0, width=60, **pad); r += 1
        
        ttk.Separator(parent, orient="horizontal").grid(row=r, column=0, columnspan=4, sticky="ew", **pad); r += 1
        
        # "Кому на подписание" вместо "Кому на согласование" (без placeholder_20)
        ttk.Label(parent, text="Кому на подписание:", font=("", 10, "bold")).grid(row=r, column=0, sticky="w", **pad)
        ttk.Combobox(parent, textvariable=self.var_issued, width=55,
                     values=generator.SIGNERS_19171, state="readonly").grid(row=r, column=1, columnspan=3, sticky="w", **pad)
        r += 1

        ttk.Separator(parent, orient="horizontal").grid(row=r, column=0, columnspan=4, sticky="ew", **pad); r += 1
        ttk.Label(parent, text="Папка для сохранения:", font=("", 10, "bold")).grid(row=r, column=0, columnspan=4, sticky="w", **pad); r += 1
        ttk.Label(parent, textvariable=self.var_output, foreground="#556B5E").grid(row=r, column=0, columnspan=4, sticky="w", **pad); r += 1

    # ------------------------------------------------------------- даты
    def _date_row(self, parent, label, var, row, **pad):
        """Строка с датой: поле с маской «ДД.ММ.ГГГГ» + кнопка календаря."""
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", **pad)
        frame = ttk.Frame(parent)
        frame.grid(row=row, column=1, columnspan=3, sticky="w", **pad)
        ent = ttk.Entry(frame, textvariable=var, width=15)
        ent.pack(side="left")
        vcmd = (self.register(lambda P: all(ch.isdigit() or ch == "." for ch in P) and P.count(".") <= 2), "%P")
        ent.configure(validate="key", validatecommand=vcmd)
        ent.bind("<KeyRelease>", lambda e, v=var, w=ent: self._apply_date_mask(v, w))
        btn = ttk.Button(frame, text="📅", width=3, command=lambda: self._pick_date(var, ent))
        btn.pack(side="left", padx=(4, 0))

    def _apply_date_mask(self, var, ent=None):
        s = var.get()
        digits = "".join(ch for ch in s if ch.isdigit())[:8]
        out = digits
        if len(digits) > 4:
            out = digits[:2] + "." + digits[2:4] + "." + digits[4:]
        elif len(digits) > 2:
            out = digits[:2] + "." + digits[2:]
        if out != s:
            pos = len(s)
            if ent:
                try:
                    pos = ent.index(tk.INSERT)
                except Exception:
                    pass
            var.set(out)
            if ent:
                new_pos = min(pos + (len(out) - len(s)), len(out))
                try:
                    ent.icursor(new_pos)
                except Exception:
                    pass

    def _pick_date(self, var, ent):
        if HAS_TKCAL:
            from tkinter import Toplevel
            top = Toplevel(self)
            top.title("Выбор даты")
            top.transient(self)
            top.grab_set()
            de = DateEntry(top, date_pattern="dd.mm.yyyy", locale="ru_RU",
                           firstweekday="monday", width=14)
            de.pack(padx=12, pady=12)
            def _set(event=None):
                try:
                    var.set(de.get_date().strftime("%d.%m.%Y"))
                except Exception:
                    pass
                top.destroy()
                ent.focus_set()
            de.bind("<<DateEntrySelected>>", _set)
            top.bind("<Escape>", lambda e: top.destroy())
            self.update_idletasks()
            x = ent.winfo_rootx()
            y = ent.winfo_rooty() + ent.winfo_height()
            scr_h = self.winfo_screenheight()
            cal_h = 250
            if y + cal_h > scr_h:
                y = ent.winfo_rooty() - cal_h
            top.geometry("+%d+%d" % (x, y))
        else:
            from tkinter import simpledialog
            val = simpledialog.askstring("Выбор даты", "Введите дату (ДД.ММ.ГГГГ):", parent=self)
            if val and len(val.replace(".", "")) == 8:
                var.set(val.strip())

    # ------------------------------------------------------------- списки лиц и авто
    def add_person(self):
        row = PersonRow(self.persons_frame, self._remove_person)
        row.grid(row=len(self.person_rows), column=0, sticky="ew", pady=2)
        self.person_rows.append(row)

    def _remove_person(self, row_widget):
        if len(self.person_rows) <= 1:
            messagebox.showwarning(APP_NAME, "Должно быть минимум одно лицо.")
            return
        row_widget.destroy()
        self.person_rows.remove(row_widget)
        # перестраиваем grid
        for i, r in enumerate(self.person_rows):
            r.grid(row=i, column=0)

    def clear_persons(self):
        for r in self.person_rows:
            r.destroy()
        self.person_rows.clear()
        self.add_person()

    def add_vehicle(self):
        row = VehicleRow(self.vehicles_frame, self._remove_vehicle)
        row.grid(row=len(self.vehicle_rows), column=0, sticky="ew", pady=2)
        self.vehicle_rows.append(row)

    def _remove_vehicle(self, row_widget):
        if len(self.vehicle_rows) <= 1:
            messagebox.showwarning(APP_NAME, "Должно быть минимум одно авто.")
            return
        row_widget.destroy()
        self.vehicle_rows.remove(row_widget)
        for i, r in enumerate(self.vehicle_rows):
            r.grid(row=i, column=0)

    def clear_vehicles(self):
        for r in self.vehicle_rows:
            r.destroy()
        self.vehicle_rows.clear()
        self.add_vehicle()

    # ------------------------------------------------------------- действия
    def _selected_districts(self):
        return [d for d, v in self.district_vars.items() if v.get()]

    def _refresh_district_checks(self):
        pass

    def _collect_persons(self):
        persons = []
        for row in self.person_rows:
            data = row.get_data()
            if any(data.values()):
                persons.append(data)
        return persons

    def _collect_vehicles(self):
        vehicles = []
        for row in self.vehicle_rows:
            data = row.get_data()
            if any(data.values()):
                vehicles.append(data)
        return vehicles

    def _build_objects_string(self):
        """Собирает строку объектов для Placeholder_5."""
        parts = []
        if self.var_include_pgrez.get():
            parts.append(PGREZ_TEXT)
        custom = self.var_custom_object.get().strip()
        if custom:
            parts.append(custom)
        return "; ".join(parts)

    # ------------------------------------------------------------- поиск в базе по организации
    def _refresh_org_db_list(self):
        """Загружает список организаций из базы и заполняет Combobox."""
        try:
            recs = generator._db_load_19171()
            values = []
            seen = set()
            for r in recs:
                name = (r.get("org_info") or "").strip()
                if name and name.lower() not in seen:
                    seen.add(name.lower())
                    values.append(name)
            self.db_org_cb["values"] = values
        except Exception:
            self.db_org_cb["values"] = []

    def _db_org_search_typed(self, event):
        """Фильтрует список организаций по вводу и открывает выпадающий список."""
        # игнорируем служебные клавиши
        if event.keysym in ("Up", "Down", "Return", "Escape", "Tab", "Left", "Right"):
            return
        try:
            recs = generator._db_load_19171()
            typed = self.db_org_cb.get().strip().lower()
            values = []
            seen = set()
            for r in recs:
                name = (r.get("org_info") or "").strip()
                if not name:
                    continue
                if typed and typed not in name.lower():
                    continue
                if name.lower() not in seen:
                    seen.add(name.lower())
                    values.append(name)
            self.db_org_cb["values"] = values[:50]
            self.db_org_cb.event_generate("<Down>")
        except Exception:
            pass

    def _on_pick_org(self, event=None):
        """Заполняет форму данными из базы по выбранной организации."""
        try:
            name = self.db_org_cb.get().strip()
            if not name:
                return
            rec = generator.db_find_19171(name)
            if not rec:
                return
            # Организация
            self.var_org_info.set(rec.get("org_info", ""))
            self.var_org_short.set(rec.get("org_short", ""))
            # Цель
            if rec.get("goal"):
                self.var_goal.set(rec.get("goal", ""))
            # Районы
            for d, v in self.district_vars.items():
                v.set(d in (rec.get("districts") or []))
            # ГПНИУ
            self.var_include_pgrez.set(bool(rec.get("include_pgrez", False)))
            # Объекты
            self.var_custom_object.set(rec.get("custom_object", ""))
            # Даты
            if rec.get("date_from"):
                self.var_date_from.set(rec.get("date_from", ""))
            if rec.get("date_to"):
                self.var_date_to.set(rec.get("date_to", ""))
            # Подписант
            if rec.get("issued_by"):
                self.var_issued.set(rec.get("issued_by", ""))
            # Лица
            persons = rec.get("persons") or []
            for r in self.person_rows:
                r.destroy()
            self.person_rows.clear()
            if persons:
                for p in persons:
                    self.add_person()
                    self.person_rows[-1].set_data(p)
            else:
                self.add_person()
            # Авто
            vehicles = rec.get("vehicles") or []
            for r in self.vehicle_rows:
                r.destroy()
            self.vehicle_rows.clear()
            if vehicles:
                for v in vehicles:
                    self.add_vehicle()
                    self.vehicle_rows[-1].set_data(v)
            else:
                self.add_vehicle()
            self.status_bar.config(text=f"База знаний: подставлены данные для «{rec.get('org_info', '')}»")
        except Exception as ex:
            self.status_bar.config(text=f"Ошибка подстановки: {ex}")

    def clear_form(self):
        """Очистка всей формы."""
        for v in (self.var_org_info, self.var_org_short,
                  self.var_goal, self.var_custom_object, self.var_date_from, self.var_date_to):
            v.set("")
        self.var_issued.set(generator.SIGNERS_19171[0])
        self.var_include_pgrez.set(False)
        for d, v in self.district_vars.items():
            v.set(False)
        # очищаем списки
        for r in self.person_rows:
            r.destroy()
        self.person_rows.clear()
        self.add_person()
        for r in self.vehicle_rows:
            r.destroy()
        self.vehicle_rows.clear()
        self.add_vehicle()
        # обновляем выпадающий список организаций (если в базе появились новые)
        try:
            self.db_org_cb.set("")
            self._refresh_org_db_list()
        except Exception:
            pass
        self.status_bar.config(text="Форма очищена.")

    def choose_output(self):
        d = filedialog.askdirectory(title="Выберите папку для сохранения документов")
        if d:
            self.var_output.set(d)

    def open_output(self):
        d = self.var_output.get()
        os.makedirs(d, exist_ok=True)
        os.startfile(d)

    def generate(self):
        try:
            org_info = self.var_org_info.get().strip()
            if not org_info:
                messagebox.showwarning(APP_NAME, "Укажите наименование организации-заявителя.")
                return

            districts = self._selected_districts()
            if not districts:
                messagebox.showwarning(APP_NAME, "Отметьте хотя бы один район.")
                return

            persons = self._collect_persons()
            if not persons:
                messagebox.showwarning(APP_NAME, "Добавьте хотя бы одно направляемое лицо.")
                return
            for p in persons:
                if not p["last_name"] or not p["first_name"]:
                    messagebox.showwarning(APP_NAME, "У всех лиц должны быть заполнены фамилия и имя.")
                    return

            vehicles = self._collect_vehicles()
            if not vehicles:
                messagebox.showwarning(APP_NAME, "Добавьте хотя бы одно транспортное средство.")
                return
            for v in vehicles:
                if not v["make"] or not v["number"]:
                    messagebox.showwarning(APP_NAME, "У всех авто должны быть заполнены марка-модель и гос. номер.")
                    return

            objects = self._build_objects_string()
            if not objects:
                messagebox.showwarning(APP_NAME, "Укажите объект(ы) — отметьте чекбокс ГПНИУ \"ПГРЭЗ\" или впишите свой.")
                return

            goal = self.var_goal.get().strip()
            if not goal:
                messagebox.showwarning(APP_NAME, "Укажите цель въезда.")
                return

            data = {
                "org_info": org_info,
                "org_short": self.var_org_short.get().strip() or org_info[:30],
                "goal": goal,
                "districts": districts,
                "objects": objects,
                "include_pgrez": self.var_include_pgrez.get(),
                "custom_object": self.var_custom_object.get().strip(),
                "date_from": self.var_date_from.get().strip(),
                "date_to": self.var_date_to.get().strip(),
                "issued_by": self.var_issued.get().strip(),
                "persons": persons,
                "vehicles": vehicles,
            }

            outdir = self.var_output.get().strip() or self._default_output()
            files = generator.generate_all_19171(data, outdir)
            msg = "Готово! Создано документов: %d\n\n%s" % (len(files), "\n".join(os.path.basename(f) for f in files))
            if messagebox.askyesno(APP_NAME, msg + "\n\nОткрыть папку с документами?"):
                os.makedirs(os.path.dirname(files[0]), exist_ok=True)
                os.startfile(os.path.dirname(files[0]))
            # обновляем список организаций в выпадающем списке
            try:
                self._refresh_org_db_list()
            except Exception:
                pass
            self.status_bar.config(text="Создано файлов: %d → %s" % (len(files), os.path.dirname(files[0])))
        except Exception as ex:
            messagebox.showerror(APP_NAME, "Ошибка генерации:\n%s" % ex)

    def _build_about_tab(self, parent):
        txt = (
            f"{APP_NAME}\n"
            f"Версия {APP_VERSION}\n\n"
            "Создатель: Соломейчук Алексей\n"
            "Creator: Salamiaichuk Aliaksei\n\n"
            "Email: al.vl.solo@yandex.by\n\n"
            "© 2025 Соломейчук Алексей / Salamiaichuk Aliaksei\n"
            "Все права защищены.\n\n"
            "Программа для генерации документов пропусков\n"
            "по п. 19.17.1 перечня административных процедур.\n"
            "Работает полностью локально, без интернета.\n\n"
            "Состав документов:\n"
            "  • Индивидуальные пропуска (по количеству лиц)\n"
            "  • Транспортные пропуска (по количеству авто)\n\n"
            "Особенности:\n"
            "  • Placeholder_5: чекбокс ГПНИУ \"ПГРЭЗ\" + свободное поле\n"
            "  • Placeholder_25: должность для каждого лица\n"
            "  • Placeholder_20: выбор из списка главных специалистов\n"
            "  • Цель въезда: свободное поле\n"
            "  • Заявление не генерируется"
        )
        lbl = ttk.Label(parent, text=txt, justify="left", anchor="nw", wraplength=750)
        lbl.pack(fill="both", expand=True, padx=16, pady=16)


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()