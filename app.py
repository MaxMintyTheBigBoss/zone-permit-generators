# -*- coding: utf-8 -*-
"""
Графический интерфейс «Генератор пропусков 14.3» по промту.
Полностью локальное приложение (без интернета и LLM).
Запуск: двойной клик по exe / python app.py
"""
import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import permit_update
import permit_db as pdb_mod
from permit_update_gui import UpdateDialog, export_db_dialog

try:
    from tkcalendar import DateEntry
    HAS_TKCAL = True
except ImportError:
    HAS_TKCAL = False

import generator

APP_NAME = "Генератор_пропусков_14.3"


class CheckListbox(ttk.Frame):
    """Скроллируемый список чекбоксов (галочек) — выбор объектов галочкой."""

    def __init__(self, master, height=12, **kw):
        super().__init__(master, **kw)
        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0, bg="#D4FCEE")
        self.scroll = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scroll.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scroll.pack(side="right", fill="y")
        self.inner = ttk.Frame(self.canvas)
        self.win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(self.win, width=e.width))
        self.items = []  # list of (text, BooleanVar)
        # колесо мыши активно только когда курсор над списком
        self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", self._on_wheel))
        self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))

    def _on_wheel(self, event):
        try:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            return "break"
        except Exception:
            return None

    def set_items(self, items):
        for w in self.inner.winfo_children():
            w.destroy()
        self.items = []
        if not items:
            tk.Label(self.inner, text="(не выбрано районов — отметьте район(ы))",
                     fg="gray", anchor="w", bg="#D4FCEE").pack(anchor="w", padx=6, pady=6)
        for it in items:
            var = tk.BooleanVar(value=False)
            cb = ttk.Checkbutton(self.inner, text=it, variable=var)
            cb.pack(anchor="w", padx=6, pady=0)
            self.items.append((it, var))

    def get_checked(self):
        return [t for t, v in self.items if v.get()]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("780x780")
        self.minsize(740, 700)
        self.configure(bg="#D4FCEE")

        self.reference = generator.load_reference()
        self.obj_to_district = {r["object"]: r["district"] for r in self.reference}

        # --- переменные данных ---
        self.var_lname = tk.StringVar()
        self.var_fname = tk.StringVar()
        self.var_mname = tk.StringVar()
        self.var_birth = tk.StringVar()
        self.var_idnum = tk.StringVar()
        self.var_goal = tk.StringVar(value=generator.GOAL_BLAGOUSTROISTVO)
        self.var_date_from = tk.StringVar()
        self.var_date_to = tk.StringVar()
        self.var_app_date = tk.StringVar(value=generator._today_dmy())
        self.var_car_on = tk.BooleanVar(value=False)
        self.var_car_make = tk.StringVar()
        self.var_car_num = tk.StringVar()
        self.var_issued = tk.StringVar(value=generator.SIGNERS[0])
        self.var_custom_objects = tk.StringVar()
        self.var_output = tk.StringVar(value=self._default_output())

        # привязка полей к событию автоподстановки из базы
        for v in (self.var_lname, self.var_fname, self.var_mname):
            v.trace_add("write", lambda *a: self._maybe_autofill())

        self.district_vars = {}
        for d in generator.DISTRICTS:
            self.district_vars[d] = tk.BooleanVar(value=False)

        self._build_ui()
        self._refresh_objects()

        # стиль для ttk — фон #D4FCEE
        style = ttk.Style(self)
        style.theme_use("default")
        for s in ("TFrame", "TLabel", "TCheckbutton", "TRadiobutton", "TNotebook", "TNotebook.Tab",
                  "TLabelFrame", "TLabelframe.Label", "TSeparator"):
            try:
                style.configure(s, background="#D4FCEE")
            except Exception:
                pass
        style.configure("TButton", background="#8CD9C6", foreground="black")
        style.map("TButton", background=[("active", "#6BCCAE")])

        self.status_bar = tk.Label(self, text="Готов к работе.  Шаблоны: %s" % self._tpl_count(),
                                   anchor="w", bg="#8CD9C6", fg="black")
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

        tab2 = ttk.Frame(nb); nb.add(tab2, text="2. Пассажиры")
        self._build_tab2(tab2, pad)

        tab3 = ttk.Frame(nb); nb.add(tab3, text="3. Авто и выдача")
        self._build_tab3(tab3, pad)

        # вкладка "О программе"
        tab_about = ttk.Frame(nb); nb.add(tab_about, text="О программе")
        self._build_about_tab(tab_about)

        # кнопки — фрейм с фиксированной высотой, не съезжает
        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=8, pady=(0, 8))
        # растягиваем колонки поровну
        for i in range(6):
            btns.columnconfigure(i, weight=1)
        ttk.Button(btns, text="Сгенерировать документы", command=self.generate).grid(row=0, column=0, sticky="ew", padx=2)
        ttk.Button(btns, text="Очистить форму", command=self.clear_form).grid(row=0, column=1, sticky="ew", padx=2)
        ttk.Button(btns, text="Куда сохранять…", command=self.choose_output).grid(row=0, column=2, sticky="ew", padx=2)
        ttk.Button(btns, text="Открыть папку", command=self.open_output).grid(row=0, column=3, sticky="ew", padx=2)
        ttk.Button(btns, text="Выгрузить БД", command=self.export_db).grid(row=0, column=4, sticky="ew", padx=2)
        ttk.Button(btns, text="Обновления…", command=self.check_updates).grid(row=0, column=5, sticky="ew", padx=2)

    def _entry(self, parent, label, var, row, col, width=30, **kw):
        ttk.Label(parent, text=label).grid(row=row, column=col, sticky="w", **kw)
        ent = ttk.Entry(parent, textvariable=var, width=width)
        ent.grid(row=row, column=col + 1, sticky="ew", **kw)
        return ent

    def _build_tab1(self, parent, pad):
        parent.columnconfigure(1, weight=1)
        parent.columnconfigure(3, weight=1)
        r = 0

        ttk.Label(parent, text="Заявитель", font=("", 10, "bold")).grid(row=r, column=0, columnspan=4, sticky="w", **pad); r += 1
        self._entry(parent, "Фамилия:", self.var_lname, r, 0, **pad); r += 1
        self._entry(parent, "Имя:", self.var_fname, r, 0, **pad); r += 1
        self._entry(parent, "Отчество:", self.var_mname, r, 0, **pad); r += 1
        self._date_row(parent, "Дата рождения:", self.var_birth, r, **pad); r += 1
        self._entry(parent, "Личный номер:", self.var_idnum, r, 0, width=20, **pad); r += 1

        # поиск по базе прошлых заявок
        ttk.Label(parent, text="Поиск в базе (введите или выберите ФИО):").grid(row=r, column=0, sticky="w", **pad)
        self.db_cb = ttk.Combobox(parent, width=38, state="normal", postcommand=self._refresh_db_list)
        self.db_cb.grid(row=r, column=1, columnspan=3, sticky="w", **pad)
        self.db_cb.bind("<KeyRelease>", self._db_search_typed)
        self.db_cb.bind("<<ComboboxSelected>>", self._on_pick_person)
        r += 1

        ttk.Separator(parent, orient="horizontal").grid(row=r, column=0, columnspan=4, sticky="ew", **pad); r += 1

        ttk.Label(parent, text="Цель въезда:", font=("", 10, "bold")).grid(row=r, column=0, sticky="w", **pad)
        goal_cb = ttk.Combobox(parent, textvariable=self.var_goal, width=34,
                               values=[generator.GOAL_BLAGOUSTROISTVO, generator.GOAL_CUSTOM],
                               state="normal")
        goal_cb.grid(row=r, column=1, columnspan=3, sticky="w", **pad)
        ttk.Label(parent, text="(выберите «благоустройство…» — из списка ниже; «свой вариант» — впишите свою цель)").grid(
            row=r + 1, column=0, columnspan=4, sticky="w", **pad)
        r += 2

        ttk.Label(parent, text="Район(ы):", font=("", 10, "bold")).grid(row=r, column=0, sticky="nw", **pad)
        rf = ttk.Frame(parent); rf.grid(row=r, column=1, columnspan=3, sticky="w", **pad)
        col = 0; row2 = 0
        for d in generator.DISTRICTS:
            cb = ttk.Checkbutton(rf, text=d, variable=self.district_vars[d],
                                 command=self._refresh_objects)
            cb.grid(row=row2, column=col, sticky="w", padx=2, pady=1)
            col += 1
            if col > 3:
                col = 0; row2 += 1
        r += 1

        ttk.Label(parent, text="Объект(ы) / кладбища (отметьте галочкой нужные):", font=("", 10, "bold")).grid(
            row=r, column=0, columnspan=4, sticky="w", **pad); r += 1
        lf = ttk.Frame(parent); lf.grid(row=r, column=0, columnspan=4, sticky="nsew", **pad)
        parent.rowconfigure(r, weight=1)
        self.obj_check = CheckListbox(lf)
        self.obj_check.pack(side="left", fill="both", expand=True)
        r += 1
        ttk.Label(parent, text="(показываются кладбища только выбранных районов)").grid(row=r, column=0, columnspan=4, sticky="w", **pad); r += 1
        self._entry(parent, "Произвольный объект (если нет в списке):", self.var_custom_objects, r, 0, width=40, **pad); r += 1

        ttk.Separator(parent, orient="horizontal").grid(row=r, column=0, columnspan=4, sticky="ew", **pad); r += 1
        self._date_row(parent, "Срок действия — с:", self.var_date_from, r, **pad); r += 1
        self._date_row(parent, "Срок действия — по:", self.var_date_to, r, **pad); r += 1
        self._date_row(parent, "Дата заявления:", self.var_app_date, r, **pad); r += 1

    def _date_row(self, parent, label, var, row, **pad):
        """Строка с датой: поле с маской «ДД.ММ.ГГГГ» + кнопка календаря."""
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", **pad)
        frame = ttk.Frame(parent)
        frame.grid(row=row, column=1, columnspan=3, sticky="w", **pad)
        ent = ttk.Entry(frame, textvariable=var, width=15)
        ent.pack(side="left")
        # маска: только цифры и точки, авто-вставка точек «  .  .    »
        vcmd = (self.register(lambda P: all(ch.isdigit() or ch == "." for ch in P) and P.count(".") <= 2), "%P")
        ent.configure(validate="key", validatecommand=vcmd)
        ent.bind("<KeyRelease>", lambda e, v=var, w=ent: self._apply_date_mask(v, w))
        btn = ttk.Button(frame, text="📅", width=3, command=lambda: self._pick_date(var, ent))
        btn.pack(side="left", padx=(4, 0))

    def _apply_date_mask(self, var, ent=None):
        """Форматирует ввод в ДД.ММ.ГГГГ по мере набора, сохраняя позицию курсора."""
        s = var.get()
        digits = "".join(ch for ch in s if ch.isdigit())[:8]
        out = digits
        if len(digits) > 4:
            out = digits[:2] + "." + digits[2:4] + "." + digits[4:]
        elif len(digits) > 2:
            out = digits[:2] + "." + digits[2:]
        if out != s:
            # сохранить позицию курсора
            pos = len(s)
            if ent:
                try:
                    pos = ent.index(tk.INSERT)
                except Exception:
                    pass
            var.set(out)
            if ent:
                # скорректировать позицию: если добавилась точка перед курсором, сдвинуть вправо
                new_pos = min(pos + (len(out) - len(s)), len(out))
                try:
                    ent.icursor(new_pos)
                except Exception:
                    pass

    def _pick_date(self, var, ent):
        """Выбор даты из встроенного календаря."""
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
            # позиционирование: под полем, но не ниже экрана
            self.update_idletasks()
            x = ent.winfo_rootx()
            y = ent.winfo_rooty() + ent.winfo_height()
            # высота экрана и предполагаемая высота календаря
            scr_h = self.winfo_screenheight()
            cal_h = 250  # примерно высота DateEntry
            if y + cal_h > scr_h:
                y = ent.winfo_rooty() - cal_h  # открыть над полем
            top.geometry("+%d+%d" % (x, y))
        else:
            # запасной вариант без tkcalendar — простое окно ручного выбора
            from tkinter import simpledialog
            val = simpledialog.askstring("Выбор даты", "Введите дату (ДД.ММ.ГГГГ):", parent=self)
            if val and len(val.replace(".", "")) == 8:
                var.set(val.strip())

    def _build_tab2(self, parent, pad):
        parent.columnconfigure(0, weight=1)
        ttk.Label(parent, text="Пассажиры (на каждого будет создан отдельный индивидуальный пропуск). Необязательно.",
                  font=("", 10, "bold")).grid(row=0, column=0, sticky="w", **pad)
        self.persons_tbl = ttk.Frame(parent)
        self.persons_tbl.grid(row=1, column=0, sticky="nsew", **pad)
        parent.rowconfigure(1, weight=1)
        for c, t in enumerate(["Фамилия", "Имя", "Отчество", "Дата рожд."]):
            ttk.Label(self.persons_tbl, text=t, font=("", 9, "bold")).grid(row=0, column=c, sticky="w", padx=4)
        self.person_rows = []
        btns = ttk.Frame(parent)
        btns.grid(row=2, column=0, sticky="w", **pad)
        ttk.Button(btns, text="+ Добавить пассажира", command=self.add_person).pack(side="left")
        ttk.Button(btns, text="- Убрать последнего", command=self.remove_person).pack(side="left", padx=6)

    def add_person(self):
        n = len(self.person_rows) + 1
        frame = ttk.Frame(self.persons_tbl)
        frame.grid(row=n, column=0, sticky="ew", pady=1)
        vars = [tk.StringVar() for _ in range(4)]
        for c, v in enumerate(vars):
            if c == 3:
                # дата рождения — поле с маской и календарём
                ent = ttk.Entry(frame, textvariable=v, width=16)
                ent.grid(row=0, column=c, padx=4)
                vcmd = (self.register(lambda P: all(ch.isdigit() or ch == "." for ch in P) and P.count(".") <= 2), "%P")
                ent.configure(validate="key", validatecommand=vcmd)
                ent.bind("<KeyRelease>", lambda e, var=v, ent=ent: self._apply_date_mask(var, ent))
                btn = ttk.Button(frame, text="📅", width=3, command=lambda var=v, ent=ent: self._pick_date(var, ent))
                btn.grid(row=0, column=c+1, padx=(0, 4))
            else:
                ttk.Entry(frame, textvariable=v, width=16).grid(row=0, column=c, padx=4)
        self.person_rows.append((frame, vars))

    def remove_person(self):
        if not self.person_rows:
            return
        frame = self.person_rows.pop()[0]
        frame.destroy()

    def _build_tab3(self, parent, pad):
        parent.columnconfigure(1, weight=1)
        r = 0
        ttk.Label(parent, text="Автомобиль (заполните, если въезд на автомобиле)", font=("", 10, "bold")).grid(row=r, column=0, columnspan=3, sticky="w", **pad); r += 1
        ttk.Checkbutton(parent, text="Въезд на автомобиле", variable=self.var_car_on).grid(row=r, column=0, columnspan=3, sticky="w", **pad); r += 1
        self._entry(parent, "Марка-модель:", self.var_car_make, r, 0, **pad); r += 1
        self._entry(parent, "Регистрационный знак:", self.var_car_num, r, 0, **pad); r += 1

        ttk.Separator(parent, orient="horizontal").grid(row=r, column=0, columnspan=3, sticky="ew", **pad); r += 1
        ttk.Label(parent, text="Кому на подписание:", font=("", 10, "bold")).grid(row=r, column=0, sticky="w", **pad)
        ttk.Combobox(parent, textvariable=self.var_issued, width=50,
                     values=generator.SIGNERS, state="readonly").grid(row=r, column=1, columnspan=2, sticky="w", **pad)
        r += 1

        ttk.Separator(parent, orient="horizontal").grid(row=r, column=0, columnspan=3, sticky="ew", **pad); r += 1
        ttk.Label(parent, text="Папка для сохранения:", font=("", 10, "bold")).grid(row=r, column=0, columnspan=3, sticky="w", **pad); r += 1
        ttk.Label(parent, textvariable=self.var_output, foreground="#1a4").grid(row=r, column=0, columnspan=3, sticky="w", **pad); r += 1

    # ------------------------------------------------------------- действия
    def _selected_districts(self):
        return [d for d, v in self.district_vars.items() if v.get()]

    def _refresh_objects(self):
        sel = self._selected_districts()
        objs = generator.filter_objects_by_districts(self.reference, sel)
        self.obj_check.set_items([o["object"] for o in objs])

    def _maybe_autofill(self):
        """Автоподстановка данных из базы знаний по ФИО (только пустые поля)."""
        fio = " ".join(x for x in [self.var_lname.get().strip(), self.var_fname.get().strip(),
                                   self.var_mname.get().strip()] if x)
        rec = generator.db_find(fio)
        if not rec:
            return
        # заполняем только незаполненные поля, чтобы не затирать ручной ввод
        if not self.var_birth.get().strip() and rec.get("birth_date"):
            self.var_birth.set(rec["birth_date"])
        if not self.var_idnum.get().strip() and rec.get("id_number"):
            self.var_idnum.set(rec["id_number"])
        if not self.var_car_make.get().strip() and rec.get("car_make"):
            self.var_car_make.set(rec["car_make"])
            self.var_car_on.set(True)
        if not self.var_car_num.get().strip() and rec.get("car_number"):
            self.var_car_num.set(rec["car_number"])
        # районы и объекты — только если ещё ничего не выбрано
        if not self._selected_districts() and rec.get("districts"):
            for d in rec["districts"]:
                if d in self.district_vars:
                    self.district_vars[d].set(True)
            self._refresh_objects()
            saved = [o.strip() for o in str(rec.get("objects", "")).split(";") if o.strip()]
            for o in saved:
                for it, var in self.obj_check.items:
                    if o in it or it in o:
                        var.set(True)
        self.status_bar.config(text="База знаний: подставлены данные для «%s»" % fio)

    # ------------------------------------------------- поиск по базе
    def _refresh_db_list(self):
        """Список всех ФИО из базы (для выпадающего списка)."""
        vals = []
        for rec in generator._db_load():
            fio = " ".join(x for x in [rec.get("last_name", ""), rec.get("first_name", ""),
                                       rec.get("middle_name", "")] if x)
            if fio:
                vals.append(fio)
        self.db_cb["values"] = vals

    def _db_search_typed(self, event):
        """Живой поиск: фильтрует список по мере ввода и открывает его."""
        if event.keysym in ("Up", "Down", "Return", "Escape"):
            return
        q = self.db_cb.get().strip().lower()
        vals = []
        for rec in generator._db_load():
            fio = " ".join(x for x in [rec.get("last_name", ""), rec.get("first_name", ""),
                                       rec.get("middle_name", "")] if x)
            if fio and (not q or q in fio.lower()):
                vals.append(fio)
        self.db_cb["values"] = vals
        if vals:
            self.db_cb.event_generate("<Down>")

    def _on_pick_person(self, event=None):
        """Выбрали ФИО из списка — подставляем все данные этой записи."""
        fio = self.db_cb.get().strip()
        rec = generator.db_find(fio)
        if not rec:
            return
        self.var_birth.set(rec.get("birth_date", ""))
        self.var_idnum.set(rec.get("id_number", ""))
        car_on = bool(rec.get("car_make") or rec.get("car_number"))
        self.var_car_on.set(car_on)
        self.var_car_make.set(rec.get("car_make", ""))
        self.var_car_num.set(rec.get("car_number", ""))
        # районы
        for d, v in self.district_vars.items():
            v.set(False)
        for d in rec.get("districts", []):
            if d in self.district_vars:
                self.district_vars[d].set(True)
        self._refresh_objects()
        # объекты прошлой заявки
        saved = [o.strip() for o in str(rec.get("objects", "")).split(";") if o.strip()]
        for o in saved:
            for it, var in self.obj_check.items:
                if o in it or it in o:
                    var.set(True)
        self.status_bar.config(text="База знаний: загружена запись «%s»" % fio)

    def clear_form(self):
        """Очистка всей формы."""
        for v in (self.var_lname, self.var_fname, self.var_mname, self.var_birth,
                  self.var_idnum, self.var_custom_objects, self.var_date_from,
                  self.var_date_to, self.var_car_make, self.var_car_num, self.db_cb):
            v.set("")
        self.var_goal.set(generator.GOAL_BLAGOUSTROISTVO)
        self.var_app_date.set(generator._today_dmy())
        self.var_issued.set(generator.SIGNERS[0])
        self.var_car_on.set(False)
        for d, v in self.district_vars.items():
            v.set(False)
        self.obj_check.set_items([])
        while self.person_rows:
            self.person_rows.pop()[0].destroy()
        self.status_bar.config(text="Форма очищена.")

    def choose_output(self):
        d = filedialog.askdirectory(title="Выберите папку для сохранения документов")
        if d:
            self.var_output.set(d)

    def open_output(self):
        d = self.var_output.get()
        os.makedirs(d, exist_ok=True)
        os.startfile(d)

    def _collect_persons(self):
        persons = []
        for frame, vars in self.person_rows:
            vals = [v.get().strip() for v in vars]
            if any(vals):
                persons.append({"last_name": vals[0], "first_name": vals[1],
                                "middle_name": vals[2], "birth_date": vals[3]})
        return persons

    def _selected_objects(self):
        checked = self.obj_check.get_checked()
        return [{"object": it, "district": self.obj_to_district.get(it, "")} for it in checked]

    def generate(self):
        try:
            lname = self.var_lname.get().strip()
            fname = self.var_fname.get().strip()
            if not lname and not fname:
                messagebox.showwarning(APP_NAME, "Укажите фамилию и имя заявителя.")
                return

            districts = self._selected_districts()
            goal = self.var_goal.get().strip() or generator.GOAL_BLAGOUSTROISTVO
            sel_objects = self._selected_objects()
            custom_objs = self.var_custom_objects.get().strip()

            if goal == generator.GOAL_BLAGOUSTROISTVO:
                if not sel_objects and not custom_objs:
                    messagebox.showwarning(APP_NAME, "Выберите хотя бы одно кладбище из списка\n(предварительно отметьте район(ы)).")
                    return
                if not districts and not sel_objects:
                    messagebox.showwarning(APP_NAME, "Отметьте хотя бы один район.")
                    return
            elif goal == generator.GOAL_CUSTOM:
                if not custom_objs:
                    messagebox.showwarning(APP_NAME, "При цели «свой вариант» укажите место визита\nв поле «Произвольный объект».")
                    return
                goal = custom_objs  # цель = свой вариант (текст)
            else:
                # произвольная цель, введённая вручную
                pass

            objects = generator.join_objects(sel_objects, custom_objs)
            car = self.var_car_on.get()

            data = {
                "last_name": lname, "first_name": fname, "middle_name": self.var_mname.get().strip(),
                "birth_date": self.var_birth.get().strip(), "id_number": self.var_idnum.get().strip(),
                "districts": districts, "objects": objects,
                "goal": goal,
                "date_from": self.var_date_from.get().strip(), "date_to": self.var_date_to.get().strip(),
                "app_date": self.var_app_date.get().strip() or generator._today_dmy(),
                "car_make": self.var_car_make.get().strip() if car else "",
                "car_number": self.var_car_num.get().strip() if car else "",
                "issued_by": self.var_issued.get().strip(),
                "persons": self._collect_persons(),
            }

            outdir = self.var_output.get().strip() or self._default_output()
            files = generator.generate_all(data, outdir)
            msg = "Готово! Создано документов: %d\n\n%s" % (len(files), "\n".join(os.path.basename(f) for f in files))
            if messagebox.askyesno(APP_NAME, msg + "\n\nОткрыть папку с документами?"):
                os.makedirs(os.path.dirname(files[0]), exist_ok=True)
                os.startfile(os.path.dirname(files[0]))
            self.status_bar.config(text="Создано файлов: %d → %s" % (len(files), os.path.dirname(files[0])))
        except Exception as ex:
            messagebox.showerror(APP_NAME, "Ошибка генерации:\n%s" % ex)

    def _build_about_tab(self, parent):
        """Вкладка «О программе»."""
        txt = (
            "Генератор_пропусков_14.3\n"
            "Версия 0.1.6\n\n"
            "Создатель: Соломейчук Алексей\n"
            "Creator: Salamiaichuk Aliaksei\n\n"
            "Email: al.vl.solo@yandex.by\n\n"
            "© 2025 Соломейчук Алексей / Salamiaichuk Aliaksei\n"
            "Все права защищены.\n\n"
            "Программа для генерации документов пропусков на въезд\n"
            "в зону эвакуации/отчуждения (п. 14.3 перечня\n"
            "административных процедур). Работает полностью локально,\n"
            "без интернета и внешних сервисов."
        )
        lbl = ttk.Label(parent, text=txt, justify="left", anchor="nw", wraplength=700)
        lbl.pack(fill="both", expand=True, padx=16, pady=16)

    # ------------------------------------------------------------------
    # Экспорт БД
    # ------------------------------------------------------------------
    def export_db(self):
        export_db_dialog(self, generator._db_path(),
                         "permit_history_143_XXXX.json")

    # ------------------------------------------------------------------
    # Проверка обновлений
    # ------------------------------------------------------------------
    def check_updates(self):
        dlg = UpdateDialog(self, APP_VERSION, "Генератор_пропусков_14.3.exe")
        self.wait_window(dlg)

def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()