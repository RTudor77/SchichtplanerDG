import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import pandas as pd
from datetime import datetime, timedelta
import os
from typing import Dict, List, Optional, Tuple

# Konstanten
DAYS_IN_PLANNING = 12
DAYS_PER_WEEK = 6
WEEKDAY_NAMES = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"]
WINDOW_GEOMETRY = "1400x900"
MIN_WINDOW_SIZE = (1200, 800)
CONFIG_FILE = "shift_config.json"

# Excel Farbpalette
COLOR_GREEN = "A9D18E"
COLOR_YELLOW = "FFD966"
COLOR_GREY = "BFBFBF"


class ShiftPlanner:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Notdienst Schichtplaner v1.2 (Optimiert)")
        self.root.geometry(WINDOW_GEOMETRY)
        self.root.minsize(*MIN_WINDOW_SIZE)

        # Konfigurationsdatei
        self.config_file = CONFIG_FILE

        # Standardkonfiguration
        self.config: Dict[str, List[str]] = {
            "pool_vm_alle": [],            # Vormittag - können alles
            "pool_vm_teilweise": [],       # Vormittag - können nicht alles (brauchen Support)
            "pool_vm_support": [],         # Vormittag - Support für Pool B
            "pool_nm_alle": [],            # Nachmittag - können alles
            "pool_freitag_abwesend": [],   # Freitags nicht verfügbar
            "feiertage": []                # [{datum: "25.12", name: "1. Weihnachtstag", mitarbeiter: "XX"}]
        }

        # Cache für Performance
        self._all_employees_cache: Optional[List[str]] = None
        self._cache_dirty = True

        # Interne Zustandsverwaltung
        self.planning_result: List[Dict[str, str]] = []
        self.absences: Dict[int, List[str]] = {}

        self.load_config()
        self.create_gui()

        # Automatisches Laden der Mitarbeiterliste nach GUI-Erstellung
        self.root.after(100, self._auto_update_employee_list)

    # -------------------- Konfiguration laden/speichern --------------------

    def load_config(self) -> None:
        """Lädt Konfiguration aus JSON-Datei"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)
                    # Sicherstellen, dass alle benötigten Pools existieren
                    for key in self.config.keys():
                        if key not in loaded_config:
                            loaded_config[key] = []
                    self.config = loaded_config
                    self._cache_dirty = True
        except json.JSONDecodeError as e:
            messagebox.showerror("Fehler", f"Ungültige JSON-Datei: {e}")
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Laden der Konfiguration: {e}")

    def save_config(self) -> None:
        """Speichert Konfiguration in JSON-Datei"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
            self._cache_dirty = True
            messagebox.showinfo("Erfolg", "Konfiguration gespeichert!")
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Speichern: {e}")

    def _get_all_employees(self) -> List[str]:
        """Gibt gecachte Liste aller Mitarbeiter zurück (Performance-Optimierung)"""
        if self._cache_dirty or self._all_employees_cache is None:
            all_employees = set()
            for pool_key in ["pool_vm_alle", "pool_vm_teilweise", "pool_vm_support", "pool_nm_alle"]:
                all_employees.update(self.config[pool_key])
            self._all_employees_cache = sorted(list(all_employees))
            self._cache_dirty = False
        return self._all_employees_cache

    def _parse_employee_input(self, input_str: str) -> List[str]:
        """Parst Mitarbeiter-Eingabe (kommagetrennt) mit Validierung"""
        return [x.strip() for x in input_str.split(",") if x.strip()]

    def _auto_update_employee_list(self) -> None:
        """Automatisches Laden der Mitarbeiterliste (ohne Messagebox)"""
        try:
            all_employees = self._get_all_employees()
            if hasattr(self, 'employee_combo'):
                self.employee_combo['values'] = all_employees
            if hasattr(self, 'holiday_employee_combo'):
                self.holiday_employee_combo['values'] = all_employees
        except Exception:
            pass  # Stilles Fehlschlagen beim automatischen Update

    # -------------------- GUI --------------------

    def create_gui(self):
        """Erstellt die Benutzeroberfläche"""
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Tab 1: Pool-Konfiguration
        self.create_pool_config_tab(notebook)
        # Tab 2: Schichtplanung
        self.create_shift_planning_tab(notebook)

    def _create_pool_entry(self, parent: ttk.Frame, row: int, label_text: str,
                          config_key: str, **label_kwargs) -> tk.Entry:
        """Hilfsmethode zum Erstellen eines Pool-Entry-Feldes (DRY)"""
        ttk.Label(parent, text=label_text, **label_kwargs).grid(
            row=row, column=0, sticky="w", padx=5, pady=5
        )
        entry = tk.Entry(parent, width=60)
        entry.grid(row=row, column=1, padx=5, pady=5)
        entry.insert(0, ",".join(self.config[config_key]))
        return entry

    def create_pool_config_tab(self, notebook: ttk.Notebook) -> None:
        """Erstellt Tab für Pool-Konfiguration mit Feiertagen"""
        pool_frame = ttk.Frame(notebook, padding=10)
        notebook.add(pool_frame, text="Pool-Konfiguration")

        # Hauptcontainer: Links Pools, Rechts Feiertage
        main_container = ttk.PanedWindow(pool_frame, orient="horizontal")
        main_container.pack(fill="both", expand=True)

        # === LINKER BEREICH: Pool-Konfiguration ===
        left_frame = ttk.LabelFrame(main_container, text=" Mitarbeiter-Pools ", padding=10)

        # Pools erstellen (DRY mit Hilfsmethode)
        self.pool_vm_alle_entry = self._create_pool_entry(
            left_frame, 0, "Pool A - Vormittag (können alles):", "pool_vm_alle"
        )
        self.pool_vm_teilweise_entry = self._create_pool_entry(
            left_frame, 1, "Pool B - Vormittag (brauchen Unterstützung):", "pool_vm_teilweise"
        )
        self.pool_vm_support_entry = self._create_pool_entry(
            left_frame, 2, "Pool C - Support für Pool B:", "pool_vm_support"
        )
        self.pool_nm_alle_entry = self._create_pool_entry(
            left_frame, 3, "Pool D - Nachmittag (können alles):", "pool_nm_alle"
        )
        self.pool_freitag_abwesend_entry = self._create_pool_entry(
            left_frame, 4, "Pool E - Freitags NICHT verfügbar:", "pool_freitag_abwesend",
            font=('TkDefaultFont', 9, 'bold'), foreground='orange'
        )

        # Speichern Button
        ttk.Button(left_frame, text="Pools speichern", command=self.save_pools).grid(
            row=5, column=1, pady=15, sticky="e"
        )

        info_text = """Hinweise:
• Mitarbeiter mit Komma trennen (z.B. RR,AN,MH)
• Reihenfolge wird bei der Planung berücksichtigt"""
        ttk.Label(left_frame, text=info_text, justify="left", font=('TkDefaultFont', 8),
                 foreground='gray').grid(row=6, column=0, columnspan=2, padx=5, pady=5, sticky="w")

        main_container.add(left_frame, weight=2)

        # === RECHTER BEREICH: Feiertage ===
        right_frame = ttk.LabelFrame(main_container, text=" Feiertage (jährlich) ", padding=10)

        # Eingabefelder für neuen Feiertag
        input_frame = ttk.Frame(right_frame)
        input_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(input_frame, text="Datum (TT.MM):").grid(row=0, column=0, sticky="w", padx=2)
        self.holiday_date_var = tk.StringVar()
        self.holiday_date_entry = ttk.Entry(input_frame, textvariable=self.holiday_date_var, width=8)
        self.holiday_date_entry.grid(row=0, column=1, padx=5)

        ttk.Label(input_frame, text="Name:").grid(row=0, column=2, sticky="w", padx=2)
        self.holiday_name_var = tk.StringVar()
        self.holiday_name_entry = ttk.Entry(input_frame, textvariable=self.holiday_name_var, width=20)
        self.holiday_name_entry.grid(row=0, column=3, padx=5)

        input_frame2 = ttk.Frame(right_frame)
        input_frame2.pack(fill="x", pady=(0, 10))

        ttk.Label(input_frame2, text="Notdienst MA:").grid(row=0, column=0, sticky="w", padx=2)
        self.holiday_employee_var = tk.StringVar()
        self.holiday_employee_combo = ttk.Combobox(input_frame2, textvariable=self.holiday_employee_var, width=8)
        self.holiday_employee_combo.grid(row=0, column=1, padx=5)

        # Buttons
        btn_frame = ttk.Frame(right_frame)
        btn_frame.pack(fill="x", pady=(0, 10))
        ttk.Button(btn_frame, text="+ Hinzufügen", command=self.add_holiday, width=12).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="- Entfernen", command=self.remove_holiday, width=12).pack(side="left", padx=2)

        # Feiertage-Liste
        holiday_columns = ("Datum", "Name", "MA")
        self.holiday_tree = ttk.Treeview(right_frame, columns=holiday_columns, show="headings", height=12)
        self.holiday_tree.heading("Datum", text="Datum")
        self.holiday_tree.heading("Name", text="Name")
        self.holiday_tree.heading("MA", text="Notdienst")
        self.holiday_tree.column("Datum", width=70, anchor="center")
        self.holiday_tree.column("Name", width=150, anchor="w")
        self.holiday_tree.column("MA", width=70, anchor="center")

        scrollbar_holiday = ttk.Scrollbar(right_frame, orient="vertical", command=self.holiday_tree.yview)
        self.holiday_tree.configure(yscrollcommand=scrollbar_holiday.set)

        self.holiday_tree.pack(side="left", fill="both", expand=True)
        scrollbar_holiday.pack(side="right", fill="y")

        # Feiertage aus Config laden
        self.update_holiday_display()

        main_container.add(right_frame, weight=1)

        info_label = ttk.Label(pool_frame, text="Feiertage werden dauerhaft gespeichert. "
                              "Der Notdienst-MA wird in der Planungswoche von regulären Schichten ausgeschlossen.",
                              font=('TkDefaultFont', 8), foreground='gray')
        info_label.pack(fill="x", pady=(10, 0))

    def create_shift_planning_tab(self, notebook: ttk.Notebook) -> None:
        """Erstellt Tab für Schichtplanung - MODERNES LAYOUT"""
        planning_frame = ttk.Frame(notebook, padding=10)
        notebook.add(planning_frame, text="Schichtplanung")

        # Modernes Styling konfigurieren
        self._configure_modern_styles()

        # === OBERER BEREICH: Grundeinstellungen (kompakt) ===
        settings_frame = ttk.LabelFrame(planning_frame, text=" Grundeinstellungen ", padding=10)
        settings_frame.pack(fill="x", pady=(0, 10))

        # Eingabefelder in einer Zeile
        input_frame = ttk.Frame(settings_frame)
        input_frame.pack(fill="x")

        # Startdatum
        ttk.Label(input_frame, text="Startdatum:").grid(row=0, column=0, sticky="w", padx=(0, 5))
        self.start_date_entry = ttk.Entry(input_frame, width=14)
        self.start_date_entry.grid(row=0, column=1, padx=(0, 15))

        # 1. Tag VM
        ttk.Label(input_frame, text="1. Tag VM:").grid(row=0, column=2, sticky="w", padx=(0, 5))
        self.first_vm_entry = ttk.Entry(input_frame, width=8)
        self.first_vm_entry.grid(row=0, column=3, padx=(0, 15))

        # 1. Tag NM
        ttk.Label(input_frame, text="1. Tag NM:").grid(row=0, column=4, sticky="w", padx=(0, 5))
        self.first_nm_entry = ttk.Entry(input_frame, width=8)
        self.first_nm_entry.grid(row=0, column=5, padx=(0, 15))

        # 1. Tag Support
        ttk.Label(input_frame, text="Support:").grid(row=0, column=6, sticky="w", padx=(0, 5))
        self.first_support_entry = ttk.Entry(input_frame, width=8)
        self.first_support_entry.grid(row=0, column=7, padx=(0, 20))

        # Buttons
        ttk.Button(input_frame, text="Planung erstellen", command=self.create_planning,
                  style="Accent.TButton").grid(row=0, column=8, padx=5)
        ttk.Button(input_frame, text="Excel Export", command=self.export_excel).grid(row=0, column=9, padx=5)

        # === HAUPTBEREICH: Planungsergebnis (links) + Abwesenheiten (rechts) ===
        main_paned = ttk.PanedWindow(planning_frame, orient="horizontal")
        main_paned.pack(fill="both", expand=True)

        # === LINKER BEREICH: Planungsergebnis (größer) ===
        result_frame = ttk.LabelFrame(main_paned, text=" Planungsergebnis ", padding=10)

        columns = ("Datum", "Tag", "VM", "NM", "Support")
        self.result_tree = ttk.Treeview(result_frame, columns=columns, show="headings",
                                        height=20, style="Modern.Treeview")

        # Spalten-Konfiguration (kompakter)
        column_config = {
            "Datum": (85, "center"),
            "Tag": (80, "center"),
            "VM": (60, "center"),
            "NM": (60, "center"),
            "Support": (60, "center")
        }
        for col_name, (width, anchor) in column_config.items():
            self.result_tree.heading(col_name, text=col_name)
            self.result_tree.column(col_name, width=width, minwidth=50, anchor=anchor)

        scrollbar_result = ttk.Scrollbar(result_frame, orient="vertical", command=self.result_tree.yview)
        self.result_tree.configure(yscrollcommand=scrollbar_result.set)

        self.result_tree.pack(side="left", fill="both", expand=True)
        scrollbar_result.pack(side="right", fill="y")

        main_paned.add(result_frame, weight=3)

        # === RECHTER BEREICH: Abwesenheiten ===
        absence_frame = ttk.LabelFrame(main_paned, text=" Abwesenheiten (temporär) ", padding=10)

        input_absence_frame = ttk.Frame(absence_frame)
        input_absence_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(input_absence_frame, text="MA:").grid(row=0, column=0, sticky="w")
        self.employee_var = tk.StringVar()
        self.employee_combo = ttk.Combobox(input_absence_frame, textvariable=self.employee_var, width=10)
        self.employee_combo.grid(row=0, column=1, padx=5)

        ttk.Label(input_absence_frame, text="Tag:").grid(row=0, column=2, sticky="w")
        self.day_var = tk.StringVar()
        self.day_combo = ttk.Combobox(input_absence_frame, textvariable=self.day_var, width=8,
                                      values=[f"Tag {i + 1}" for i in range(DAYS_IN_PLANNING)])
        self.day_combo.grid(row=0, column=3, padx=5)

        btn_frame = ttk.Frame(absence_frame)
        btn_frame.pack(fill="x", pady=(0, 10))
        ttk.Button(btn_frame, text="+ Hinzufügen", command=self.add_absence, width=12).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="- Entfernen", command=self.remove_absence, width=12).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="Laden", command=self.update_employee_list, width=8).pack(side="left", padx=2)

        absence_columns = ("Tag", "MA")
        self.absence_tree = ttk.Treeview(absence_frame, columns=absence_columns, show="headings",
                                         height=16, style="Modern.Treeview")
        self.absence_tree.heading("Tag", text="Tag")
        self.absence_tree.heading("MA", text="MA")
        self.absence_tree.column("Tag", width=70, anchor="center")
        self.absence_tree.column("MA", width=70, anchor="center")

        scrollbar_absence = ttk.Scrollbar(absence_frame, orient="vertical", command=self.absence_tree.yview)
        self.absence_tree.configure(yscrollcommand=scrollbar_absence.set)
        self.absence_tree.pack(side="left", fill="both", expand=True)
        scrollbar_absence.pack(side="right", fill="y")

        main_paned.add(absence_frame, weight=1)

    def _configure_modern_styles(self) -> None:
        """Konfiguriert moderne ttk Styles"""
        style = ttk.Style()

        # Versuche ein moderneres Theme zu verwenden
        available_themes = style.theme_names()
        if 'clam' in available_themes:
            style.theme_use('clam')
        elif 'vista' in available_themes:
            style.theme_use('vista')

        # Moderne Treeview-Styles
        style.configure("Modern.Treeview",
                       rowheight=25,
                       font=('Segoe UI', 9))
        style.configure("Modern.Treeview.Heading",
                       font=('Segoe UI', 9, 'bold'),
                       padding=5)

        # Accent Button Style
        style.configure("Accent.TButton",
                       font=('Segoe UI', 9, 'bold'))

        # LabelFrame Style
        style.configure("TLabelframe.Label",
                       font=('Segoe UI', 10, 'bold'))

    # -------------------- Abwesenheiten --------------------

    def update_employee_list(self) -> None:
        """Aktualisiert die Mitarbeiterliste in der Combobox"""
        self._cache_dirty = True
        all_employees = self._get_all_employees()
        self.employee_combo['values'] = all_employees
        messagebox.showinfo("Info", f"Mitarbeiterliste aktualisiert: {len(all_employees)} Mitarbeiter gefunden")

    def add_absence(self) -> None:
        """Fügt eine Abwesenheit hinzu (unterstützt mehrere, komma-getrennte Kürzel)."""
        employee = self.employee_var.get().strip()
        day_str = self.day_var.get().strip()

        if not employee or not day_str:
            messagebox.showwarning("Warnung", "Bitte Mitarbeiter und Tag auswählen!")
            return

        try:
            day_nr = int(day_str.split()[1]) - 1  # "Tag 1" -> 0
            if not (0 <= day_nr < DAYS_IN_PLANNING):
                raise ValueError("Tag außerhalb des gültigen Bereichs")
        except (ValueError, IndexError) as e:
            messagebox.showerror("Fehler", f"Ungültiger Tag: {e}")
            return

        employees = self._parse_employee_input(employee)
        if not employees:
            messagebox.showwarning("Warnung", "Keine gültigen Mitarbeiter eingegeben!")
            return

        if day_nr not in self.absences:
            self.absences[day_nr] = []

        # Nur neue Mitarbeiter hinzufügen
        added_count = 0
        for emp in employees:
            if emp not in self.absences[day_nr]:
                self.absences[day_nr].append(emp)
                added_count += 1

        self.update_absence_display()

        # Feedback und automatischer Sprung zum nächsten Tag
        if added_count > 0:
            current_day = day_nr + 1
            if current_day < DAYS_IN_PLANNING:
                self.day_var.set(f"Tag {current_day + 1}")
            # MA bleibt stehen für schnellere Mehrfacheingabe

    def remove_absence(self) -> None:
        """Entfernt eine Abwesenheit"""
        selected = self.absence_tree.selection()
        if not selected:
            messagebox.showwarning("Warnung", "Bitte einen Eintrag auswählen!")
            return

        removed_count = 0
        for item in selected:
            values = self.absence_tree.item(item, 'values')
            day_str, employee = values[0], values[1]
            try:
                day_nr = int(day_str.split()[1]) - 1
                if day_nr in self.absences and employee in self.absences[day_nr]:
                    self.absences[day_nr].remove(employee)
                    if not self.absences[day_nr]:
                        del self.absences[day_nr]
                    removed_count += 1
            except (ValueError, IndexError):
                continue

        self.update_absence_display()
        if removed_count > 0:
            messagebox.showinfo("Erfolg", f"{removed_count} Abwesenheit(en) entfernt")

    def update_absence_display(self) -> None:
        """Aktualisiert die Anzeige der Abwesenheiten"""
        for item in self.absence_tree.get_children():
            self.absence_tree.delete(item)

        for day_nr in sorted(self.absences.keys()):
            for employee in sorted(self.absences[day_nr]):
                self.absence_tree.insert("", tk.END, values=(f"Tag {day_nr + 1}", employee))

    # -------------------- Feiertage (dauerhaft gespeichert) --------------------

    def add_holiday(self) -> None:
        """Fügt einen Feiertag hinzu und speichert in Config"""
        date_str = self.holiday_date_var.get().strip()
        name = self.holiday_name_var.get().strip()
        employee = self.holiday_employee_var.get().strip()

        if not date_str or not name:
            messagebox.showwarning("Warnung", "Bitte Datum und Name eingeben!")
            return

        # Datum validieren (TT.MM Format)
        try:
            parts = date_str.split(".")
            if len(parts) != 2:
                raise ValueError("Format muss TT.MM sein")
            day, month = int(parts[0]), int(parts[1])
            if not (1 <= day <= 31 and 1 <= month <= 12):
                raise ValueError("Ungültiges Datum")
            date_str = f"{day:02d}.{month:02d}"  # Normalisieren
        except (ValueError, IndexError) as e:
            messagebox.showerror("Fehler", f"Ungültiges Datum: {e}\nFormat: TT.MM (z.B. 25.12)")
            return

        # Prüfen ob Datum bereits existiert
        for holiday in self.config["feiertage"]:
            if holiday["datum"] == date_str:
                # Update statt neu hinzufügen
                holiday["name"] = name
                holiday["mitarbeiter"] = employee
                self.update_holiday_display()
                self.save_config()
                self._clear_holiday_inputs()
                return

        # Neuen Feiertag hinzufügen
        self.config["feiertage"].append({
            "datum": date_str,
            "name": name,
            "mitarbeiter": employee
        })

        # Nach Datum sortieren
        self.config["feiertage"].sort(key=lambda x: (int(x["datum"].split(".")[1]), int(x["datum"].split(".")[0])))

        self.update_holiday_display()
        self.save_config()
        self._clear_holiday_inputs()

    def _clear_holiday_inputs(self) -> None:
        """Leert die Feiertags-Eingabefelder"""
        self.holiday_date_var.set("")
        self.holiday_name_var.set("")
        self.holiday_employee_var.set("")

    def remove_holiday(self) -> None:
        """Entfernt einen Feiertag und speichert Config"""
        selected = self.holiday_tree.selection()
        if not selected:
            messagebox.showwarning("Warnung", "Bitte einen Eintrag auswählen!")
            return

        for item in selected:
            values = self.holiday_tree.item(item, 'values')
            date_str = values[0]
            # Feiertag aus Liste entfernen
            self.config["feiertage"] = [
                h for h in self.config["feiertage"] if h["datum"] != date_str
            ]

        self.update_holiday_display()
        self.save_config()

    def update_holiday_display(self) -> None:
        """Aktualisiert die Anzeige der Feiertage aus Config"""
        for item in self.holiday_tree.get_children():
            self.holiday_tree.delete(item)

        for holiday in self.config.get("feiertage", []):
            self.holiday_tree.insert("", tk.END, values=(
                holiday["datum"],
                holiday["name"],
                holiday.get("mitarbeiter", "")
            ))

    def _get_holiday_employees_for_week(self, tag_nr: int, start_date: datetime) -> List[str]:
        """Gibt Mitarbeiter zurück, die in der Woche des Tags Feiertags-Notdienst haben"""
        # Woche berechnen (0-5 = Woche 1, 6-11 = Woche 2)
        week_start = (tag_nr // DAYS_PER_WEEK) * DAYS_PER_WEEK
        week_end = week_start + DAYS_PER_WEEK

        holiday_employees = []

        # Alle Tage der Woche durchgehen
        for day_in_week in range(week_start, week_end):
            week_number = day_in_week // DAYS_PER_WEEK
            current_date = start_date + timedelta(days=day_in_week + week_number)
            date_str = current_date.strftime("%d.%m")

            # Prüfen ob dieser Tag ein Feiertag ist
            for holiday in self.config.get("feiertage", []):
                if holiday["datum"] == date_str and holiday.get("mitarbeiter"):
                    holiday_employees.append(holiday["mitarbeiter"])

        return holiday_employees

    def save_pools(self) -> None:
        """Speichert die Pool-Konfiguration (optimiert mit DRY)"""
        try:
            pool_entries = {
                "pool_vm_alle": self.pool_vm_alle_entry,
                "pool_vm_teilweise": self.pool_vm_teilweise_entry,
                "pool_vm_support": self.pool_vm_support_entry,
                "pool_nm_alle": self.pool_nm_alle_entry,
                "pool_freitag_abwesend": self.pool_freitag_abwesend_entry
            }

            for pool_key, entry_widget in pool_entries.items():
                self.config[pool_key] = self._parse_employee_input(entry_widget.get())

            self.save_config()
            self._auto_update_employee_list()  # Automatisch Mitarbeiterliste aktualisieren
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Speichern der Pools: {e}")

    def parse_absent_employees(self) -> Dict[int, List[str]]:
        """Gibt die Abwesenheitsstruktur zurück"""
        return self.absences

    # -------------------- Planung --------------------

    def _validate_planning_inputs(self) -> Optional[Tuple[datetime, str, str, str]]:
        """Validiert Eingaben für Planung. Gibt (start_date, first_vm, first_nm, first_support) oder None zurück"""
        start_date_str = self.start_date_entry.get().strip()
        if not start_date_str:
            messagebox.showerror("Fehler", "Bitte Startdatum eingeben!")
            return None

        try:
            start_date = datetime.strptime(start_date_str, "%d.%m.%Y")
        except ValueError:
            messagebox.showerror("Fehler", "Ungültiges Datumsformat! Bitte TT.MM.YYYY verwenden.")
            return None

        if start_date.weekday() != 0:  # Montag
            messagebox.showerror("Fehler", "Startdatum muss ein Montag sein!")
            return None

        first_vm = self.first_vm_entry.get().strip()
        first_nm = self.first_nm_entry.get().strip()
        first_support = self.first_support_entry.get().strip()

        if not all([first_vm, first_nm]):
            messagebox.showerror("Fehler", "Bitte mindestens Vormittag und Nachmittag für den ersten Tag ausfüllen!")
            return None

        return start_date, first_vm, first_nm, first_support

    def _initialize_pool_positions(self, first_vm: str, first_nm: str, first_support: str) -> Dict[str, int]:
        """Initialisiert Pool-Positionen basierend auf ersten Mitarbeitern"""
        pool_positions = {"vm_alle": 0, "vm_teilweise": 0, "vm_support": 0, "nm_alle": 0}

        # Effizienter: nur einmal suchen
        pool_configs = [
            ("vm_alle", first_vm, self.config["pool_vm_alle"]),
            ("nm_alle", first_nm, self.config["pool_nm_alle"]),
            ("vm_support", first_support, self.config["pool_vm_support"]),
            ("vm_teilweise", first_vm, self.config["pool_vm_teilweise"])
        ]

        for pool_key, employee, pool_list in pool_configs:
            if employee and pool_list and employee in pool_list:
                idx = pool_list.index(employee)
                pool_positions[pool_key] = (idx + 1) % len(pool_list)

        return pool_positions

    def _get_absent_for_day(self, tag_nr: int, is_friday: bool, start_date: datetime) -> List[str]:
        """Gibt Liste der abwesenden Mitarbeiter für einen Tag zurück"""
        absent_today = self.absences.get(tag_nr, []).copy()

        if is_friday:
            # Freitags zusätzlich die generellen "Freitag nicht verfügbar"
            absent_today = list(set(absent_today + self.config["pool_freitag_abwesend"]))

        # Feiertags-Mitarbeiter für diese Woche ausschließen
        holiday_employees = self._get_holiday_employees_for_week(tag_nr, start_date)
        absent_today = list(set(absent_today + holiday_employees))

        return absent_today

    def _find_employee_from_pool(self, pool: List[str], start_pos: int,
                                 excluded: List[str], absent: List[str]) -> Optional[Tuple[str, int]]:
        """Findet nächsten verfügbaren Mitarbeiter aus Pool. Gibt (employee, new_position) zurück"""
        if not pool:
            return None

        pool_size = len(pool)
        for i in range(pool_size):
            candidate_idx = (start_pos + i) % pool_size
            candidate = pool[candidate_idx]
            if candidate not in absent and candidate not in excluded:
                return candidate, (candidate_idx + 1) % pool_size
        return None

    def create_planning(self) -> None:
        """Erstellt die Schichtplanung für 2 Wochen Mo-Sa (12 Tage) - OPTIMIERT"""
        try:
            # Input-Validierung
            validation_result = self._validate_planning_inputs()
            if not validation_result:
                return
            start_date, first_vm, first_nm, first_support = validation_result

            self.planning_result = []
            pool_positions = self._initialize_pool_positions(first_vm, first_nm, first_support)

            # 12 Tage: Mo-Sa, Mo-Sa (2 Wochen ohne Sonntag)
            for tag_nr in range(DAYS_IN_PLANNING):
                # Wochentag berechnen
                weekday_in_week = tag_nr % DAYS_PER_WEEK
                week_number = tag_nr // DAYS_PER_WEEK
                current_date = start_date + timedelta(days=tag_nr + week_number)

                is_friday = weekday_in_week == 4
                is_saturday = weekday_in_week == 5
                weekday_german = WEEKDAY_NAMES[weekday_in_week]

                absent_today = self._get_absent_for_day(tag_nr, is_friday, start_date)

                # Samstag: keine Schichten
                if is_saturday:
                    vm_employee = ""
                    nm_employee = ""
                    support_employee = ""
                elif tag_nr == 0:
                    # Erster Tag: vordefinierte Mitarbeiter
                    vm_employee = first_vm
                    nm_employee = first_nm
                    support_employee = first_support if first_support else ""
                else:
                    # Alle anderen Tage: automatische Planung
                    vm_employee, nm_employee, support_employee = self._plan_day(
                        tag_nr, absent_today, pool_positions
                    )

                self.planning_result.append({
                    "Datum": current_date.strftime("%d.%m.%Y"),
                    "Wochentag": weekday_german,
                    "Vormittag": vm_employee or "",
                    "Nachmittag": nm_employee or "",
                    "Support": support_employee or ""
                })

            self.display_results()
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler bei der Planung: {e}")

    def _plan_day(self, tag_nr: int, absent_today: List[str],
                  pool_positions: Dict[str, int]) -> Tuple[Optional[str], Optional[str], str]:
        """Plant einen einzelnen Tag. Gibt (vm_employee, nm_employee, support_employee) zurück"""
        yesterday_nm = self.planning_result[tag_nr - 1]["Nachmittag"] if tag_nr > 0 else None
        forbidden_vm = [yesterday_nm] if yesterday_nm else []
        used_today = []

        # Vormittag planen
        vm_result = self._find_employee_from_pool(
            self.config["pool_vm_alle"],
            pool_positions["vm_alle"],
            forbidden_vm,
            absent_today
        )

        vm_employee = None
        support_employee = ""

        if vm_result:
            vm_employee, pool_positions["vm_alle"] = vm_result
            used_today.append(vm_employee)

            # Support nötig?
            if vm_employee in self.config["pool_vm_teilweise"]:
                support_result = self._find_employee_from_pool(
                    self.config["pool_vm_support"],
                    pool_positions["vm_support"],
                    [vm_employee] + forbidden_vm,
                    absent_today
                )
                if support_result:
                    support_employee, pool_positions["vm_support"] = support_result
                    used_today.append(support_employee)

        # Nachmittag planen
        nm_result = self._find_employee_from_pool(
            self.config["pool_nm_alle"],
            pool_positions["nm_alle"],
            used_today,
            absent_today
        )

        nm_employee = None
        if nm_result:
            nm_employee, pool_positions["nm_alle"] = nm_result

        return vm_employee, nm_employee, support_employee

    # -------------------- Anzeige & Export --------------------

    def display_results(self) -> None:
        """Zeigt die Planungsergebnisse in der Treeview an (optimiert)"""
        # Alte Einträge löschen
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        # Neue Einträge hinzufügen (angepasst an neue kompakte Spalten)
        for row in self.planning_result:
            # Wochentag abkürzen für kompaktere Darstellung
            tag_kurz = row["Wochentag"][:2] if row["Wochentag"] else ""
            self.result_tree.insert(
                "", tk.END,
                values=(row["Datum"], tag_kurz, row["Vormittag"],
                       row["Nachmittag"], row["Support"])
            )

    def _prepare_export_data(self) -> pd.DataFrame:
        """Bereitet Daten für Export vor"""
        rows = []
        for r in self.planning_result:
            try:
                datum_dt = datetime.strptime(r["Datum"], "%d.%m.%Y")
            except Exception:
                datum_dt = r["Datum"]
            rows.append({
                "Datum": datum_dt,
                "Wochentag": r["Wochentag"],
                "Vormittag": r["Vormittag"],
                "Nachmittag": r["Nachmittag"],
                "Support": r["Support"]
            })
        return pd.DataFrame(rows, columns=["Datum", "Wochentag", "Vormittag", "Nachmittag", "Support"])

    def _export_csv(self, filename: str, df: pd.DataFrame) -> None:
        """Exportiert als CSV"""
        df_copy = df.copy()
        if not df_copy.empty and isinstance(df_copy.loc[0, "Datum"], datetime):
            df_copy["Datum"] = df_copy["Datum"].dt.strftime("%d.%m.%Y")
        df_copy.to_csv(filename, index=False, encoding="utf-8-sig")
        messagebox.showinfo("Erfolg", f"CSV-Datei erfolgreich gespeichert: {filename}")

    def export_excel(self) -> None:
        """Exportiert die Planung nach Excel im Layout:
        Spalten: Datum | Wochentag | Vormittag (grün) | Nachmittag (grün) | Support (gelb)
        Samstag wird aufgeführt (ohne Planung), nach jedem Samstag eine graue Trennzeile.
        """
        if not self.planning_result:
            messagebox.showwarning("Warnung", "Keine Planung vorhanden! Bitte erst Planung erstellen.")
            return

        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not filename:
            return

        df = self._prepare_export_data()

        # CSV-Fallback ohne Formatierung
        if filename.lower().endswith(".csv"):
            self._export_csv(filename, df)
            return

        # Excel mit Formatierung
        try:
            self._export_excel_formatted(filename, df)
        except ImportError:
            # Fallback zu CSV wenn openpyxl nicht verfügbar
            csv_filename = filename.rsplit(".", 1)[0] + ".csv"
            self._export_csv(csv_filename, df)
            messagebox.showinfo(
                "Info",
                "openpyxl ist nicht installiert. CSV gespeichert: "
                f"{csv_filename}\nFür formatiertes Excel bitte installieren: pip install openpyxl"
            )
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Export: {e}")

    def _export_excel_formatted(self, filename: str, df: pd.DataFrame) -> None:
        """Exportiert formatiertes Excel mit Farben und Trennzeilen"""
        from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
        from openpyxl.utils import get_column_letter

        # Farbpalette
        GREEN = PatternFill(start_color=COLOR_GREEN, end_color=COLOR_GREEN, fill_type="solid")
        YELLOW = PatternFill(start_color=COLOR_YELLOW, end_color=COLOR_YELLOW, fill_type="solid")
        GREY = PatternFill(start_color=COLOR_GREY, end_color=COLOR_GREY, fill_type="solid")

        thin = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin")
        )

        with pd.ExcelWriter(filename, engine="openpyxl", date_format="DD.MM.YYYY") as writer:
            df.to_excel(writer, sheet_name="Notdienst_Planung", index=False)
            ws = writer.sheets["Notdienst_Planung"]

            # Spaltenbreiten setzen
            column_widths = [12, 14, 14, 14, 14]  # A..E
            for idx, width in enumerate(column_widths, start=1):
                ws.column_dimensions[get_column_letter(idx)].width = width

            # Header formatieren
            for cell in ws[1]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center")
                cell.border = thin

            # Datenzellen formatieren
            max_row = ws.max_row
            max_col = ws.max_column

            # Farbzuordnung für Spalten
            column_fills = {3: GREEN, 4: GREEN, 5: YELLOW}

            for r in range(2, max_row + 1):
                for c in range(1, max_col + 1):
                    cell = ws.cell(row=r, column=c)
                    cell.border = thin
                    cell.alignment = Alignment(horizontal=("left" if c in (1, 2) else "center"))

                    # Füllfarbe anwenden
                    if c in column_fills:
                        cell.fill = column_fills[c]

            # Trennzeilen nach Samstagen einfügen
            self._insert_saturday_separators(ws, df, GREY, thin, max_col)

        messagebox.showinfo("Erfolg", f"Excel-Datei erfolgreich gespeichert: {filename}")

    def _insert_saturday_separators(self, ws, df: pd.DataFrame, grey_fill, border, max_col: int) -> None:
        """Fügt graue Trennzeilen nach jedem Samstag ein"""
        saturday_rows = []
        for i, weekday in enumerate(df["Wochentag"].tolist(), start=2):  # +1 Header, +1 1-basiert
            if weekday == "Samstag":
                saturday_rows.append(i)

        # Von unten nach oben einfügen um Zeilenindizes nicht zu verschieben
        for row_idx in sorted(saturday_rows, reverse=True):
            insert_row = row_idx + 1
            ws.insert_rows(insert_row)
            for c in range(1, max_col + 1):
                cell = ws.cell(row=insert_row, column=c)
                cell.fill = grey_fill
                cell.border = border


def main():
    root = tk.Tk()
    app = ShiftPlanner(root)
    root.mainloop()


if __name__ == "__main__":
    main()
