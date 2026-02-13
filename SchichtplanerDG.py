import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import json
import pandas as pd
from datetime import datetime, timedelta
import os
from typing import Dict, List, Optional, Tuple

# Konstanten
DAYS_IN_PLANNING = 12
DAYS_PER_WEEK = 6
WORKING_DAYS = 10  # Nur Arbeitstage Mo-Fr (ohne Samstage)
WEEKDAY_NAMES = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"]
WINDOW_GEOMETRY = "1400x900"
MIN_WINDOW_SIZE = (1200, 800)
CONFIG_FILE = "shift_config.json"
HISTORY_FILE = "shift_history.json"
CATCHUP_DELAY_DAYS = 2  # Tage bis Nachhol-Versuch für übersprungene Mitarbeiter
CATCHUP_PRIORITY_DAYS = 4  # Nach diesen Tagen bekommt Nachhol absolute Priorität

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
            "pool_mo_mi_abwesend": [],     # Montag/Mittwoch nicht verfügbar
            "feiertage": []                # [{datum: "25.12", name: "1. Weihnachtstag", mitarbeiter: "XX"}]
        }

        # Cache für Performance
        self._all_employees_cache: Optional[List[str]] = None
        self._cache_dirty = True

        # Interne Zustandsverwaltung
        self.planning_result: List[Dict[str, str]] = []
        self.absences: Dict[int, List[str]] = {}
        self.saved_plans: List[Dict] = []  # Gespeicherte Pläne für Auswertung
        self.manual_stats_corrections: Dict[str, Dict[str, int]] = {} # Manuelle Korrekturen {Name: {VM: +1, ...}}
        self.history_file = HISTORY_FILE
        self._plan_manually_edited = False  # Flag: Plan wurde nachträglich geändert

        self.load_config()
        self.load_history()
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

    def load_history(self) -> None:
        """Lädt gespeicherte Pläne aus History-Datei"""
        try:
            if os.path.exists(self.history_file):
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.saved_plans = data.get("plans", [])
                    self.manual_stats_corrections = data.get("manual_corrections", {})
                    self._eval_filter_from = data.get("filter_from", "")
                    self._eval_filter_to = data.get("filter_to", "")
            else:
                self.saved_plans = []
                self.manual_stats_corrections = {}
                self._eval_filter_from = ""
                self._eval_filter_to = ""
        except Exception:
            self.saved_plans = []
            self.manual_stats_corrections = {}
            self._eval_filter_from = ""
            self._eval_filter_to = ""

    def save_history(self) -> None:
        """Speichert Pläne in History-Datei"""
        try:
            data = {
                "plans": self.saved_plans,
                "manual_corrections": getattr(self, 'manual_stats_corrections', {}),
                "filter_from": getattr(self, '_eval_filter_from', ""),
                "filter_to": getattr(self, '_eval_filter_to', "")
            }
            with open(self.history_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Speichern der Historie: {e}")


    def _get_last_tag10_data(self) -> Optional[Dict[str, str]]:
        """
        Holt Daten basierend auf letztem Plan:
        - Datum: Letzter Eintrag (Sa) + 2 Tage = Nächster Montag
        - MA: Rotation wird FORTGESETZT basierend auf Tag 10 (Do)
        - Support: Tag 1 (Mo) erzwingt Support (wie Montag Woche 2)
        """
        if not self.saved_plans:
            return None
        
        try:
            # 1. Letzten Plan finden
            sorted_plans = sorted(
                self.saved_plans,
                key=lambda p: datetime.strptime(p.get("start_date", "01.01.1900"), "%d.%m.%Y"),
                reverse=True
            )
            last_plan = sorted_plans[0]
            entries = last_plan.get("entries", [])
            
            if len(entries) < 12:
                return None

            # 2. Datum berechnen (Ende + 2 Tage)
            last_entry = entries[11] # Samstag
            last_date = datetime.strptime(last_entry.get("Datum", ""), "%d.%m.%Y")
            new_start_date = last_date + timedelta(days=2)
            
            # 3. Mitarbeiter aus Tag 10 (Index 9) finden
            tag10 = entries[9]
            prev_vm = tag10.get("Vormittag", "")
            prev_nm = tag10.get("Nachmittag", "")
            prev_support = tag10.get("Support", "")

            # Falls Support an Tag 10 leer war, rückwärts suchen für Rotation
            if not prev_support:
                for i in range(8, -1, -1):
                    s = entries[i].get("Support", "")
                    if s:
                        prev_support = s
                        break
            
            # 4. Nachfolger bestimmen
            def get_successor(name: str, pool_key: str) -> str:
                pool = self.config.get(pool_key, [])
                if not pool: return ""
                if name in pool:
                    idx = pool.index(name)
                    return pool[(idx + 1) % len(pool)]
                return pool[0] # Fallback: Erster der Liste

            next_vm = get_successor(prev_vm, "pool_vm_alle")
            next_nm = get_successor(prev_nm, "pool_nm_alle")
            # Support ist für Tag 1 (Montag) PFLICHT ("wie Montag Woche 2")
            next_support = get_successor(prev_support, "pool_vm_support")

            return {
                "start_date": new_start_date.strftime("%d.%m.%Y"),
                "vm": next_vm,
                "nm": next_nm,
                "support": next_support
            }
            
        except Exception:
            pass
        
        return None



    def _on_auto_mode_toggle(self) -> None:
        """Handler für Automatik-Modus Toggle"""
        if self.auto_mode_var.get():
            # Automatik aktiviert: Felder aus letztem Tag 10 befüllen
            tag10_data = self._get_last_tag10_data()
            
            if tag10_data:
                # Felder leeren und neu befüllen
                self.start_date_entry.delete(0, tk.END)
                self.start_date_entry.insert(0, tag10_data["start_date"])
                
                self.first_vm_entry.delete(0, tk.END)
                self.first_vm_entry.insert(0, tag10_data["vm"])
                
                self.first_nm_entry.delete(0, tk.END)
                self.first_nm_entry.insert(0, tag10_data["nm"])
                
                self.first_support_entry.delete(0, tk.END)
                self.first_support_entry.insert(0, tag10_data["support"])
            else:
                messagebox.showinfo(
                    "Automatik-Modus",
                    "Keine gespeicherten Pläne gefunden.\n"
                    "Bitte zuerst einen Plan speichern oder manuell eingeben."
                )
                self.auto_mode_var.set(False)

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
        # Tab 3: Auswertung
        self.create_evaluation_tab(notebook)
        # Tab 4: Historie
        self.create_history_tab(notebook)
        # Tab 5: Hilfe
        self.create_help_tab(notebook)

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
        self.pool_mo_mi_abwesend_entry = self._create_pool_entry(
            left_frame, 5, "Pool F - Mo/Mi NICHT verfügbar:", "pool_mo_mi_abwesend",
            font=('TkDefaultFont', 9, 'bold'), foreground='#e67300'
        )

        # Speichern Button
        ttk.Button(left_frame, text="Pools speichern", command=self.save_pools).grid(
            row=6, column=1, pady=15, sticky="e"
        )

        info_text = """Hinweise:
• Mitarbeiter mit Komma trennen (z.B. RR,AN,MH)
• Reihenfolge wird bei der Planung berücksichtigt"""
        ttk.Label(left_frame, text=info_text, justify="left", font=('TkDefaultFont', 8),
                 foreground='gray').grid(row=7, column=0, columnspan=2, padx=5, pady=5, sticky="w")

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
        ttk.Button(input_frame, text="Plan speichern", command=self.save_current_plan).grid(row=0, column=9, padx=5)
        ttk.Button(input_frame, text="Excel Export", command=self.export_excel).grid(row=0, column=10, padx=5)

        # Automatik-Modus (zweite Zeile)
        auto_frame = ttk.Frame(settings_frame)
        auto_frame.pack(fill="x", pady=(10, 0))
        
        self.auto_mode_var = tk.BooleanVar(value=False)
        self.auto_mode_check = ttk.Checkbutton(
            auto_frame, 
            text="Automatik-Modus (übernimmt Werte vom letzten gespeicherten Tag 10)",
            variable=self.auto_mode_var,
            command=self._on_auto_mode_toggle
        )
        self.auto_mode_check.pack(side="left")

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

        # Double-click zum Bearbeiten
        self.result_tree.bind("<Double-1>", self._on_result_double_click)
        self._edit_entry = None  # Aktives Eingabefeld

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
                                      values=[f"Tag {i + 1}" for i in range(WORKING_DAYS)])
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

    def create_evaluation_tab(self, notebook: ttk.Notebook) -> None:
        """Erstellt Tab für Auswertung der Notdienste"""
        eval_frame = ttk.Frame(notebook, padding=10)
        notebook.add(eval_frame, text="Auswertung")

        # === FILTERBEREICH ===
        filter_frame = ttk.LabelFrame(eval_frame, text=" Datumsfilter ", padding=10)
        filter_frame.pack(fill="x", pady=(0, 10))

        filter_input = ttk.Frame(filter_frame)
        filter_input.pack(fill="x")

        ttk.Label(filter_input, text="Von (TT.MM.YYYY):").grid(row=0, column=0, sticky="w", padx=(0, 5))
        self.eval_from_var = tk.StringVar(value=getattr(self, '_eval_filter_from', ''))
        self.eval_from_entry = ttk.Entry(filter_input, textvariable=self.eval_from_var, width=14)
        self.eval_from_entry.grid(row=0, column=1, padx=(0, 15))

        ttk.Label(filter_input, text="Bis (TT.MM.YYYY):").grid(row=0, column=2, sticky="w", padx=(0, 5))
        self.eval_to_var = tk.StringVar(value=getattr(self, '_eval_filter_to', ''))
        self.eval_to_entry = ttk.Entry(filter_input, textvariable=self.eval_to_var, width=14)
        self.eval_to_entry.grid(row=0, column=3, padx=(0, 15))

        ttk.Button(filter_input, text="Aktualisieren", command=self.update_evaluation_display).grid(row=0, column=4, padx=5)
        ttk.Button(filter_input, text="Filter speichern", command=self._save_filter).grid(row=0, column=5, padx=5)

        # === STATISTIK-ANZEIGE ===
        stats_frame = ttk.LabelFrame(eval_frame, text=" Statistik pro Mitarbeiter ", padding=10)
        stats_frame.pack(fill="both", expand=True)

        columns = ("MA", "VM", "NM", "Support", "Gesamt")
        self.stats_tree = ttk.Treeview(stats_frame, columns=columns, show="headings",
                                       height=20, style="Modern.Treeview")

        column_config = {
            "MA": (100, "center"),
            "VM": (80, "center"),
            "NM": (80, "center"),
            "Support": (80, "center"),
            "Gesamt": (80, "center")
        }
        for col_name, (width, anchor) in column_config.items():
            self.stats_tree.heading(col_name, text=col_name)
            self.stats_tree.column(col_name, width=width, minwidth=60, anchor=anchor)

        scrollbar_stats = ttk.Scrollbar(stats_frame, orient="vertical", command=self.stats_tree.yview)
        self.stats_tree.configure(yscrollcommand=scrollbar_stats.set)

        self.stats_tree.pack(side="left", fill="both", expand=True)
        scrollbar_stats.pack(side="right", fill="y")

        # Double-click zum Bearbeiten der Statistik
        self.stats_tree.bind("<Double-1>", self._on_stats_double_click)
        self._stats_edit_entry = None

        # === INFO-BEREICH ===
        info_frame = ttk.Frame(eval_frame)
        info_frame.pack(fill="x", pady=(10, 0))

        self.stats_info_label = ttk.Label(info_frame, text="Keine Pläne gespeichert.", font=('Segoe UI', 9))
        self.stats_info_label.pack(side="left")

        ttk.Button(info_frame, text="Alle Pläne löschen", command=self._clear_all_plans).pack(side="right")

        # Initial aktualisieren
        self.root.after(200, self.update_evaluation_display)

    # -------------------- Historie --------------------

    def create_history_tab(self, notebook: ttk.Notebook) -> None:
        """Erstellt Tab für historische Planungsübersicht"""
        history_frame = ttk.Frame(notebook, padding=10)
        notebook.add(history_frame, text="Historie")

        # === OBERER BEREICH: Planübersicht ===
        overview_frame = ttk.LabelFrame(history_frame, text=" Gespeicherte Pläne ", padding=10)
        overview_frame.pack(fill="x", pady=(0, 10))

        # Treeview für Planübersicht
        hist_columns = ("Zeitraum", "Gespeichert", "Geändert", "Anmerkung")
        self.history_tree = ttk.Treeview(overview_frame, columns=hist_columns, show="headings",
                                          height=8, style="Modern.Treeview")
        self.history_tree.heading("Zeitraum", text="Zeitraum")
        self.history_tree.heading("Gespeichert", text="Gespeichert am")
        self.history_tree.heading("Geändert", text="Geändert")
        self.history_tree.heading("Anmerkung", text="Anmerkung")
        self.history_tree.column("Zeitraum", width=200, anchor="center")
        self.history_tree.column("Gespeichert", width=140, anchor="center")
        self.history_tree.column("Geändert", width=70, anchor="center")
        self.history_tree.column("Anmerkung", width=300, anchor="w")

        scrollbar_hist = ttk.Scrollbar(overview_frame, orient="vertical", command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=scrollbar_hist.set)
        self.history_tree.pack(side="left", fill="both", expand=True)
        scrollbar_hist.pack(side="right", fill="y")

        # Buttons
        btn_frame = ttk.Frame(history_frame)
        btn_frame.pack(fill="x", pady=(0, 10))
        ttk.Button(btn_frame, text="Plan anzeigen", command=self._view_plan_detail).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="Anmerkung bearbeiten", command=self._edit_plan_note).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="Plan löschen", command=self._delete_single_plan).pack(side="left", padx=2)

        # === UNTERER BEREICH: Detailansicht ===
        detail_frame = ttk.LabelFrame(history_frame, text=" Detailansicht ", padding=10)
        detail_frame.pack(fill="both", expand=True)

        detail_columns = ("Datum", "Tag", "VM", "NM", "Support")
        self.history_detail_tree = ttk.Treeview(detail_frame, columns=detail_columns, show="headings",
                                                 height=14, style="Modern.Treeview")
        for col_name, width in [("Datum", 85), ("Tag", 80), ("VM", 80), ("NM", 80), ("Support", 80)]:
            self.history_detail_tree.heading(col_name, text=col_name)
            self.history_detail_tree.column(col_name, width=width, minwidth=50, anchor="center")

        scrollbar_detail = ttk.Scrollbar(detail_frame, orient="vertical", command=self.history_detail_tree.yview)
        self.history_detail_tree.configure(yscrollcommand=scrollbar_detail.set)
        self.history_detail_tree.pack(side="left", fill="both", expand=True)
        scrollbar_detail.pack(side="right", fill="y")

        # Doppelklick zum Bearbeiten von Einträgen in der Detailansicht
        self.history_detail_tree.bind("<Double-1>", self._on_history_detail_double_click)
        self._hist_edit_entry = None
        self._hist_current_plan_idx = None  # Index des aktuell angezeigten Plans

        # Initial aktualisieren
        self.root.after(300, self.update_history_display)

    def update_history_display(self) -> None:
        """Aktualisiert die Übersicht der gespeicherten Pläne"""
        if not hasattr(self, 'history_tree'):
            return
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)

        # Pläne chronologisch sortiert anzeigen (neueste zuerst)
        sorted_plans = sorted(
            enumerate(self.saved_plans),
            key=lambda x: datetime.strptime(x[1].get("start_date", "01.01.1900"), "%d.%m.%Y"),
            reverse=True
        )

        for orig_idx, plan in sorted_plans:
            zeitraum = f"{plan.get('start_date', '?')} – {plan.get('end_date', '?')}"
            gespeichert = plan.get("saved_at", "?")
            geaendert = "✓" if plan.get("manuell_geaendert", False) else "✗"
            anmerkung = plan.get("anmerkung", "")
            # Speichere den Original-Index als Tag für spätere Referenz
            self.history_tree.insert("", tk.END, values=(zeitraum, gespeichert, geaendert, anmerkung),
                                      tags=(str(orig_idx),))

    def _get_selected_plan_index(self) -> Optional[int]:
        """Gibt den Index des ausgewählten Plans zurück"""
        selected = self.history_tree.selection()
        if not selected:
            messagebox.showwarning("Warnung", "Bitte einen Plan auswählen!")
            return None
        tags = self.history_tree.item(selected[0], "tags")
        if tags:
            return int(tags[0])
        return None

    def _view_plan_detail(self) -> None:
        """Zeigt die Einträge des ausgewählten Plans in der Detailansicht"""
        plan_idx = self._get_selected_plan_index()
        if plan_idx is None:
            return

        self._hist_current_plan_idx = plan_idx
        plan = self.saved_plans[plan_idx]

        # Detail-Treeview leeren
        for item in self.history_detail_tree.get_children():
            self.history_detail_tree.delete(item)

        # Einträge anzeigen
        for entry in plan.get("entries", []):
            tag_kurz = entry.get("Wochentag", "")[:2] if entry.get("Wochentag") else ""
            self.history_detail_tree.insert(
                "", tk.END,
                values=(entry.get("Datum", ""), tag_kurz,
                       entry.get("Vormittag", ""), entry.get("Nachmittag", ""),
                       entry.get("Support", ""))
            )

    def _edit_plan_note(self) -> None:
        """Öffnet ein Dialogfenster zur Bearbeitung der Anmerkung"""
        plan_idx = self._get_selected_plan_index()
        if plan_idx is None:
            return

        plan = self.saved_plans[plan_idx]
        current_note = plan.get("anmerkung", "")

        # Dialog erstellen
        dialog = tk.Toplevel(self.root)
        dialog.title("Anmerkung bearbeiten")
        dialog.geometry("400x200")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="Anmerkung:", font=('Segoe UI', 10, 'bold')).pack(padx=10, pady=(10, 5), anchor="w")

        text_widget = tk.Text(dialog, wrap=tk.WORD, width=45, height=5, font=('Segoe UI', 10))
        text_widget.pack(padx=10, pady=5, fill="both", expand=True)
        text_widget.insert("1.0", current_note)
        text_widget.focus_set()

        def save_note():
            new_note = text_widget.get("1.0", tk.END).strip()
            self.saved_plans[plan_idx]["anmerkung"] = new_note
            self.save_history()
            self.update_history_display()
            dialog.destroy()

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill="x", padx=10, pady=10)
        ttk.Button(btn_frame, text="Speichern", command=save_note).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Abbrechen", command=dialog.destroy).pack(side="left", padx=5)

        dialog.bind("<Escape>", lambda e: dialog.destroy())

    def _delete_single_plan(self) -> None:
        """Löscht einen einzelnen ausgewählten Plan"""
        plan_idx = self._get_selected_plan_index()
        if plan_idx is None:
            return

        plan = self.saved_plans[plan_idx]
        zeitraum = f"{plan.get('start_date', '?')} – {plan.get('end_date', '?')}"

        if messagebox.askyesno("Bestätigung", f"Plan '{zeitraum}' wirklich löschen?"):
            self.saved_plans.pop(plan_idx)
            self._hist_current_plan_idx = None
            # Detail-Treeview leeren
            for item in self.history_detail_tree.get_children():
                self.history_detail_tree.delete(item)
            self.save_history()
            self.update_history_display()
            self.update_evaluation_display()
            messagebox.showinfo("Erfolg", f"Plan '{zeitraum}' gelöscht!")

    def _on_history_detail_double_click(self, event) -> None:
        """Handler für Doppelklick zum Bearbeiten einer Zelle in der Detailansicht"""
        if self._hist_current_plan_idx is None:
            return

        if self._hist_edit_entry:
            self._hist_confirm_edit()

        region = self.history_detail_tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        column = self.history_detail_tree.identify_column(event.x)
        item = self.history_detail_tree.identify_row(event.y)

        if not item or not column:
            return

        col_idx = int(column.replace("#", ""))
        # Nur VM, NM, Support editierbar (Spalten 3, 4, 5)
        if col_idx < 3:
            return

        values = list(self.history_detail_tree.item(item, "values"))
        current_value = values[col_idx - 1] if col_idx - 1 < len(values) else ""

        bbox = self.history_detail_tree.bbox(item, column)
        if not bbox:
            return

        x, y, width, height = bbox

        self._hist_edit_entry = tk.Entry(self.history_detail_tree, width=width // 8)
        self._hist_edit_entry.place(x=x, y=y, width=width, height=height)
        self._hist_edit_entry.insert(0, current_value)
        self._hist_edit_entry.select_range(0, tk.END)
        self._hist_edit_entry.focus_set()

        self._hist_edit_item = item
        self._hist_edit_col_idx = col_idx
        self._hist_edit_row_idx = self.history_detail_tree.index(item)

        self._hist_edit_entry.bind("<Return>", lambda e: self._hist_confirm_edit())
        self._hist_edit_entry.bind("<Escape>", lambda e: self._hist_cancel_edit())
        self._hist_edit_entry.bind("<FocusOut>", lambda e: self._hist_confirm_edit())

    def _hist_confirm_edit(self) -> None:
        """Bestätigt die Bearbeitung in der Detailansicht und aktualisiert den gespeicherten Plan"""
        if not self._hist_edit_entry or self._hist_current_plan_idx is None:
            return

        new_value = self._hist_edit_entry.get().strip()

        # Treeview aktualisieren
        values = list(self.history_detail_tree.item(self._hist_edit_item, "values"))
        values[self._hist_edit_col_idx - 1] = new_value
        self.history_detail_tree.item(self._hist_edit_item, values=values)

        # Gespeicherten Plan aktualisieren
        col_map = {3: "Vormittag", 4: "Nachmittag", 5: "Support"}
        if self._hist_edit_col_idx in col_map:
            plan = self.saved_plans[self._hist_current_plan_idx]
            entries = plan.get("entries", [])
            if self._hist_edit_row_idx < len(entries):
                entries[self._hist_edit_row_idx][col_map[self._hist_edit_col_idx]] = new_value
                plan["manuell_geaendert"] = True
                self.save_history()
                self.update_history_display()

        self._hist_cleanup_edit()

    def _hist_cancel_edit(self) -> None:
        """Bricht die Bearbeitung in der Detailansicht ab"""
        self._hist_cleanup_edit()

    def _hist_cleanup_edit(self) -> None:
        """Räumt das Edit-Widget in der Detailansicht auf"""
        if self._hist_edit_entry:
            self._hist_edit_entry.destroy()
            self._hist_edit_entry = None
        self._hist_edit_item = None
        self._hist_edit_col_idx = None
        self._hist_edit_row_idx = None

    def create_help_tab(self, notebook: ttk.Notebook) -> None:
        """Erstellt Tab für Hilfe und Anleitung"""
        help_frame = ttk.Frame(notebook, padding=10)
        notebook.add(help_frame, text="Hilfe")

        text_area = scrolledtext.ScrolledText(help_frame, wrap=tk.WORD, width=100, height=30,
                                              font=('Segoe UI', 10))
        text_area.pack(fill="both", expand=True)

        # Tags für Formatierung
        text_area.tag_config("h1", font=('Segoe UI', 14, 'bold'), foreground="#2c3e50", spacing3=10)
        text_area.tag_config("h2", font=('Segoe UI', 12, 'bold'), foreground="#34495e", spacing3=5)
        text_area.tag_config("bold", font=('Segoe UI', 10, 'bold'))
        text_area.tag_config("italic", font=('Segoe UI', 10, 'italic'))

        help_content = [
            ("Anleitung & Funktionsweise", "h1"),
            
            ("\n1. Konfiguration (Pools)", "h2"),
            ("Zuerst müssen die Mitarbeiter in die entsprechenden Pools eingetragen werden (Tab 'Pool-Konfiguration').\n"
             "• ", ""), ("Pool A (VM Alle):", "bold"), (" Mitarbeiter, die vormittags alles machen können.\n"
             "• ", ""), ("Pool B (VM Teilweise):", "bold"), (" Mitarbeiter, die vormittags Unterstützung brauchen.\n"
             "• ", ""), ("Pool C (VM Support):", "bold"), (" Mitarbeiter, die Pool B unterstützen können.\n"
             "• ", ""), ("Pool D (NM Alle):", "bold"), (" Mitarbeiter für den Nachmittag.\n"
             "• ", ""), ("Pool E:", "bold"), (" Mitarbeiter, die Freitags nie verfügbar sind.\n"
             "• ", ""), ("Pool F:", "bold"), (" Mitarbeiter, die Mo/Mi nie verfügbar sind.\n"
             "Die Reihenfolge der Eingabe bestimmt die initiale Rotationsreihenfolge.", ""),

            ("\n2. Planung erstellen", "h2"),
            ("Im Tab 'Schichtplanung':\n"
             "1. Startdatum (Montag) eingeben.\n"
             "2. Die Mitarbeiter für den ", ""), ("aller ersten Tag", "bold"), (" manuell festlegen (als Startpunkt).\n"
             "3. Auf 'Planung erstellen' klicken.\n"
             "Das System berechnet nun automatisch die folgenden 12 Tage (2 Wochen).", ""),

            ("\n3. Automatische Auswahlkriterien", "h2"),
            ("Das Programm entscheidet nach folgenden Regeln, wer wann eingeteilt wird:\n", ""),
            
            ("A) Rotation (Gerechtigkeit)", "bold"),
            ("\nDas System merkt sich für jeden Pool die letzte Position. Beim nächsten Tag wird einfach der nächste Mitarbeiter in der Liste gewählt. "
             "Dadurch kommt jeder gleich oft dran (Reihum-Prinzip).\n", ""),

            ("B) Abwesenheiten (Priorität 1)", "bold"),
            ("\nEin Mitarbeiter wird ÜBERSPRUNGEN, wenn:\n"
             "• Er im Tab 'Schichtplanung' rechts unter 'Abwesenheiten' eingetragen ist.\n"
             "• Er im Pool 'Freitag abwesend' ist und der Tag ein Freitag ist.\n"
             "• Er im Pool 'Mo/Mi abwesend' ist und der Tag Mo oder Mi ist.\n"
             "• Er für die aktuelle Woche als Feiertags-Notdienst eingeteilt ist.\n", ""),

            ("C) Ruhezeiten (Priorität 2)", "bold"),
            ("\n• Wer am Vortag Nachmittagsschicht hatte, wird am nächsten Tag NICHT für den Vormittag eingeteilt.\n", ""),

            ("D) Support-Logik (Vormittag)", "bold"),
            ("\n• Wird ein Mitarbeiter aus ", ""), ("Pool B (Teilweise)", "bold"), (" gewählt, MUSS zwingend ein Support-Mitarbeiter aus ", ""), ("Pool C", "bold"), (" dazu kommen.\n"
             "• Ist der gewählte Mitarbeiter aus Pool A ('Vollprofi'), wird kein Support benötigt.\n"
             "• ", ""), ("Ausnahme Montag Woche 2:", "bold"), (" Am Montag der zweiten Woche wird IMMER ein Support eingeplant (höheres Aufkommen).\n", ""),

            ("\n4. Feiertage", "h2"),
            ("Feiertage werden im ersten Tab definiert. Für einen Feiertag wird ein EINER Mitarbeiter festgelegt, "
             "der dann den ", ""), ("kompletten Tag", "italic"), (" übernimmt. Dieser Mitarbeiter wird für den Rest der Woche aus der regulären Rotation genommen, damit er nicht überlastet wird.", ""),
             
            ("\n5. Auswertung", "h2"),
            ("Fertige Pläne können gespeichert werden. Im Tab 'Auswertung' sehen Sie dann eine Statistik, wer wie oft welche Schicht hatte.", ""),

            ("\n6. Nachhol-System (NEU)", "h2"),
            ("Wenn ein Mitarbeiter übersprungen wird (wegen Abwesenheit oder weil er bereits eine andere Schicht am selben Tag hat), "
             "wird er automatisch in eine ", ""), ("Nachhol-Queue", "bold"), (" aufgenommen.\n\n", ""),
            ("• ", ""), ("Nach 2 Tagen:", "bold"), (" Der übersprungene MA wird bevorzugt eingeplant, wenn der reguläre Kandidat nicht verfügbar ist.\n", ""),
            ("• ", ""), ("Nach 4 Tagen:", "bold"), (" Der übersprungene MA bekommt absolute Priorität und wird auch vor einem verfügbaren regulären Kandidaten eingeplant.\n", ""),
            ("• ", ""), ("Keine Dopplung:", "bold"), (" Nach einer Nachhol-Einplanung wird die Pool-Position so angepasst, dass derselbe MA nicht zweimal in kurzer Zeit drankommt.\n", ""),
            ("\nDieses System sorgt für maximale Fairness auch bei häufigen Abwesenheiten.", "")
        ]

        for text, tag in help_content:
            text_area.insert(tk.END, text, tag)
        
        text_area.configure(state='disabled')  # Schreibgeschützt

    def save_current_plan(self) -> None:
        """Speichert den aktuellen Plan in die Historie"""
        if not self.planning_result:
            messagebox.showwarning("Warnung", "Keine Planung vorhanden!")
            return

        # Datum des Plans ermitteln
        start_date = self.planning_result[0]["Datum"] if self.planning_result else ""
        end_date = self.planning_result[-1]["Datum"] if self.planning_result else ""

        # Prüfen ob Plan bereits existiert
        for plan in self.saved_plans:
            if plan.get("start_date") == start_date and plan.get("end_date") == end_date:
                if messagebox.askyesno("Plan existiert", "Ein Plan für diesen Zeitraum existiert bereits. Überschreiben?"):
                    self.saved_plans.remove(plan)
                    break
                else:
                    return

        # Plan speichern
        new_plan = {
            "start_date": start_date,
            "end_date": end_date,
            "saved_at": datetime.now().strftime("%d.%m.%Y %H:%M"),
            "entries": [row.copy() for row in self.planning_result],
            "anmerkung": "",
            "manuell_geaendert": self._plan_manually_edited
        }
        self.saved_plans.append(new_plan)
        self._plan_manually_edited = False  # Reset nach Speichern
        self.save_history()
        messagebox.showinfo("Erfolg", f"Plan vom {start_date} bis {end_date} gespeichert!")
        self.update_evaluation_display()
        self.update_history_display()

    def update_evaluation_display(self) -> None:
        """Aktualisiert die Statistik-Anzeige basierend auf Filter"""
        # Filter lesen
        filter_from = self.eval_from_var.get().strip() if hasattr(self, 'eval_from_var') else ""
        filter_to = self.eval_to_var.get().strip() if hasattr(self, 'eval_to_var') else ""

        # Statistik berechnen
        stats = self._calculate_statistics(filter_from, filter_to)

        # Treeview leeren
        for item in self.stats_tree.get_children():
            self.stats_tree.delete(item)

        # Statistik sortieren: Nach Gesamt (absteigend), dann Name (aufsteigend)
        # Wir erstellen erst eine Liste mit (Name, Counts, Total)
        stats_list = []
        for ma, counts in stats.items():
            total = counts["VM"] + counts["NM"] + counts["Support"]
            stats_list.append((ma, counts, total))

        # Sortieren: -total für absteigend, name für aufsteigend
        stats_list.sort(key=lambda x: (-x[2], x[0]))

        # Statistik anzeigen
        for ma, counts, total in stats_list:
            self.stats_tree.insert("", tk.END, values=(
                ma, counts["VM"], counts["NM"], counts["Support"], total
            ))

        # Info aktualisieren
        plan_count = len(self.saved_plans)
        if plan_count > 0:
            date_range = ""
            if filter_from or filter_to:
                date_range = f" (Filter: {filter_from or '...'} bis {filter_to or '...'})"
            self.stats_info_label.config(text=f"{plan_count} Plan(e) gespeichert.{date_range}")
        else:
            self.stats_info_label.config(text="Keine Pläne gespeichert.")

    def _calculate_statistics(self, filter_from: str, filter_to: str) -> Dict[str, Dict[str, int]]:
        """Berechnet Statistik über alle gespeicherten Pläne"""
        stats: Dict[str, Dict[str, int]] = {}

        # Filter-Daten parsen
        from_date = None
        to_date = None
        try:
            if filter_from:
                from_date = datetime.strptime(filter_from, "%d.%m.%Y")
            if filter_to:
                to_date = datetime.strptime(filter_to, "%d.%m.%Y")
        except ValueError:
            pass  # Ungültige Filter ignorieren

        for plan in self.saved_plans:
            for entry in plan.get("entries", []):
                # Datum prüfen
                try:
                    entry_date = datetime.strptime(entry["Datum"], "%d.%m.%Y")
                    if from_date and entry_date < from_date:
                        continue
                    if to_date and entry_date > to_date:
                        continue
                except ValueError:
                    continue

                # Zählen
                for shift_type, key in [("Vormittag", "VM"), ("Nachmittag", "NM"), ("Support", "Support")]:
                    ma = entry.get(shift_type, "").strip()
                    if ma:
                        if ma not in stats:
                            stats[ma] = {"VM": 0, "NM": 0, "Support": 0}
                        stats[ma][key] += 1

                        if ma not in stats:
                            stats[ma] = {"VM": 0, "NM": 0, "Support": 0}
                        stats[ma][key] += 1

        # Manuelle Korrekturen anwenden
        for ma, corrections in self.manual_stats_corrections.items():
            if ma not in stats:
                 # Wenn MA nicht in den Plänen, aber eine Korrektur existiert, trotzdem anzeigen
                 if any(corrections.values()): # Nur wenn Korrektur != 0
                     stats[ma] = {"VM": 0, "NM": 0, "Support": 0}
            
            if ma in stats:
                for key in ["VM", "NM", "Support"]:
                    stats[ma][key] += corrections.get(key, 0)

        return stats

    def _save_filter(self) -> None:
        """Speichert den aktuellen Datumsfilter"""
        self._eval_filter_from = self.eval_from_var.get().strip()
        self._eval_filter_to = self.eval_to_var.get().strip()
        self.save_history()
        messagebox.showinfo("Erfolg", "Filter gespeichert!")

    def _on_stats_double_click(self, event) -> None:
        """Handler für Doppleclick auf Statistik"""
        # Aktives Eingabefeld schließen
        if getattr(self, '_stats_edit_entry', None):
            self._stats_cancel_edit()

        region = self.stats_tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        column = self.stats_tree.identify_column(event.x)
        item = self.stats_tree.identify_row(event.y)

        if not item or not column:
            return

        # Spaltenindex: #1=MA, #2=VM, #3=NM, #4=Support, #5=Gesamt
        col_idx = int(column.replace("#", ""))
        
        # MA und Gesamt nicht editierbar
        if col_idx == 1 or col_idx == 5:
            return

        # Aktuelle Werte
        values = list(self.stats_tree.item(item, "values"))
        current_val = values[col_idx - 1]

        bbox = self.stats_tree.bbox(item, column)
        if not bbox:
            return
        
        x, y, width, height = bbox

        self._stats_edit_entry = tk.Entry(self.stats_tree, width=width // 8)
        self._stats_edit_entry.place(x=x, y=y, width=width, height=height)
        self._stats_edit_entry.insert(0, current_val)
        self._stats_edit_entry.select_range(0, tk.END)
        self._stats_edit_entry.focus_set()

        # Metadaten speichern
        self._stats_edit_item = item
        self._stats_edit_col_idx = col_idx
        # MA Name für Speicherung
        self._stats_edit_ma = values[0] 

        self._stats_edit_entry.bind("<Return>", lambda e: self._confirm_stats_edit())
        self._stats_edit_entry.bind("<Escape>", lambda e: self._stats_cancel_edit())
        self._stats_edit_entry.bind("<FocusOut>", lambda e: self._confirm_stats_edit())

    def _confirm_stats_edit(self) -> None:
        """Bestätigt Statistik-Bearbeitung"""
        if not getattr(self, '_stats_edit_entry', None):
            return

        try:
            new_val_str = self._stats_edit_entry.get().strip()
            new_val = int(new_val_str)
        except ValueError:
            # Ungültige Eingabe -> Abbruch
            self._stats_cancel_edit()
            return
        
        # Wir müssen den Delta berechnen. 
        # Das Problem: Der angezeigte Wert ist (Berechnet + Manuell).
        # Wir wollen Manuell_Neu so setzen, dass (Berechnet + Manuell_Neu) = Eingabe_Neu
        # => Manuell_Neu = Eingabe_Neu - Berechnet
        
        # 1. Berechneten Wert (ohne Korrektur) neu ermitteln
        filter_from = self.eval_from_var.get().strip() if hasattr(self, 'eval_from_var') else ""
        filter_to = self.eval_to_var.get().strip() if hasattr(self, 'eval_to_var') else ""
        
        # Temporär Korrekturen deaktivieren um Basiswert zu bekommen
        current_corrections = self.manual_stats_corrections
        self.manual_stats_corrections = {} 
        base_stats = self._calculate_statistics(filter_from, filter_to)
        self.manual_stats_corrections = current_corrections # Restore

        ma = self._stats_edit_ma
        col_map = {2: "VM", 3: "NM", 4: "Support"}
        key = col_map[self._stats_edit_col_idx]

        base_val = 0
        if ma in base_stats:
            base_val = base_stats[ma].get(key, 0)
        
        # Neue Korrektur berechnen
        correction = new_val - base_val

        # Speichern
        if ma not in self.manual_stats_corrections:
            self.manual_stats_corrections[ma] = {}
        
        self.manual_stats_corrections[ma][key] = correction
        
        # Wenn Korrektur 0 ist, Eintrag ggf. bereinigen
        if correction == 0:
            if key in self.manual_stats_corrections[ma]:
                del self.manual_stats_corrections[ma][key]
            if not self.manual_stats_corrections[ma]:
                del self.manual_stats_corrections[ma]

        self.save_history()
        self._stats_cancel_edit()
        self.update_evaluation_display()

    def _stats_cancel_edit(self) -> None:
        if getattr(self, '_stats_edit_entry', None):
            self._stats_edit_entry.destroy()
            self._stats_edit_entry = None

    def _clear_all_plans(self) -> None:
        """Löscht alle gespeicherten Pläne"""
        if not self.saved_plans:
            messagebox.showinfo("Info", "Keine Pläne vorhanden.")
            return

        if messagebox.askyesno("Bestätigung", f"{len(self.saved_plans)} Plan(e) wirklich löschen?"):
            self.saved_plans = []
            self.save_history()
            self.update_evaluation_display()
            self.update_history_display()
            messagebox.showinfo("Erfolg", "Alle Pläne gelöscht!")

    # -------------------- Abwesenheiten --------------------

    def update_employee_list(self) -> None:
        """Aktualisiert die Mitarbeiterliste in der Combobox"""
        self._cache_dirty = True
        all_employees = self._get_all_employees()
        self.employee_combo['values'] = all_employees
        messagebox.showinfo("Info", f"Mitarbeiterliste aktualisiert: {len(all_employees)} Mitarbeiter gefunden")

    @staticmethod
    def _working_day_to_tag_nr(working_day: int) -> int:
        """Wandelt Arbeitstag-Index (0-9) in internen tag_nr (0-11, überspringe Sa=5,11) um."""
        week = working_day // 5
        day_in_week = working_day % 5
        return week * DAYS_PER_WEEK + day_in_week

    @staticmethod
    def _tag_nr_to_working_day(tag_nr: int) -> int:
        """Wandelt internen tag_nr (0-11) in Arbeitstag-Index (0-9) um."""
        week = tag_nr // DAYS_PER_WEEK
        day_in_week = tag_nr % DAYS_PER_WEEK
        if day_in_week >= 5:  # Samstag
            return -1  # Kein gültiger Arbeitstag
        return week * 5 + day_in_week

    def add_absence(self) -> None:
        """Fügt eine Abwesenheit hinzu (unterstützt mehrere, komma-getrennte Kürzel)."""
        employee = self.employee_var.get().strip()
        day_str = self.day_var.get().strip()

        if not employee or not day_str:
            messagebox.showwarning("Warnung", "Bitte Mitarbeiter und Tag auswählen!")
            return

        try:
            working_day = int(day_str.split()[1]) - 1  # "Tag 1" -> 0 (Arbeitstag-Index)
            if not (0 <= working_day < WORKING_DAYS):
                raise ValueError("Tag außerhalb des gültigen Bereichs")
            day_nr = self._working_day_to_tag_nr(working_day)  # -> interner tag_nr
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
            next_working_day = working_day + 1
            if next_working_day < WORKING_DAYS:
                self.day_var.set(f"Tag {next_working_day + 1}")
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
                working_day = int(day_str.split()[1]) - 1
                day_nr = self._working_day_to_tag_nr(working_day)
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
            working_day = self._tag_nr_to_working_day(day_nr)
            if working_day < 0:
                continue  # Samstage überspringen
            for employee in sorted(self.absences[day_nr]):
                self.absence_tree.insert("", tk.END, values=(f"Tag {working_day + 1}", employee))

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
                "pool_freitag_abwesend": self.pool_freitag_abwesend_entry,
                "pool_mo_mi_abwesend": self.pool_mo_mi_abwesend_entry
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

    def _get_absent_for_day(self, tag_nr: int, is_friday: bool, start_date: datetime, is_monday: bool = False, is_wednesday: bool = False) -> List[str]:
        """Gibt Liste der abwesenden Mitarbeiter für einen Tag zurück"""
        absent_today = self.absences.get(tag_nr, []).copy()

        if is_friday:
            # Freitags zusätzlich die generellen "Freitag nicht verfügbar"
            absent_today = list(set(absent_today + self.config["pool_freitag_abwesend"]))

        if is_monday or is_wednesday:
            # Montag/Mittwoch zusätzlich die generellen "Mo/Mi nicht verfügbar"
            absent_today = list(set(absent_today + self.config.get("pool_mo_mi_abwesend", [])))

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

    def _find_employee_with_catchup(self, pool: List[str], pool_key: str, start_pos: int,
                                     excluded: List[str], absent: List[str],
                                     catchup_queues: Dict[str, List[Dict]], current_tag_nr: int,
                                     pool_positions: Dict[str, int]) -> Tuple[Optional[str], int, List[str]]:
        """
        Sucht Mitarbeiter mit Berücksichtigung der Nachhol-Queue.
        Gibt (employee, new_position, skipped_employees) zurück.
        
        Logik:
        0. PRIORITÄTS-ESKALATION: Wer >= 4 Tage wartet, bekommt absolute Priorität
           - Wenn excluded wegen anderer Schicht am selben Tag: trotzdem einplanen!
        1. Prüfe ob der reguläre Kandidat (an start_pos) verfügbar ist
        2. Wenn ja → verwende diesen (normale Rotation)
        3. Wenn nein → prüfe Nachhol-Queue, dann andere aus Rotation
        4. Alle übersprungenen werden zur Queue hinzugefügt
        """
        skipped_employees: List[str] = []
        queue = catchup_queues.get(pool_key, [])
        
        if not pool:
            return None, start_pos, skipped_employees
        
        pool_size = len(pool)
        
        # 0. PRIORITÄTS-ESKALATION: Prüfe ob jemand in Queue zu lange wartet
        # Wichtig: Kopie der Queue durchlaufen, da wir sie modifizieren
        for i in range(len(queue) - 1, -1, -1):  # Rückwärts iterieren für sicheres pop()
            entry = queue[i]
            employee = entry["employee"]
            skipped_at = entry["skipped_at_tag"]
            waiting_days = current_tag_nr - skipped_at
            
            if waiting_days >= CATCHUP_PRIORITY_DAYS:
                # Dieser MA wartet zu lange - absolute Priorität!
                if employee not in absent:
                    # Auch wenn excluded (z.B. wegen VM am selben Tag) - trotzdem einplanen!
                    # Das ist der Fix: Nach 4 Tagen Warten hat Catch-up Vorrang
                    queue.pop(i)
                    # Pool-Position NACH diesem MA setzen (um Dopplung zu vermeiden)
                    if employee in pool:
                        emp_idx = pool.index(employee)
                        new_pos = (emp_idx + 1) % pool_size
                    else:
                        new_pos = (start_pos + 1) % pool_size
                    return employee, new_pos, skipped_employees
                else:
                    # Wirklich abwesend (Urlaub etc.) → Wartezeit zurücksetzen
                    entry["skipped_at_tag"] = current_tag_nr  # FIX: skipped_at zurücksetzen!
                    entry["available_from_tag"] = current_tag_nr + CATCHUP_DELAY_DAYS
        
        # 1. Prüfe zuerst den regulären Kandidaten (an start_pos)
        regular_candidate = pool[start_pos]
        if regular_candidate not in absent and regular_candidate not in excluded:
            # Regulärer Kandidat ist verfügbar! Normale Rotation.
            new_pos = (start_pos + 1) % pool_size
            return regular_candidate, new_pos, skipped_employees
        
        # 2. Regulärer Kandidat ist NICHT verfügbar → zur Nachhol-Queue hinzufügen
        already_in_queue = any(e["employee"] == regular_candidate for e in queue)
        if not already_in_queue:
            skipped_employees.append(regular_candidate)
        
        # 3. Jetzt prüfe Nachhol-Queue für Ersatz
        for i in range(len(queue) - 1, -1, -1):  # Rückwärts iterieren für sicheres pop()
            entry = queue[i]
            employee = entry["employee"]
            available_from = entry["available_from_tag"]
            
            if current_tag_nr >= available_from:
                # Dieser MA ist bereit für Nachhol-Versuch
                if employee not in absent and employee not in excluded:
                    # Verfügbar! Aus Queue entfernen und zurückgeben
                    queue.pop(i)
                    # Pool-Position NACH diesem MA setzen (um Dopplung zu vermeiden)
                    if employee in pool:
                        emp_idx = pool.index(employee)
                        new_pos = (emp_idx + 1) % pool_size
                    else:
                        new_pos = (start_pos + 1) % pool_size
                    return employee, new_pos, skipped_employees
                # Nicht verfügbar, aber available_from erreicht → available_from aktualisieren
                elif employee in absent:
                    entry["available_from_tag"] = current_tag_nr + CATCHUP_DELAY_DAYS
                # Wenn excluded (nicht absent), warten wir auf Prioritäts-Eskalation
        
        # 4. Keine Nachhol-Kandidaten verfügbar → nächsten aus normaler Rotation suchen
        for i in range(1, pool_size):  # Start bei 1, da 0 (regulärer Kandidat) bereits geprüft
            candidate_idx = (start_pos + i) % pool_size
            candidate = pool[candidate_idx]
            
            if candidate not in absent and candidate not in excluded:
                # Gefunden! Gib zurück mit neuer Position
                new_pos = (candidate_idx + 1) % pool_size
                return candidate, new_pos, skipped_employees
            else:
                # Auch diesen als übersprungen markieren (wenn nicht schon in Queue)
                already_in_queue = any(e["employee"] == candidate for e in queue)
                if not already_in_queue and candidate not in skipped_employees:
                    skipped_employees.append(candidate)
        
        return None, start_pos, skipped_employees

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
            
            # Nachhol-Queue für übersprungene Mitarbeiter
            catchup_queues: Dict[str, List[Dict]] = {
                "vm_alle": [],
                "nm_alle": [],
                "vm_support": []
            }

            # 12 Tage: Mo-Sa, Mo-Sa (2 Wochen ohne Sonntag)
            for tag_nr in range(DAYS_IN_PLANNING):
                # Wochentag berechnen
                weekday_in_week = tag_nr % DAYS_PER_WEEK
                week_number = tag_nr // DAYS_PER_WEEK
                current_date = start_date + timedelta(days=tag_nr + week_number)

                is_monday = weekday_in_week == 0
                is_wednesday = weekday_in_week == 2
                is_friday = weekday_in_week == 4
                is_saturday = weekday_in_week == 5
                weekday_german = WEEKDAY_NAMES[weekday_in_week]

                absent_today = self._get_absent_for_day(tag_nr, is_friday, start_date, is_monday, is_wednesday)

                # Feiertag prüfen: Datum im Format TT.MM vergleichen
                holiday_employee = None
                current_date_str = current_date.strftime("%d.%m")
                for holiday in self.config.get("feiertage", []):
                    if holiday["datum"] == current_date_str:
                        holiday_employee = holiday.get("mitarbeiter", "")
                        break

                if holiday_employee:
                    # Feiertag: Notdienst-MA deckt alle Schichten ab
                    vm_employee = holiday_employee
                    nm_employee = holiday_employee
                    support_employee = holiday_employee
                elif is_saturday:
                    # Samstag: keine Schichten
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
                        tag_nr, absent_today, pool_positions, catchup_queues
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
                  pool_positions: Dict[str, int],
                  catchup_queues: Dict[str, List[Dict]]) -> Tuple[Optional[str], Optional[str], str]:
        """Plant einen einzelnen Tag mit Nachhol-Queue. Gibt (vm_employee, nm_employee, support_employee) zurück"""
        yesterday_nm = self.planning_result[tag_nr - 1]["Nachmittag"] if tag_nr > 0 else None
        forbidden_vm = [yesterday_nm] if yesterday_nm else []
        used_today = []

        # Vormittag planen (mit Nachhol-Queue)
        vm_employee, new_vm_pos, vm_skipped = self._find_employee_with_catchup(
            self.config["pool_vm_alle"],
            "vm_alle",
            pool_positions["vm_alle"],
            forbidden_vm,
            absent_today,
            catchup_queues,
            tag_nr,
            pool_positions
        )
        
        pool_positions["vm_alle"] = new_vm_pos
        support_employee = ""

        # Übersprungene zur Queue hinzufügen
        for skipped in vm_skipped:
            catchup_queues["vm_alle"].append({
                "employee": skipped,
                "skipped_at_tag": tag_nr,
                "available_from_tag": tag_nr + CATCHUP_DELAY_DAYS
            })

        if vm_employee:
            used_today.append(vm_employee)

            # Support nötig?
            # 1. Immer wenn VM aus Pool B (braucht Unterstützung)
            # 2. Montag in Woche 2 (tag_nr == 6) - höheres Aufkommen
            needs_support = (vm_employee in self.config["pool_vm_teilweise"]) or (tag_nr == 6)
            
            if needs_support:
                support_employee, new_support_pos, support_skipped = self._find_employee_with_catchup(
                    self.config["pool_vm_support"],
                    "vm_support",
                    pool_positions["vm_support"],
                    [vm_employee] + forbidden_vm,
                    absent_today,
                    catchup_queues,
                    tag_nr,
                    pool_positions
                )
                pool_positions["vm_support"] = new_support_pos
                
                # Übersprungene Support-MA zur Queue
                for skipped in support_skipped:
                    catchup_queues["vm_support"].append({
                        "employee": skipped,
                        "skipped_at_tag": tag_nr,
                        "available_from_tag": tag_nr + CATCHUP_DELAY_DAYS
                    })
                
                if support_employee:
                    used_today.append(support_employee)
            else:
                # Kein separater Support nötig → VM-MA auch in Support-Spalte eintragen (Übersichtlichkeit)
                support_employee = vm_employee

        # Nachmittag planen (mit Nachhol-Queue)
        nm_employee, new_nm_pos, nm_skipped = self._find_employee_with_catchup(
            self.config["pool_nm_alle"],
            "nm_alle",
            pool_positions["nm_alle"],
            used_today,
            absent_today,
            catchup_queues,
            tag_nr,
            pool_positions
        )
        
        pool_positions["nm_alle"] = new_nm_pos

        # Übersprungene zur Queue hinzufügen
        for skipped in nm_skipped:
            catchup_queues["nm_alle"].append({
                "employee": skipped,
                "skipped_at_tag": tag_nr,
                "available_from_tag": tag_nr + CATCHUP_DELAY_DAYS
            })

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

    def _on_result_double_click(self, event) -> None:
        """Handler für Doppelklick zum Bearbeiten einer Zelle"""
        # Aktives Eingabefeld schließen
        if self._edit_entry:
            self._cancel_edit()

        # Klick-Position ermitteln
        region = self.result_tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        column = self.result_tree.identify_column(event.x)
        item = self.result_tree.identify_row(event.y)

        if not item or not column:
            return

        # Spaltenindex (1-basiert): #1=Datum, #2=Tag, #3=VM, #4=NM, #5=Support
        col_idx = int(column.replace("#", ""))

        # Nur VM, NM, Support editierbar (Spalten 3, 4, 5)
        if col_idx < 3:
            return

        # Aktuelle Werte holen
        values = list(self.result_tree.item(item, "values"))
        current_value = values[col_idx - 1] if col_idx - 1 < len(values) else ""

        # Zellen-Bounding-Box ermitteln
        bbox = self.result_tree.bbox(item, column)
        if not bbox:
            return

        x, y, width, height = bbox

        # Entry-Widget erstellen
        self._edit_entry = tk.Entry(self.result_tree, width=width // 8)
        self._edit_entry.place(x=x, y=y, width=width, height=height)
        self._edit_entry.insert(0, current_value)
        self._edit_entry.select_range(0, tk.END)
        self._edit_entry.focus_set()

        # Metadaten speichern
        self._edit_item = item
        self._edit_col_idx = col_idx
        self._edit_row_idx = self.result_tree.index(item)

        # Event-Bindings
        self._edit_entry.bind("<Return>", lambda e: self._confirm_edit())
        self._edit_entry.bind("<Escape>", lambda e: self._cancel_edit())
        self._edit_entry.bind("<FocusOut>", lambda e: self._confirm_edit())

    def _confirm_edit(self) -> None:
        """Bestätigt die Bearbeitung und aktualisiert die Daten"""
        if not self._edit_entry:
            return

        new_value = self._edit_entry.get().strip()

        # Treeview aktualisieren
        values = list(self.result_tree.item(self._edit_item, "values"))
        values[self._edit_col_idx - 1] = new_value
        self.result_tree.item(self._edit_item, values=values)

        # planning_result aktualisieren
        col_map = {3: "Vormittag", 4: "Nachmittag", 5: "Support"}
        if self._edit_col_idx in col_map:
            self.planning_result[self._edit_row_idx][col_map[self._edit_col_idx]] = new_value
            self._plan_manually_edited = True  # Flag setzen bei nachträglicher Änderung

        self._cleanup_edit()

    def _cancel_edit(self) -> None:
        """Bricht die Bearbeitung ab"""
        self._cleanup_edit()

    def _cleanup_edit(self) -> None:
        """Räumt das Edit-Widget auf"""
        if self._edit_entry:
            self._edit_entry.destroy()
            self._edit_entry = None
        self._edit_item = None
        self._edit_col_idx = None
        self._edit_row_idx = None

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
