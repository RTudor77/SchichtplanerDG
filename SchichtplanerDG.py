import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import pandas as pd
from datetime import datetime, timedelta
import os


class ShiftPlanner:
    def __init__(self, root):
        self.root = root
        self.root.title("Notdienst Schichtplaner v1.1")
        # Größeres Startfenster + sinnvolle Mindestgröße
        self.root.geometry("1400x900")
        self.root.minsize(1200, 800)

        # Konfigurationsdatei
        self.config_file = "shift_config.json"

        # Standardkonfiguration
        self.config = {
            "pool_vm_alle": [],            # Vormittag - können alles
            "pool_vm_teilweise": [],       # Vormittag - können nicht alles (brauchen Support)
            "pool_vm_support": [],         # Vormittag - Support für Pool B
            "pool_nm_alle": [],            # Nachmittag - können alles
            "pool_freitag_abwesend": []    # Freitags nicht verfügbar
        }

        self.load_config()
        self.create_gui()

    # -------------------- Konfiguration laden/speichern --------------------

    def load_config(self):
        """Lädt Konfiguration aus JSON-Datei"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)
                    if "pool_freitag_abwesend" not in loaded_config:
                        loaded_config["pool_freitag_abwesend"] = []
                    self.config = loaded_config
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Laden der Konfiguration: {e}")

    def save_config(self):
        """Speichert Konfiguration in JSON-Datei"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
            messagebox.showinfo("Erfolg", "Konfiguration gespeichert!")
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Speichern: {e}")

    # -------------------- GUI --------------------

    def create_gui(self):
        """Erstellt die Benutzeroberfläche"""
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Tab 1: Pool-Konfiguration
        self.create_pool_config_tab(notebook)
        # Tab 2: Schichtplanung
        self.create_shift_planning_tab(notebook)

    def create_pool_config_tab(self, notebook):
        """Erstellt Tab für Pool-Konfiguration"""
        pool_frame = ttk.Frame(notebook)
        notebook.add(pool_frame, text="Pool-Konfiguration")

        # Scrollbarer Frame
        canvas = tk.Canvas(pool_frame)
        scrollbar_pool = ttk.Scrollbar(pool_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar_pool.set)

        # Pool A
        ttk.Label(scrollable_frame, text="Pool A - Vormittag (können alles):").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.pool_vm_alle_entry = tk.Entry(scrollable_frame, width=60)
        self.pool_vm_alle_entry.grid(row=0, column=1, padx=5, pady=5)
        self.pool_vm_alle_entry.insert(0, ",".join(self.config["pool_vm_alle"]))

        # Pool B
        ttk.Label(scrollable_frame, text="Pool B - Vormittag (brauchen Unterstützung):").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.pool_vm_teilweise_entry = tk.Entry(scrollable_frame, width=60)
        self.pool_vm_teilweise_entry.grid(row=1, column=1, padx=5, pady=5)
        self.pool_vm_teilweise_entry.insert(0, ",".join(self.config["pool_vm_teilweise"]))

        # Pool C
        ttk.Label(scrollable_frame, text="Pool C - Support für Pool B:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.pool_vm_support_entry = tk.Entry(scrollable_frame, width=60)
        self.pool_vm_support_entry.grid(row=2, column=1, padx=5, pady=5)
        self.pool_vm_support_entry.insert(0, ",".join(self.config["pool_vm_support"]))

        # Pool D
        ttk.Label(scrollable_frame, text="Pool D - Nachmittag (können alles):").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.pool_nm_alle_entry = tk.Entry(scrollable_frame, width=60)
        self.pool_nm_alle_entry.grid(row=3, column=1, padx=5, pady=5)
        self.pool_nm_alle_entry.insert(0, ",".join(self.config["pool_nm_alle"]))

        # Pool E
        ttk.Label(scrollable_frame, text="Pool E - Freitags NICHT verfügbar:", font=('TkDefaultFont', 9, 'bold'), foreground='red').grid(row=4, column=0, sticky="w", padx=5, pady=5)
        self.pool_freitag_abwesend_entry = tk.Entry(scrollable_frame, width=60)
        self.pool_freitag_abwesend_entry.grid(row=4, column=1, padx=5, pady=5)
        self.pool_freitag_abwesend_entry.insert(0, ",".join(self.config["pool_freitag_abwesend"]))

        # Speichern Button
        ttk.Button(scrollable_frame, text="Pools speichern", command=self.save_pools).grid(row=5, column=1, pady=20)

        info_text = """
Hinweise zur Pool-Konfiguration:
• Mitarbeiter mit Komma trennen (z.B. RR,AN,MH,TL)
• Reihenfolge wird bei der Planung berücksichtigt
• Pool A: Mitarbeiter die vormittags alles alleine können
• Pool B: Mitarbeiter die vormittags Unterstützung brauchen
• Pool C: Mitarbeiter die Pool B unterstützen können  
• Pool D: Mitarbeiter für Nachmittagsdienst
• Pool E: Mitarbeiter die FREITAGS NICHT eingeteilt werden dürfen
  (z.B. Homeoffice, Teilzeit, externe Termine)
"""
        ttk.Label(scrollable_frame, text=info_text, justify="left", font=('TkDefaultFont', 9)).grid(row=6, column=0, columnspan=2, padx=5, pady=10, sticky="w")

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar_pool.pack(side="right", fill="y")

    def create_shift_planning_tab(self, notebook):
        """Erstellt Tab für Schichtplanung"""
        planning_frame = ttk.Frame(notebook)
        notebook.add(planning_frame, text="Schichtplanung")

        # Oben: linke & rechte Spalte (Bedienelemente), unten: Ergebnisliste
        top_frame = ttk.Frame(planning_frame)
        top_frame.pack(side="top", fill="x", expand=False)

        left_frame = ttk.Frame(top_frame)
        left_frame.pack(side="left", fill="y", padx=(0, 10))

        right_frame = ttk.Frame(top_frame)
        right_frame.pack(side="left", fill="both", expand=True, padx=(10, 0))

        # Linke Seite – Grundeinstellungen
        ttk.Label(left_frame, text="Grundeinstellungen", font=('TkDefaultFont', 11, 'bold'), foreground='pink').grid(row=0, column=0, columnspan=2, pady=(0, 10), sticky="w")

        ttk.Label(left_frame, text="Startdatum (Montag, TT.MM.YYYY):").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.start_date_entry = tk.Entry(left_frame, width=20)
        self.start_date_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(left_frame, text="1. Tag Vormittag:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.first_vm_entry = tk.Entry(left_frame, width=20)
        self.first_vm_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(left_frame, text="1. Tag Nachmittag:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.first_nm_entry = tk.Entry(left_frame, width=20)
        self.first_nm_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(left_frame, text="1. Tag Support:").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        self.first_support_entry = tk.Entry(left_frame, width=20)
        self.first_support_entry.grid(row=4, column=1, padx=5, pady=5, sticky="w")

        button_frame = ttk.Frame(left_frame)
        button_frame.grid(row=5, column=0, columnspan=2, pady=20)
        ttk.Button(button_frame, text="Planung erstellen", command=self.create_planning).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Excel exportieren", command=self.export_excel).pack(side="left", padx=5)

        # Rechte Seite – Abwesenheiten
        ttk.Label(right_frame, text="Abwesenheiten verwalten", font=('TkDefaultFont', 11, 'bold')).grid(row=0, column=0, columnspan=3, pady=(0, 10), sticky="w")

        ttk.Label(right_frame, text="Mitarbeiter:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.employee_var = tk.StringVar()
        self.employee_combo = ttk.Combobox(right_frame, textvariable=self.employee_var, width=18)
        self.employee_combo.grid(row=1, column=1, padx=5, pady=2, sticky="w")
        ttk.Label(right_frame, text="(Mehrere: IL,AN,RR)", font=('TkDefaultFont', 8), foreground='gray').grid(row=1, column=2, sticky="w", padx=2, pady=2)

        ttk.Label(right_frame, text="Tag:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        self.day_var = tk.StringVar()
        # 12 Tage (Mo-Sa, Mo-Sa)
        self.day_combo = ttk.Combobox(right_frame, textvariable=self.day_var, width=18, values=[f"Tag {i + 1}" for i in range(12)])
        self.day_combo.grid(row=2, column=1, padx=5, pady=2, sticky="w")

        ttk.Button(right_frame, text="➕ Hinzufügen", command=self.add_absence).grid(row=3, column=0, padx=5, pady=5, sticky="w")
        ttk.Button(right_frame, text="🗑️ Entfernen", command=self.remove_absence).grid(row=3, column=1, padx=5, pady=5, sticky="w")
        ttk.Button(right_frame, text="Pools laden", command=self.update_employee_list).grid(row=3, column=2, padx=5, pady=5, sticky="w")

        ttk.Label(right_frame, text="Eingetragene Abwesenheiten:").grid(row=4, column=0, columnspan=3, sticky="w", padx=5, pady=(15, 5))

        # Abwesenheitenliste: mindestens 26 Zeilen
        absence_columns = ("Tag", "Mitarbeiter")
        self.absence_tree = ttk.Treeview(right_frame, columns=absence_columns, show="headings", height=26)
        self.absence_tree.heading("Tag", text="Tag")
        self.absence_tree.heading("Mitarbeiter", text="Mitarbeiter")
        self.absence_tree.column("Tag", width=90, stretch=False)
        self.absence_tree.column("Mitarbeiter", width=140, stretch=True)
        self.absence_tree.grid(row=5, column=0, columnspan=3, padx=5, pady=5, sticky="nsew")

        absence_scrollbar = ttk.Scrollbar(right_frame, orient="vertical", command=self.absence_tree.yview)
        absence_scrollbar.grid(row=5, column=3, sticky="ns", pady=5)
        self.absence_tree.configure(yscrollcommand=absence_scrollbar.set)

        right_frame.grid_rowconfigure(5, weight=1)
        right_frame.grid_columnconfigure(2, weight=1)

        # Ergebnisbereich
        result_frame = ttk.Frame(planning_frame)
        result_frame.pack(side="top", fill="both", expand=True, pady=(20, 0))

        ttk.Label(result_frame, text="Planungsergebnis:", font=('TkDefaultFont', 11, 'bold')).pack(anchor="w", pady=(0, 5))

        # Spaltenreihenfolge: Datum, Wochentag, Vormittag, Nachmittag, Support
        columns = ("Datum", "Wochentag", "Vormittag", "Nachmittag", "Support")
        self.result_tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=12)
        for c in columns:
            self.result_tree.heading(c, text=c)
        self.result_tree.column("Datum", width=120, stretch=False)
        self.result_tree.column("Wochentag", width=160, stretch=False)
        self.result_tree.column("Vormittag", width=180, stretch=True)
        self.result_tree.column("Nachmittag", width=180, stretch=True)
        self.result_tree.column("Support", width=180, stretch=True)

        self.result_tree.pack(side="left", fill="both", expand=True)

        scrollbar_y = ttk.Scrollbar(result_frame, orient="vertical", command=self.result_tree.yview)
        scrollbar_y.pack(side="right", fill="y")
        self.result_tree.configure(yscrollcommand=scrollbar_y.set)

        self.planning_result = []
        self.absences = {}  # {tag_nr: [mitarbeiter_liste]}

    # -------------------- Abwesenheiten --------------------

    def update_employee_list(self):
        """Aktualisiert die Mitarbeiterliste in der Combobox"""
        all_employees = set()
        for pool in [self.config["pool_vm_alle"], self.config["pool_vm_teilweise"], self.config["pool_vm_support"], self.config["pool_nm_alle"]]:
            all_employees.update(pool)
        self.employee_combo['values'] = sorted(list(all_employees))
        messagebox.showinfo("Info", f"Mitarbeiterliste aktualisiert: {len(all_employees)} Mitarbeiter gefunden")

    def add_absence(self):
        """Fügt eine Abwesenheit hinzu (unterstützt mehrere, komma-getrennte Kürzel)."""
        employee = self.employee_var.get().strip()
        day_str = self.day_var.get().strip()
        if not employee or not day_str:
            messagebox.showwarning("Warnung", "Bitte Mitarbeiter und Tag auswählen!")
            return
        try:
            day_nr = int(day_str.split()[1]) - 1  # "Tag 1" -> 0
        except (ValueError, IndexError):
            messagebox.showerror("Fehler", "Ungültiger Tag!")
            return
        employees = [e.strip() for e in employee.split(",") if e.strip()]
        if day_nr not in self.absences:
            self.absences[day_nr] = []
        for emp in employees:
            if emp not in self.absences[day_nr]:
                self.absences[day_nr].append(emp)
        self.update_absence_display()

        # Optional: direkt zum nächsten Tag springen
        current_day = day_nr + 1
        if current_day < 12:
            self.day_var.set(f"Tag {current_day + 1}")

    def remove_absence(self):
        """Entfernt eine Abwesenheit"""
        selected = self.absence_tree.selection()
        if not selected:
            messagebox.showwarning("Warnung", "Bitte einen Eintrag auswählen!")
            return
        for item in selected:
            values = self.absence_tree.item(item, 'values')
            day_str, employee = values[0], values[1]
            day_nr = int(day_str.split()[1]) - 1
            if day_nr in self.absences and employee in self.absences[day_nr]:
                self.absences[day_nr].remove(employee)
                if not self.absences[day_nr]:
                    del self.absences[day_nr]
        self.update_absence_display()
        messagebox.showinfo("Erfolg", "Abwesenheit(en) entfernt")

    def update_absence_display(self):
        """Aktualisiert die Anzeige der Abwesenheiten"""
        for item in self.absence_tree.get_children():
            self.absence_tree.delete(item)
        for day_nr in sorted(self.absences.keys()):
            for employee in sorted(self.absences[day_nr]):
                self.absence_tree.insert("", tk.END, values=(f"Tag {day_nr + 1}", employee))

    def save_pools(self):
        """Speichert die Pool-Konfiguration"""
        try:
            self.config["pool_vm_alle"] = [x.strip() for x in self.pool_vm_alle_entry.get().split(",") if x.strip()]
            self.config["pool_vm_teilweise"] = [x.strip() for x in self.pool_vm_teilweise_entry.get().split(",") if x.strip()]
            self.config["pool_vm_support"] = [x.strip() for x in self.pool_vm_support_entry.get().split(",") if x.strip()]
            self.config["pool_nm_alle"] = [x.strip() for x in self.pool_nm_alle_entry.get().split(",") if x.strip()]
            self.config["pool_freitag_abwesend"] = [x.strip() for x in self.pool_freitag_abwesend_entry.get().split(",") if x.strip()]
            self.save_config()
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Speichern der Pools: {e}")

    def parse_absent_employees(self):
        """Gibt die Abwesenheitsstruktur zurück"""
        return self.absences

    # -------------------- Planung --------------------

    def create_planning(self):
        """Erstellt die Schichtplanung für 2 Wochen Mo-Sa (12 Tage)"""
        try:
            start_date_str = self.start_date_entry.get().strip()
            if not start_date_str:
                messagebox.showerror("Fehler", "Bitte Startdatum eingeben!")
                return
            start_date = datetime.strptime(start_date_str, "%d.%m.%Y")
            if start_date.weekday() != 0:  # Montag
                messagebox.showerror("Fehler", "Startdatum muss ein Montag sein!")
                return

            first_vm = self.first_vm_entry.get().strip()
            first_nm = self.first_nm_entry.get().strip()
            first_support = self.first_support_entry.get().strip()
            if not all([first_vm, first_nm]):
                messagebox.showerror("Fehler", "Bitte mindestens Vormittag und Nachmittag für den ersten Tag ausfüllen!")
                return

            absent_by_day = self.parse_absent_employees()
            self.planning_result = []

            pool_positions = {"vm_alle": 0, "vm_teilweise": 0, "vm_support": 0, "nm_alle": 0}

            if first_vm and first_vm in self.config["pool_vm_alle"]:
                vm_start_idx = self.config["pool_vm_alle"].index(first_vm)
                pool_positions["vm_alle"] = (vm_start_idx + 1) % len(self.config["pool_vm_alle"])
            if first_nm and first_nm in self.config["pool_nm_alle"]:
                nm_start_idx = self.config["pool_nm_alle"].index(first_nm)
                pool_positions["nm_alle"] = (nm_start_idx + 1) % len(self.config["pool_nm_alle"])
            if first_support and first_support in self.config["pool_vm_support"]:
                support_start_idx = self.config["pool_vm_support"].index(first_support)
                pool_positions["vm_support"] = (support_start_idx + 1) % len(self.config["pool_vm_support"])
            if first_vm and first_vm in self.config["pool_vm_teilweise"]:
                vm_b_idx = self.config["pool_vm_teilweise"].index(first_vm)
                pool_positions["vm_teilweise"] = (vm_b_idx + 1) % len(self.config["pool_vm_teilweise"])

            # 12 Tage: Mo-Sa, Mo-Sa (2 Wochen ohne Sonntag)
            for tag_nr in range(12):
                # Wochentag innerhalb der Woche (0=Mo .. 5=Sa) und Wochenoffset (für Sonntage)
                weekday_in_week = tag_nr % 6
                week_number = tag_nr // 6
                current_date = start_date + timedelta(days=tag_nr + week_number)

                is_friday = weekday_in_week == 4
                is_saturday = weekday_in_week == 5

                weekday_names = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"]
                weekday_german = weekday_names[weekday_in_week]

                absent_today = absent_by_day.get(tag_nr, []).copy()
                if is_friday:
                    # Freitags zusätzlich die generellen "Freitag nicht verfügbar"
                    absent_today = list(set(absent_today + self.config["pool_freitag_abwesend"]))

                # Samstag: keine Schichten
                if is_saturday:
                    vm_employee = ""
                    nm_employee = ""
                    support_employee = ""
                elif tag_nr == 0:
                    vm_employee = first_vm
                    nm_employee = first_nm
                    support_employee = first_support if first_support else ""
                else:
                    yesterday_nm = self.planning_result[tag_nr - 1]["Nachmittag"] if tag_nr > 0 else None
                    forbidden_vm = [yesterday_nm] if yesterday_nm else []
                    used_today = []
                    vm_employee = None
                    support_employee = ""

                    # Vormittag aus Pool A
                    if self.config["pool_vm_alle"]:
                        start_pos = pool_positions["vm_alle"]
                        pool_size = len(self.config["pool_vm_alle"])
                        for i in range(pool_size):
                            candidate_idx = (start_pos + i) % pool_size
                            candidate = self.config["pool_vm_alle"][candidate_idx]
                            if (candidate not in absent_today and candidate not in forbidden_vm):
                                vm_employee = candidate
                                used_today.append(vm_employee)
                                pool_positions["vm_alle"] = (candidate_idx + 1) % pool_size

                                # Support nötig?
                                if candidate in self.config["pool_vm_teilweise"] and self.config["pool_vm_support"]:
                                    start_pos_c = pool_positions["vm_support"]
                                    pool_c_size = len(self.config["pool_vm_support"])
                                    for j in range(pool_c_size):
                                        support_idx = (start_pos_c + j) % pool_c_size
                                        candidate_c = self.config["pool_vm_support"][support_idx]
                                        if (candidate_c not in absent_today and candidate_c != candidate and candidate_c not in forbidden_vm):
                                            support_employee = candidate_c
                                            used_today.append(support_employee)
                                            pool_positions["vm_support"] = (support_idx + 1) % pool_c_size
                                            break
                                break

                    # Nachmittag aus Pool D
                    nm_employee = None
                    if self.config["pool_nm_alle"]:
                        start_pos = pool_positions["nm_alle"]
                        pool_size = len(self.config["pool_nm_alle"])
                        for i in range(pool_size):
                            candidate_idx = (start_pos + i) % pool_size
                            candidate = self.config["pool_nm_alle"][candidate_idx]
                            if (candidate not in absent_today and candidate not in used_today):
                                nm_employee = candidate
                                pool_positions["nm_alle"] = (candidate_idx + 1) % pool_size
                                break

                self.planning_result.append({
                    "Datum": current_date.strftime("%d.%m.%Y"),
                    "Wochentag": weekday_german,
                    "Vormittag": vm_employee if not is_saturday else "",
                    "Nachmittag": nm_employee if not is_saturday else "",
                    "Support": support_employee if not is_saturday else ""
                })

            self.display_results()
        except ValueError:
            messagebox.showerror("Fehler", "Ungültiges Datumsformat! Bitte TT.MM.YYYY verwenden.")
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler bei der Planung: {e}")

    # -------------------- Anzeige & Export --------------------

    def display_results(self):
        """Zeigt die Planungsergebnisse in der Treeview an"""
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        for row in self.planning_result:
            self.result_tree.insert(
                "", tk.END,
                values=(row["Datum"], row["Wochentag"], row["Vormittag"], row["Nachmittag"], row["Support"])
            )

    def export_excel(self):
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

        # Daten in gewünschter Reihenfolge aufbauen
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

        df = pd.DataFrame(rows, columns=["Datum", "Wochentag", "Vormittag", "Nachmittag", "Support"])

        # CSV-Fallback ohne Formatierung
        if filename.lower().endswith(".csv"):
            df2 = df.copy()
            if isinstance(df2.loc[0, "Datum"], datetime):
                df2["Datum"] = df2["Datum"].dt.strftime("%d.%m.%Y")
            df2.to_csv(filename, index=False, encoding="utf-8-sig")
            messagebox.showinfo("Erfolg", f"CSV-Datei erfolgreich gespeichert: {filename}")
            return

        # Excel mit Formatierung
        try:
            from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
            from openpyxl.utils import get_column_letter

            # Farbpalette (an Screenshot angelehnt)
            GREEN  = PatternFill(start_color="A9D18E", end_color="A9D18E", fill_type="solid")  # grün für Vormittag & Nachmittag
            YELLOW = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")  # gelb für Support
            GREY   = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid")  # Trennzeile

            thin = Border(
                left=Side(style="thin"), right=Side(style="thin"),
                top=Side(style="thin"), bottom=Side(style="thin")
            )

            with pd.ExcelWriter(filename, engine="openpyxl", date_format="DD.MM.YYYY") as writer:
                df.to_excel(writer, sheet_name="Notdienst_Planung", index=False)
                ws = writer.sheets["Notdienst_Planung"]

                # Spaltenbreiten
                widths = [12, 14, 14, 14, 14]  # A..E
                for idx, w in enumerate(widths, start=1):
                    ws.column_dimensions[get_column_letter(idx)].width = w

                # Header fett & zentriert
                for c in ws[1]:
                    c.font = Font(bold=True)
                    c.alignment = Alignment(horizontal="center")
                    c.border = thin

                max_row = ws.max_row
                max_col = ws.max_column

                # Datenzellen: Rahmen + Füllfarben (3&4 grün, 5 gelb; 1&2 weiß)
                for r in range(2, max_row + 1):
                    for c in range(1, max_col + 1):
                        cell = ws.cell(row=r, column=c)
                        cell.border = thin
                        cell.alignment = Alignment(horizontal=("left" if c in (1, 2) else "center"))
                        if c == 3:        # Vormittag
                            cell.fill = GREEN
                        elif c == 4:      # Nachmittag
                            cell.fill = GREEN
                        elif c == 5:      # Support
                            cell.fill = YELLOW
                        # Spalten 1 & 2 bleiben ohne Fill (weiß)

                # Trennzeilen NACH jedem Samstag (von unten nach oben einfügen)
                saturday_rows = []
                for i, wd in enumerate(df["Wochentag"].tolist(), start=2):  # +1 Header, +1 weil 1-basiert
                    if wd == "Samstag":
                        saturday_rows.append(i)

                for row_idx in sorted(saturday_rows, reverse=True):
                    ins = row_idx + 1  # Zeile nach Samstag
                    ws.insert_rows(ins)
                    for c in range(1, max_col + 1):
                        cell = ws.cell(row=ins, column=c)
                        cell.fill = GREY
                        cell.border = thin

            messagebox.showinfo("Erfolg", f"Excel-Datei erfolgreich gespeichert: {filename}")

        except ImportError:
            csv_filename = filename.rsplit(".", 1)[0] + ".csv"
            df2 = df.copy()
            if isinstance(df2.loc[0, "Datum"], datetime):
                df2["Datum"] = df2["Datum"].dt.strftime("%d.%m.%Y")
            df2.to_csv(csv_filename, index=False, encoding="utf-8-sig")
            messagebox.showinfo(
                "Info",
                "openpyxl ist nicht installiert. CSV gespeichert: "
                f"{csv_filename}\nFür formatiertes Excel bitte installieren: pip install openpyxl"
            )
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Export: {e}")


def main():
    root = tk.Tk()
    app = ShiftPlanner(root)
    root.mainloop()


if __name__ == "__main__":
    main()
