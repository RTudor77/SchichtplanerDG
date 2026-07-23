"""Verifikations-Tests für SchichtplanerDG.

Läuft in einem Temp-Verzeichnis, damit die echten Dateien
(shift_config.json / shift_history.json) unangetastet bleiben.

Aufruf:  py test_SchichtplanerDG.py
"""
import json
import os
import sys
import tempfile
import tkinter as tk
from datetime import datetime

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, PROJECT_DIR)

# In Temp-Verzeichnis wechseln, BEVOR das Modul geladen wird (relative Dateipfade!)
WORK_DIR = tempfile.mkdtemp(prefix="schichtplaner_test_")
os.chdir(WORK_DIR)

import SchichtplanerDG as mod


class _SilentBox:
    """Ersetzt messagebox im Test (keine Dialoge, Fehler werden geloggt)."""
    def showerror(self, *a, **k):
        print("    [messagebox.showerror]", a)
    def showwarning(self, *a, **k):
        print("    [messagebox.showwarning]", a)
    def askyesno(self, *a, **k):
        return True
    def __getattr__(self, name):
        return lambda *a, **k: None


mod.messagebox = _SilentBox()

FAILURES = []


def check(name, cond, info=""):
    status = "OK  " if cond else "FAIL"
    extra = f" — {info}" if (info and not cond) else ""
    print(f"[{status}] {name}{extra}")
    if not cond:
        FAILURES.append(name)


def fresh_app(config, history=None):
    if os.path.exists(mod.HISTORY_FILE):
        os.remove(mod.HISTORY_FILE)
    if os.path.exists(mod.ABSENCES_FILE):
        os.remove(mod.ABSENCES_FILE)
    if history is not None:
        with open(mod.HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(history, f)
    with open(mod.CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f)
    root = tk.Tk()
    root.withdraw()
    app = mod.ShiftPlanner(root)
    return root, app


BASE_POOLS = {
    "pool_vm_alle": ["AA", "BB", "CC", "DD", "EE", "FF", "GG"],
    "pool_vm_teilweise": [],
    "pool_vm_support": ["AA", "BB", "CC", "DD", "EE", "FF", "GG"],
    "pool_nm_alle": ["AA", "BB", "CC", "DD", "EE", "FF", "GG"],
    "pool_freitag_abwesend": [],
    "pool_mo_mi_abwesend": [],
    "feiertage": []
}

REAL_POOLS = {
    "pool_vm_alle": ["MH", "RI", "TR", "JB", "FA", "RR", "IL"],
    "pool_vm_teilweise": ["MH", "RI"],
    "pool_vm_support": ["IL", "RR", "FA", "JB", "AN", "TR"],
    "pool_nm_alle": ["IL", "RR", "FA", "JB", "AN", "TR", "RI", "MH", "CA"],
    "pool_freitag_abwesend": [],
    "pool_mo_mi_abwesend": ["CA"],
    "feiertage": []
}


def week_loads(entries, employees, week):
    rows = entries[week * 5:(week + 1) * 5]
    loads = {ma: 0 for ma in employees}
    for e in rows:
        vm = e["Vormittag"]
        for st in ("Vormittag", "Nachmittag", "Support"):
            ma = e[st]
            if ma and not (st == "Support" and ma == vm) and ma in loads:
                loads[ma] += 1
    return loads


# ---------------------------------------------------------------- Test 1
print("\n--- Test 1: Statistik-Doppelzählung behoben ---")
root, app = fresh_app(BASE_POOLS)
app.saved_plans = [{
    "start_date": "02.02.2026", "end_date": "13.02.2026",
    "entries": [
        {"Datum": "02.02.2026", "Wochentag": "Montag",
         "Vormittag": "AA", "Nachmittag": "BB", "Support": "AA"},   # Anzeige-Duplikat
        {"Datum": "03.02.2026", "Wochentag": "Dienstag",
         "Vormittag": "CC", "Nachmittag": "DD", "Support": "EE"},   # echter Support
    ]}]
stats = app._calculate_statistics("", "")
check("Support==VM zählt nicht als Support-Einsatz",
      stats["AA"]["Support"] == 0 and stats["AA"]["VM"] == 1, str(stats.get("AA")))
check("Echter Support zählt weiterhin", stats["EE"]["Support"] == 1, str(stats.get("EE")))
counts = app._load_recent_history_counts(-1)
check("Fairness-Historie nutzt dieselbe Dedup-Regel",
      counts["AA"]["Support"] == 0 and counts["EE"]["Support"] == 1, str(counts))
root.destroy()

# ---------------------------------------------------------------- Test 2
print("\n--- Test 2: Halbtägige Abwesenheiten (Termine) ---")
root, app = fresh_app(BASE_POOLS)
app.absences = [
    {"ma": "AA", "von": datetime(2026, 2, 3), "bis": datetime(2026, 2, 3), "zeit": "vm"},
    {"ma": "BB", "von": datetime(2026, 2, 2), "bis": datetime(2026, 2, 6), "zeit": "nm"},
    {"ma": "CC", "von": datetime(2026, 2, 3), "bis": datetime(2026, 2, 4), "zeit": "ganz"},
]
vm, nm = app._absence_sets_for_date(datetime(2026, 2, 3))
check("'Nur vormittags' blockiert VM/Support, nicht NM", "AA" in vm and "AA" not in nm)
check("'Nur nachmittags' blockiert NM, nicht VM", "BB" in nm and "BB" not in vm)
check("'Ganztags' blockiert beides", "CC" in vm and "CC" in nm)
vm2, nm2 = app._absence_sets_for_date(datetime(2026, 2, 5))
check("Zeitraum-Ende wird respektiert",
      "CC" not in vm2 and "AA" not in vm2 and "BB" in nm2)
root.destroy()

# ---------------------------------------------------------------- Test 3
print("\n--- Test 3: Fair-Modus — Wochen-Quote, leerer Tag 1, Abwesenheiten ---")
root, app = fresh_app(BASE_POOLS)
app.start_date_entry.delete(0, tk.END)
app.start_date_entry.insert(0, "02.02.2026")  # Montag
app.fair_mode_var.set(True)
app.absences = [
    {"ma": "AA", "von": datetime(2026, 2, 3), "bis": datetime(2026, 2, 3), "zeit": "vm"},
    {"ma": "BB", "von": datetime(2026, 2, 9), "bis": datetime(2026, 2, 13), "zeit": "ganz"},
]
app.create_planning()
check("Plan mit leerem Tag 1 erzeugt (10 Tage)", len(app.planning_result) == 10,
      f"len={len(app.planning_result)}")
fehler = [v for v in app._validate_current_plan() if v["schwere"] == "fehler"]
check("Keine Regel-Fehler im erzeugten Plan", not fehler, str(fehler))
row_di = app.planning_result[1]
check("Termin vormittags: kein VM/Support am 03.02.",
      row_di["Vormittag"] != "AA" and row_di["Support"] != "AA", str(row_di))
w2_rows = app.planning_result[5:]
bb_used = any("BB" in (r["Vormittag"], r["Nachmittag"], r["Support"]) for r in w2_rows)
check("Ganztags-Abwesenheit komplette Woche 2 respektiert", not bb_used)
emps = BASE_POOLS["pool_vm_alle"]
l1 = week_loads(app.planning_result, emps, 0)
l2 = week_loads(app.planning_result, [m for m in emps if m != "BB"], 1)
s1 = max(l1.values()) - min(l1.values())
s2 = max(l2.values()) - min(l2.values())
check("Woche 1: Spread <= 1", s1 <= 1, f"loads={l1}")
check("Woche 2: Spread <= 1 (ohne BB)", s2 <= 1, f"loads={l2}")
root.destroy()

# ---------------------------------------------------------------- Test 4
print("\n--- Test 4: Plan-Prüfung erkennt Regelverstöße ---")
root, app = fresh_app({**BASE_POOLS, "pool_vm_teilweise": ["BB"]})
app.absences = [
    {"ma": "DD", "von": datetime(2026, 2, 2), "bis": datetime(2026, 2, 2), "zeit": "ganz"},
]
app.planning_result = [
    {"Datum": "02.02.2026", "Wochentag": "Montag",
     "Vormittag": "DD", "Nachmittag": "DD", "Support": "XX"},   # abwesend + doppelt + unbekannt
    {"Datum": "03.02.2026", "Wochentag": "Dienstag",
     "Vormittag": "DD", "Nachmittag": "EE", "Support": ""},      # Ruhezeit (gestern NM)
    {"Datum": "04.02.2026", "Wochentag": "Mittwoch",
     "Vormittag": "BB", "Nachmittag": "FF", "Support": ""},      # Pool B ohne Support
]
viol = app._validate_current_plan()
check("Erkennt abwesenden MA",
      any("abwesend" in v["text"] and v["row"] == 0 for v in viol))
check("Erkennt VM==NM am selben Tag", any("VM und NM" in v["text"] for v in viol))
check("Erkennt unbekanntes Kürzel (Tippfehler)",
      any("Unbekanntes Kürzel" in v["text"] for v in viol))
check("Erkennt Ruhezeit-Verstoß", any("Ruhezeit" in v["text"] for v in viol))
check("Erkennt Pool B ohne Support", any("Pool B" in v["text"] for v in viol))
root.destroy()

# ---------------------------------------------------------------- Test 5
print("\n--- Test 5: Echte Pool-Konfiguration (Smoke-Test) ---")
root, app = fresh_app(REAL_POOLS)
app.start_date_entry.delete(0, tk.END)
app.start_date_entry.insert(0, "02.02.2026")
app.fair_mode_var.set(True)
app.absences = [
    {"ma": "RR", "von": datetime(2026, 2, 4), "bis": datetime(2026, 2, 4), "zeit": "vm"},
    {"ma": "FA", "von": datetime(2026, 2, 10), "bis": datetime(2026, 2, 11), "zeit": "nm"},
]
app.create_planning()
fehler = [v for v in app._validate_current_plan() if v["schwere"] == "fehler"]
check("Keine Regel-Fehler mit echten Pools", not fehler, str(fehler))
check("Alle Tage besetzt",
      all(r["Vormittag"] and r["Nachmittag"] for r in app.planning_result))
all_emps = sorted(set(sum([REAL_POOLS[k] for k in
                           ("pool_vm_alle", "pool_vm_support", "pool_nm_alle")], [])))
l1 = week_loads(app.planning_result, all_emps, 0)
l2 = week_loads(app.planning_result, all_emps, 1)
print(f"    Wochenlast W1: {l1}")
print(f"    Wochenlast W2: {l2}")
check("W1: Spread <= 2 (Pools ungleich besetzt)",
      max(l1.values()) - min(l1.values()) <= 2, str(l1))
check("W2: Spread <= 2 (Pools ungleich besetzt)",
      max(l2.values()) - min(l2.values()) <= 2, str(l2))
root.destroy()

# ---------------------------------------------------------------- Test 6
print("\n--- Test 6: Klassischer Modus weiterhin funktionsfähig ---")
root, app = fresh_app(REAL_POOLS)
app.start_date_entry.delete(0, tk.END)
app.start_date_entry.insert(0, "02.02.2026")
app.fair_mode_var.set(False)
app.first_vm_entry.insert(0, "MH")
app.first_nm_entry.insert(0, "IL")
app.first_support_entry.insert(0, "RR")
app.create_planning()
check("Klassik: 10 Tage erzeugt", len(app.planning_result) == 10,
      f"len={len(app.planning_result)}")
check("Klassik: Tag 1 wie vorgegeben",
      app.planning_result[0]["Vormittag"] == "MH"
      and app.planning_result[0]["Nachmittag"] == "IL")
root.destroy()

# ---------------------------------------------------------------- Test 7
print("\n--- Test 7: Abwesenheits-Eingabe über GUI-Funktionen ---")
root, app = fresh_app(REAL_POOLS)
app.employee_var.set("RR,IL")
app.absence_from_var.set("03.02.2026")
app.absence_to_var.set("05.02.2026")
app.absence_scope_var.set("Nur nachmittags")
app.add_absence()
check("Zwei MA in einem Rutsch erfasst", len(app.absences) == 2,
      f"len={len(app.absences)}")
check("Zeitraum + Zeitfenster korrekt übernommen",
      app.absences[0]["zeit"] == "nm" and app.absences[0]["bis"].day == 5,
      str(app.absences))
app.absence_from_var.set("06.02.2026")
app.absence_to_var.set("")
app.absence_scope_var.set("Ganztags")
app.employee_var.set("RR")
app.add_absence()
check("'Bis' leer = Einzeltag", app.absences[-1]["von"] == app.absences[-1]["bis"],
      str(app.absences[-1]))
root.destroy()

# ---------------------------------------------------------------- Test 8
print("\n--- Test 8: Vortags-Regel (Abwesenheit am Folgetag => kein NM am Vortag) ---")
root, app = fresh_app(BASE_POOLS)
app.start_date_entry.delete(0, tk.END)
app.start_date_entry.insert(0, "02.02.2026")  # Montag
app.fair_mode_var.set(True)
app.absences = [
    # AA ganztags abwesend am Dienstag 03.02. => Montag 02.02. kein NM für AA
    {"ma": "AA", "von": datetime(2026, 2, 3), "bis": datetime(2026, 2, 3), "zeit": "ganz"},
    # BB vormittags abwesend am Montag 09.02. (W2) => Freitag 06.02. kein NM für BB
    {"ma": "BB", "von": datetime(2026, 2, 9), "bis": datetime(2026, 2, 9), "zeit": "vm"},
    # CC NUR NACHMITTAGS abwesend am Donnerstag 05.02. => Mittwoch-NM bleibt ERLAUBT
    {"ma": "CC", "von": datetime(2026, 2, 5), "bis": datetime(2026, 2, 5), "zeit": "nm"},
]
app.create_planning()
check("Plan erzeugt", len(app.planning_result) == 10)
fehler = [v for v in app._validate_current_plan() if v["schwere"] == "fehler"]
check("Keine Regel-Fehler", not fehler, str(fehler))
check("Mo 02.02.: AA nicht NM (Di ganztags abwesend)",
      app.planning_result[0]["Nachmittag"] != "AA", str(app.planning_result[0]))
check("Fr 06.02.: BB nicht NM (Mo vormittags abwesend)",
      app.planning_result[4]["Nachmittag"] != "BB", str(app.planning_result[4]))

# Prüf-Engine: handgebauter Verstoß wird konkret gemeldet
app.planning_result[0]["Nachmittag"] = "AA"
viol = app._validate_current_plan()
check("Prüfung meldet Vortags-Verstoß konkret",
      any("Folgetag" in v["text"] and v["row"] == 0 for v in viol),
      " | ".join(v["text"] for v in viol))

# Gegentest: 'nur nachmittags' am Folgetag blockiert den Vortag NICHT
app.planning_result[2]["Nachmittag"] = "CC"  # Mittwoch 04.02.
viol = app._validate_current_plan()
check("'Nur nachmittags' am Folgetag blockiert Vortag-NM nicht",
      not any("Folgetag" in v["text"] and v["row"] == 2 for v in viol),
      " | ".join(v["text"] for v in viol))
root.destroy()

# ---------------------------------------------------------------- Test 9
print("\n--- Test 9: Abwesenheiten überleben Neustart (Persistierung) ---")
root, app = fresh_app(REAL_POOLS)
app.employee_var.set("RR")
app.absence_from_var.set("03.02.2026")
app.absence_to_var.set("05.02.2026")
app.absence_scope_var.set("Nur vormittags")
app.add_absence()
root.destroy()

# Neue Instanz simuliert Programm-Neustart (Config/History NICHT löschen!)
root2 = tk.Tk()
root2.withdraw()
app2 = mod.ShiftPlanner(root2)
check("Abwesenheit nach Neustart wieder geladen", len(app2.absences) == 1,
      str(app2.absences))
if app2.absences:
    a = app2.absences[0]
    check("Inhalt korrekt (MA, Zeitraum, Zeitfenster)",
          a["ma"] == "RR" and a["zeit"] == "vm"
          and a["von"] == datetime(2026, 2, 3) and a["bis"] == datetime(2026, 2, 5),
          str(a))
    check("Anzeige-Liste beim Start befüllt",
          len(app2.absence_tree.get_children()) == 1)
root2.destroy()

# ---------------------------------------------------------------- Test 10
print("\n--- Test 10: Pool F = max. 1 Einsatz pro Woche ---")
POOLS_F = {
    "pool_vm_alle": ["CC", "DD", "EE"],
    "pool_vm_teilweise": [],
    "pool_vm_support": ["CC", "DD", "EE"],
    "pool_nm_alle": ["AA", "BB", "CC", "DD", "EE"],
    "pool_freitag_abwesend": [],
    "pool_mo_mi_abwesend": ["AA"],   # AA: nie Mo/Mi UND max. 1 Einsatz/Woche
    "feiertage": []
}
root, app = fresh_app(POOLS_F)
app.start_date_entry.delete(0, tk.END)
app.start_date_entry.insert(0, "02.02.2026")
app.fair_mode_var.set(True)
app.create_planning()
check("Plan erzeugt", len(app.planning_result) == 10)
fehler = [v for v in app._validate_current_plan() if v["schwere"] == "fehler"]
check("Keine Regel-Fehler", not fehler, str(fehler))
for w in (0, 1):
    loads = week_loads(app.planning_result, ["AA"], w)
    check(f"Woche {w + 1}: AA hat max. 1 Einsatz", loads["AA"] <= 1, f"AA={loads['AA']}")

# Handgebauter Verstoß: AA zweimal in Woche 1 -> Prüfung meldet Pool-F-Limit
app.planning_result[1]["Nachmittag"] = "AA"  # Dienstag
app.planning_result[3]["Nachmittag"] = "AA"  # Donnerstag
viol = app._validate_current_plan()
check("Prüfung meldet Pool-F-Limit-Verstoß",
      any("Pool F" in v["text"] and "max. 1" in v["text"] for v in viol),
      " | ".join(v["text"] for v in viol))
root.destroy()

# ---------------------------------------------------------------- Test 11
print("\n--- Test 11: Rückkehr-Regel (gestern ganztags abwesend => heute kein VM/Support) ---")
root, app = fresh_app(BASE_POOLS)
app.start_date_entry.delete(0, tk.END)
app.start_date_entry.insert(0, "02.02.2026")
app.fair_mode_var.set(True)
app.absences = [
    {"ma": "AA", "von": datetime(2026, 2, 3), "bis": datetime(2026, 2, 3), "zeit": "ganz"},   # Di
    {"ma": "BB", "von": datetime(2026, 2, 6), "bis": datetime(2026, 2, 6), "zeit": "ganz"},   # Fr
]
app.create_planning()
check("Plan erzeugt", len(app.planning_result) == 10)
fehler = [v for v in app._validate_current_plan() if v["schwere"] == "fehler"]
check("Keine Regel-Fehler", not fehler, str(fehler))
row_mi = app.planning_result[2]   # Mittwoch nach AAs Dienstag-Abwesenheit
check("Mi: AA nicht VM/Support (gestern ganztags abwesend)",
      row_mi["Vormittag"] != "AA" and row_mi["Support"] != "AA", str(row_mi))
row_mo2 = app.planning_result[5]  # Montag W2 nach BBs Freitag-Abwesenheit
check("Mo W2: BB nicht VM/Support (Fr ganztags abwesend, gilt über WE)",
      row_mo2["Vormittag"] != "BB" and row_mo2["Support"] != "BB", str(row_mo2))

# Handgebauter Verstoß -> konkrete Meldung
app.planning_result[2]["Vormittag"] = "AA"
viol = app._validate_current_plan()
check("Prüfung meldet Rückkehr-Verstoß konkret",
      any("Vortag ganztags abwesend" in v["text"] and v["row"] == 2 for v in viol),
      " | ".join(v["text"] for v in viol))
root.destroy()

# ---------------------------------------------------------------- Test 12
print("\n--- Test 12: Screenshot-Szenario — keine Schichttyp-Häufungen durch Optimierer ---")
root, app = fresh_app(REAL_POOLS)
app.start_date_entry.delete(0, tk.END)
app.start_date_entry.insert(0, "06.07.2026")  # Montag
app.fair_mode_var.set(True)
app.fairness_horizon_var.set("Nur dieser Plan")
app.first_vm_entry.insert(0, "TR")
app.first_nm_entry.insert(0, "FA")
app.first_support_entry.insert(0, "IL")
app.absences = [
    {"ma": "RR", "von": datetime(2026, 7, 6), "bis": datetime(2026, 7, 10), "zeit": "ganz"},
    {"ma": "JB", "von": datetime(2026, 7, 14), "bis": datetime(2026, 7, 15), "zeit": "ganz"},
    {"ma": "AN", "von": datetime(2026, 7, 15), "bis": datetime(2026, 7, 16), "zeit": "ganz"},
    {"ma": "MH", "von": datetime(2026, 7, 15), "bis": datetime(2026, 7, 16), "zeit": "ganz"},
    {"ma": "TR", "von": datetime(2026, 7, 17), "bis": datetime(2026, 7, 17), "zeit": "ganz"},
]
app.create_planning()
for r in app.planning_result:
    print(f"    {r['Datum']} {r['Wochentag'][:2]}: VM={r['Vormittag']:3s} NM={r['Nachmittag']:3s} Sup={r['Support']}")
check("Plan erzeugt", len(app.planning_result) == 10)
fehler = [v for v in app._validate_current_plan() if v["schwere"] == "fehler"]
check("Keine Regel-Fehler", not fehler, str(fehler))

# Kein MA hat denselben Schichttyp an zwei direkt aufeinanderfolgenden Tagen
consecutive_same_type = []
for i in range(1, len(app.planning_result)):
    prev_e, cur_e = app.planning_result[i - 1], app.planning_result[i]
    for st in ("Vormittag", "Nachmittag", "Support"):
        p = prev_e[st].strip()
        c = cur_e[st].strip()
        if st == "Support":
            if p == prev_e["Vormittag"].strip():
                p = ""
            if c == cur_e["Vormittag"].strip():
                c = ""
        if p and p == c:
            consecutive_same_type.append((cur_e["Datum"], st, c))
check("Kein gleicher Schichttyp an Folgetagen", not consecutive_same_type,
      str(consecutive_same_type))
il_total = sum(
    1 for e in app.planning_result
    for st in ("Vormittag", "Nachmittag", "Support")
    if e[st].strip() == "IL" and not (st == "Support" and e["Vormittag"].strip() == "IL")
)
check("IL ist angemessen eingeplant (>= 2 Einsätze)", il_total >= 2, f"IL={il_total}")
root.destroy()

# ---------------------------------------------------------------- Test 13
print("\n--- Test 13: Gewichtung — Pool C trägt ca. einen Einsatz mehr als Pool B ---")
root, app = fresh_app(REAL_POOLS)
app.start_date_entry.delete(0, tk.END)
app.start_date_entry.insert(0, "06.07.2026")
app.fair_mode_var.set(True)
app.fairness_horizon_var.set("Nur dieser Plan")
app.create_planning()
fehler = [v for v in app._validate_current_plan() if v["schwere"] == "fehler"]
check("Keine Regel-Fehler", not fehler, str(fehler))
totals = {}
for e in app.planning_result:
    vm = e["Vormittag"].strip()
    for st in ("Vormittag", "Nachmittag", "Support"):
        ma = e[st].strip()
        if ma and not (st == "Support" and ma == vm):
            totals[ma] = totals.get(ma, 0) + 1
print(f"    Gesamtlast: { {ma: totals.get(ma, 0) for ma in sorted(set(sum([REAL_POOLS[k] for k in ('pool_vm_alle','pool_vm_support','pool_nm_alle')], [])))} }")
pool_b = ["MH", "RI"]
core_c = ["IL", "RR", "FA", "JB", "TR"]
b_vals = [totals.get(ma, 0) for ma in pool_b]
c_vals = [totals.get(ma, 0) for ma in core_c]
check("Kein Pool-C-MA unter Pool-B-Niveau", min(c_vals) >= max(b_vals),
      f"B={dict(zip(pool_b, b_vals))} C={dict(zip(core_c, c_vals))}")
avg_b = sum(b_vals) / len(b_vals)
avg_c = sum(c_vals) / len(c_vals)
check("Pool C im Schnitt ca. 1 Einsatz über Pool B", round(avg_c - avg_b, 2) >= 0.8,
      f"avg_B={avg_b:.2f} avg_C={avg_c:.2f}")
root.destroy()

# ---------------------------------------------------------------- Ergebnis
print("\n=================================================")
if FAILURES:
    print(f"{len(FAILURES)} TEST(S) FEHLGESCHLAGEN:")
    for name in FAILURES:
        print(f"  - {name}")
    sys.exit(1)
print("ALLE TESTS BESTANDEN")
