#!/usr/bin/env python3
# (Same code as before, included here for the repo)
import sys
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

def read_table(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix in ['.xlsx', '.xls']:
        df = pd.read_excel(path, dtype=str)
    elif suffix == '.csv':
        try:
            df = pd.read_csv(path, dtype=str, sep=None, engine='python', encoding='utf-8-sig')
        except Exception:
            try:
                df = pd.read_csv(path, dtype=str, sep=';', encoding='utf-8-sig')
            except Exception:
                df = pd.read_csv(path, dtype=str, sep=';', encoding='cp1252')
    else:
        raise ValueError(f"Unsupported file type: {suffix}. Use CSV or Excel.")
    if df.columns.to_list() == list(range(len(df.columns))):
        df.columns = [f"Kolom_{i}" for i in range(len(df.columns))]
    return df

def natural_key(s: str):
    import re
    parts = re.findall(r'\d+|\D+', s)
    key = []
    for p in parts:
        if p.isdigit():
            key.append(int(p))
        else:
            key.append(p.upper())
    return key

def normalize_series(ser: pd.Series) -> pd.Series:
    ser = ser.astype(str).fillna("")
    ser = ser.str.strip()
    return ser

def compute_stands_without_orders(col_orders: pd.Series, col_all: pd.Series):
    """
    Geeft alle standnummers terug die wel in 'alle_standen' (CAD) staan,
    maar niet in 'bestellingen'.

    - col_orders: kolom uit het bestellingenbestand (kan dubbelen bevatten)
    - col_all:    kolom uit het CAD-bestand (unieke nummers)
    """
    s_orders = normalize_series(col_orders)
    s_all = normalize_series(col_all)

    # lege waarden negeren
    s_orders = s_orders[s_orders.str.len() > 0]
    s_all = s_all[s_all.str.len() > 0]

    set_orders = set(s_orders.tolist())
    set_all = set(s_all.tolist())

    missing = list(set_all - set_orders)  # CAD \ Bestellingen
    return sorted(missing, key=natural_key)


def ask_for_file(title: str):
    from tkinter import filedialog
    path_str = filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("Alle bestanden", "*.*")]
    )
    return Path(path_str) if path_str else None

def ask_for_column(root: tk.Tk, df: pd.DataFrame, title: str):
    dialog = tk.Toplevel(root); dialog.title(title); dialog.grab_set(); dialog.geometry("420x360")
    tk.Label(dialog, text=title, font=("Segoe UI", 11, "bold")).pack(pady=(10, 6))
    tk.Label(dialog, text="Kies één kolom:").pack()
    frame = tk.Frame(dialog); frame.pack(fill="both", expand=True, padx=10, pady=8)
    lb = tk.Listbox(frame, selectmode="single")
    sc = tk.Scrollbar(frame, orient="vertical", command=lb.yview); lb.config(yscrollcommand=sc.set)
    lb.pack(side="left", fill="both", expand=True); sc.pack(side="right", fill="y")
    for col in df.columns: lb.insert("end", str(col))
    sel = {"v": None}
    def on_ok():
        if not lb.curselection():
            messagebox.showwarning("Let op", "Selecteer een kolom."); return
        sel["v"] = lb.get(lb.curselection()[0]); dialog.destroy()
    def on_cancel(): sel["v"] = None; dialog.destroy()
    btns = tk.Frame(dialog); btns.pack(pady=8)
    tk.Button(btns, text="OK", width=10, command=on_ok).pack(side="left", padx=5)
    tk.Button(btns, text="Annuleren", width=10, command=on_cancel).pack(side="left", padx=5)
    dialog.wait_window(); return sel["v"]

def main():
    # Intro
messagebox.showinfo(
    "Excel/CSV Vergelijker",
    "Kies eerst het BESTELLINGEN-bestand (kan dubbelen bevatten) en de kolom met standnummers.\n"
    "Kies daarna het CAD-bestand met ALLE STANDNUMMERS (uniek) en de kolom.\n"
    "Output: standen die nog GEEN bestelling hebben."
)

# 1e bestand = Bestellingen (kan dubbelen bevatten)
f1 = ask_for_file("Kies het BESTELLINGEN-bestand (CSV of Excel)")
...
col1 = ask_for_column(root, df1, f"Kolom met standnummers (BESTELLINGEN) uit: {f1.name}")

# 2e bestand = CAD (alle standen, uniek)
f2 = ask_for_file("Kies het CAD-bestand met ALLE STANDNUMMERS (CSV of Excel)")
...
col2 = ask_for_column(root, df2, f"Kolom met standnummers (CAD/ALLE) uit: {f2.name}")

# Vergelijken
try:
    # let op: eerst bestellingen, dan alle_standen
    result = compute_stands_without_orders(df1[col1], df2[col2])
except Exception as e:
    ...

