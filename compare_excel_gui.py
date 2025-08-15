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

def compute_unique_values(col1: pd.Series, col2: pd.Series):
    s1 = normalize_series(col1)
    s2 = normalize_series(col2)
    combined = pd.concat([s1, s2], ignore_index=True)
    combined = combined[combined.astype(str).str.len() > 0]
    counts = combined.value_counts(dropna=False)
    uniques = counts[counts == 1].index.tolist()
    uniques_sorted = sorted(uniques, key=natural_key)
    return uniques_sorted

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
    root = tk.Tk(); root.withdraw()
    messagebox.showinfo("Excel/CSV Vergelijker",
        "Deze tool vergelijkt twee bestanden op basis van een gekozen kolom.\n"
        "- Dubbelen (ook meerdere keren in één bestand) worden verwijderd.\n"
        "- Alleen unieke waarden over beide bestanden blijven over.\n"
        "- Resultaat wordt gesorteerd en als .xlsx opgeslagen.")
    f1 = ask_for_file("Kies het 1e doelbestand (CSV of Excel)")
    if not f1: return
    try: df1 = read_table(f1)
    except Exception as e: messagebox.showerror("Fout bij lezen bestand 1", f"{e}"); return
    col1 = ask_for_column(root, df1, f"Kolom kiezen uit: {f1.name}")
    if not col1: return
    f2 = ask_for_file("Kies het 2e doelbestand (CSV of Excel)")
    if not f2: return
    try: df2 = read_table(f2)
    except Exception as e: messagebox.showerror("Fout bij lezen bestand 2", f"{e}"); return
    col2 = ask_for_column(root, df2, f"Kolom kiezen uit: {f2.name}")
    if not col2: return
    try: uniques = compute_unique_values(df1[col1], df2[col2])
    except Exception as e: messagebox.showerror("Fout bij vergelijken", f"{e}"); return
    out_path = filedialog.asksaveasfilename(
        title="Kies naam en map voor het output bestand",
        defaultextension=".xlsx",
        filetypes=[("Excel bestand", "*.xlsx")]
    )
    if not out_path: return
    try:
        out_df = pd.DataFrame({"Unieke_waarden": uniques})
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            out_df.to_excel(writer, index=False, sheet_name="Resultaat")
        messagebox.showinfo("Klaar", f"Output weggeschreven naar:\n{out_path}\n\nAantal rijen: {len(out_df)}")
    except Exception as e:
        messagebox.showerror("Fout bij wegschrijven", f"{e}")

if __name__ == "__main__":
    try: main()
    except Exception as e:
        try: messagebox.showerror("Onverwachte fout", f"{e}")
        except: pass
        raise
