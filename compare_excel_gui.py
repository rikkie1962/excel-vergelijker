#!/usr/bin/env python3
r"""
Excel/CSV Vergelijker – CAD (alle stands) MIN Bestellingen

Workflow:
1) Kies het BESTELLINGEN-bestand (mag dubbelen hebben) en de kolom met standnummers.
2) Kies het CAD-bestand met ALLE STANDNUMMERS (uniek) en de kolom met standnummers.
Resultaat: standnummers die wel in CAD staan maar niet in Bestellingen (CAD \\ Bestellingen),
oplopend gesorteerd (natuurlijke sortering). Output = .xlsx
"""

from pathlib import Path
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import Optional
import pandas as pd

# =========================
# Helper functies
# =========================

STAND_LIKE_RE = re.compile(r"^\s*\d+(?:\.[A-Za-z0-9]+)*[A-Za-z0-9]*\s*$")
# matcht o.a.: 0.A01, 8.F49, 9.A02, 0.A17a, 0.BAR2

def is_stand_like(s: str) -> bool:
    if s is None:
        return False
    return bool(STAND_LIKE_RE.match(str(s)))

def natural_key(s: str):
    """
    Natuurlijke sortering met type-tags zodat er nooit str vs int vergeleken wordt.
    Voorbeelden: 0.A9 < 0.A10 < 1.A01 < 8.F49
    """
    parts = re.findall(r"\d+|\D+", "" if s is None else str(s))
    key = []
    for p in parts:
        if p.isdigit():
            key.append((0, int(p)))
        else:
            key.append((1, p.upper()))
    return key

def normalize_series(ser: pd.Series) -> pd.Series:
    """Naar string, strip spaties."""
    ser = ser.astype(str).fillna("").str.strip()
    return ser

def _read_csv(path: Path):
    """CSV lezen met autodetect delimiter en encoding fallbacks."""
    try:
        df = pd.read_csv(path, dtype=str, sep=None, engine="python", encoding="utf-8-sig")
        used = ("auto", "utf-8-sig")
    except Exception:
        try:
            df = pd.read_csv(path, dtype=str, sep=";", encoding="utf-8-sig")
            used = (";", "utf-8-sig")
        except Exception:
            df = pd.read_csv(path, dtype=str, sep=";", encoding="cp1252")
            used = (";", "cp1252")
    return df, used

def read_table(path: Path) -> pd.DataFrame:
    """
    Lees CSV/XLS(X) naar DataFrame (strings). Detecteert automatisch:
    - CSV/Excel zonder header waarbij de 1e rij standnummers bevat:
      dan herlaadt met header=None zodat kolomnamen niet fout zijn.
    """
    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xls"):
        df = pd.read_excel(path, dtype=str, header=0)
        if any(is_stand_like(c) for c in df.columns):
            df = pd.read_excel(path, dtype=str, header=None)
            df.columns = [f"Kolom_{i}" for i in range(df.shape[1])]
    elif suffix == ".csv":
        df, used = _read_csv(path)
        if any(is_stand_like(c) for c in df.columns):
            sep, enc = used
            if sep == "auto":
                try:
                    df = pd.read_csv(path, dtype=str, sep=None, engine="python",
                                     encoding=enc, header=None)
                except Exception:
                    df = pd.read_csv(path, dtype=str, sep=";", encoding=enc, header=None)
            else:
                df = pd.read_csv(path, dtype=str, sep=sep, encoding=enc, header=None)
            df.columns = [f"Kolom_{i}" for i in range(df.shape[1])]
    else:
        raise ValueError(f"Niet-ondersteund bestandstype: {suffix} (gebruik .csv of .xlsx/.xls)")

    if list(df.columns) == list(range(len(df.columns))):
        df.columns = [f"Kolom_{i}" for i in range(len(df.columns))]
    return df

def ask_for_file(title: str) -> Optional[Path]:
    p = filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("Alle bestanden", "*.*")]
    )
    return Path(p) if p else None

def ask_for_column(root: tk.Tk, df: pd.DataFrame, title: str) -> Optional[str]:
    """Eenvoudige listbox om 1 kolom te kiezen."""
    dialog = tk.Toplevel(root)
    dialog.title(title)
    dialog.grab_set()
    dialog.geometry("440x380")

    tk.Label(dialog, text=title, font=("Segoe UI", 11, "bold")).pack(pady=(10, 6))
    tk.Label(dialog, text="Kies één kolom:").pack()

    frame = tk.Frame(dialog)
    frame.pack(fill="both", expand=True, padx=10, pady=8)

    lb = tk.Listbox(frame, selectmode="single")
    scrollbar = tk.Scrollbar(frame, orient="vertical", command=lb.yview)
    lb.config(yscrollcommand=scrollbar.set)
    lb.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    for col in df.columns:
        lb.insert("end", str(col))

    state = {"value": None}

    def ok_():
        sel = lb.curselection()
        if not sel:
            messagebox.showwarning("Let op", "Selecteer een kolom.")
            return
        state["value"] = lb.get(sel[0])
        dialog.destroy()

    def cancel_():
        state["value"] = None
        dialog.destroy()

    btns = tk.Frame(dialog)
    btns.pack(pady=8)
    tk.Button(btns, text="OK", width=10, command=ok_).pack(side="left", padx=5)
    tk.Button(btns, text="Annuleren", width=10, command=cancel_).pack(side="left", padx=5)

    dialog.wait_window()
    return state["value"]

def cad_minus_orders(orders_col: pd.Series, cad_col: pd.Series):
    """
    Retourneert lijst met standnummers die wél in CAD staan en NIET in Bestellingen.
    (CAD-set MIN Orders-set), natuurlijk gesorteerd.
    """
    s_orders = normalize_series(orders_col)
    s_cad = normalize_series(cad_col)

    # lege waarden negeren
    s_orders = s_orders[s_orders.str.len() > 0]
    s_cad = s_cad[s_cad.str.len() > 0]

    set_orders = set(s_orders.tolist())  # dubbelen vallen vanzelf weg
    set_cad = set(s_cad.tolist())

    result = sorted(list(set_cad - set_orders), key=natural_key)
    return result

# =========================
# Main (GUI flow)
# =========================

def main():
    root = tk.Tk()
    root.withdraw()

    messagebox.showinfo(
        "Excel/CSV Vergelijker",
        "Workflow:\n"
        "1) Kies het BESTELLINGEN-bestand (mag dubbelen hebben) en de kolom met standnummers.\n"
        "2) Kies het CAD-bestand met ALLE STANDNUMMERS (uniek) en de kolom met standnummers.\n\n"
        "Resultaat: standen die nog GEEN bestelling hebben (CAD minus Bestellingen)."
    )

    # 1) Bestellingen
    f_orders = ask_for_file("Kies het BESTELLINGEN-bestand (CSV of Excel)")
    if not f_orders:
        return
    try:
        df_orders = read_table(f_orders)
    except Exception as e:
        messagebox.showerror("Fout bij lezen BESTELLINGEN", f"{e}")
        return

    col_orders = ask_for_column(root, df_orders, f"Kolom met standnummers (BESTELLINGEN) uit: {f_orders.name}")
    if not col_orders:
        return

    # 2) CAD / Alle standnummers
    f_cad = ask_for_file("Kies het CAD-bestand met ALLE STANDNUMMERS (CSV of Excel)")
    if not f_cad:
        return
    try:
        df_cad = read_table(f_cad)
    except Exception as e:
        messagebox.showerror("Fout bij lezen CAD/ALLE", f"{e}")
        return

    col_cad = ask_for_column(root, df_cad, f"Kolom met standnummers (CAD/ALLE) uit: {f_cad.name}")
    if not col_cad:
        return

    # 3) Vergelijken
    try:
        result = cad_minus_orders(df_orders[col_orders], df_cad[col_cad])
    except Exception as e:
        messagebox.showerror("Fout bij vergelijken", f"{e}")
        return

    # 4) Output-pad
    out_path = filedialog.asksaveasfilename(
        title="Kies naam en map voor het output bestand",
        defaultextension=".xlsx",
        filetypes=[("Excel bestand", "*.xlsx")]
    )
    if not out_path:
        return

    # 5) Wegschrijven
    try:
        out_df = pd.DataFrame({"Standnummers_zonder_bestelling": result})
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            out_df.to_excel(writer, index=False, sheet_name="Resultaat")
        messagebox.showinfo(
            "Klaar",
            f"Output weggeschreven naar:\n{out_path}\n\nAantal rijen: {len(out_df)}"
        )
    except Exception as e:
        messagebox.showerror("Fout bij wegschrijven", f"{e}")
        return

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        try:
            messagebox.showerror("Onverwachte fout", f"{e}")
        except Exception:
            pass
        raise



