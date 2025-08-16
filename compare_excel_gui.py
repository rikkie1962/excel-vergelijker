#!/usr/bin/env python3
r"""
Excel/CSV Vergelijker – Beurs use-case

Doel:
- 1e bestand = BESTELLINGEN (kan dubbelen bevatten) → kies kolom met standnummers
- 2e bestand = CAD/ALLE STANDNUMMERS (uniek) → kies kolom met standnummers
- Output = alle standnummers die wél in CAD staan maar géén bestelling hebben
          (dus: CAD \\ BESTELLINGEN), oplopend gesorteerd (natuurlijke sortering)
- Output wordt als .xlsx weggeschreven.

Ondersteunt .csv, .xlsx en .xls
GUI met tkinter.
"""

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import Optional

import pandas as pd


# =========================
# Helpers
# =========================

def read_table(path: Path) -> pd.DataFrame:
    """
    Lees CSV of Excel naar DataFrame. Alles als string om type-issues te voorkomen.
    Probeert delimiter/encoding voor CSV.
    """
    suffix = path.suffix.lower()
    if suffix in [".xlsx", ".xls"]:
        df = pd.read_excel(path, dtype=str)
    elif suffix == ".csv":
        try:
            # autodetect delimiter
            df = pd.read_csv(path, dtype=str, sep=None, engine="python", encoding="utf-8-sig")
        except Exception:
            # common fallbacks in NL/EU
            try:
                df = pd.read_csv(path, dtype=str, sep=";", encoding="utf-8-sig")
            except Exception:
                df = pd.read_csv(path, dtype=str, sep=";", encoding="cp1252")
    else:
        raise ValueError(f"Niet-ondersteund bestandstype: {suffix} (gebruik .csv of .xlsx/.xls)")

    # Als er geen header was en Pandas 0..N kolomnamen gaf, maak dan Kolom_0, Kolom_1, ...
    try:
        if list(df.columns) == list(range(len(df.columns))):
            df.columns = [f"Kolom_{i}" for i in range(len(df.columns))]
    except Exception:
        pass

    return df


def ask_for_file(title: str) -> Optional[Path]:
    path_str = filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("Alle bestanden", "*.*")]
    )
    if not path_str:
        return None
    return Path(path_str)


def ask_for_column(root: tk.Tk, df: pd.DataFrame, title: str) -> Optional[str]:
    """
    Eenvoudige listbox om 1 kolom te kiezen.
    """
    dialog = tk.Toplevel(root)
    dialog.title(title)
    dialog.grab_set()
    dialog.geometry("420x360")

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

    selection = {"value": None}

    def on_ok():
        sel = lb.curselection()
        if not sel:
            messagebox.showwarning("Let op", "Selecteer een kolom.")
            return
        selection["value"] = lb.get(sel[0])
        dialog.destroy()

    def on_cancel():
        selection["value"] = None
        dialog.destroy()

    btns = tk.Frame(dialog)
    btns.pack(pady=8)
    tk.Button(btns, text="OK", width=10, command=on_ok).pack(side="left", padx=5)
    tk.Button(btns, text="Annuleren", width=10, command=on_cancel).pack(side="left", padx=5)

    dialog.wait_window()
    return selection["value"]


def natural_key(s: str):
    """
    Sorteersleutel die gemengde cijfers/letters netjes sorteert.
    Voorbeelden: 3A21 < 11F33 < 100A1
    """
    import re
    if s is None:
        return []
    s = str(s)
    parts = re.findall(r"\d+|\D+", s)
    key = []
    for p in parts:
        if p.isdigit():
            key.append(int(p))
        else:
            key.append(p.upper())
    return key


def normalize_series(ser: pd.Series) -> pd.Series:
    """
    Alles naar string, strip spaties, None->"".
    """
    ser = ser.astype(str)
    ser = ser.fillna("")
    ser = ser.str.strip()
    return ser


def compute_stands_without_orders(col_orders: pd.Series, col_all: pd.Series):
    """
    Retourneert alle standnummers die:
      - WÉL in CAD/ALLE_STANDEN staan
      - NIET in BESTELLINGEN voorkomen (ook niet 1x)

    Formeel: resultaat = set(col_all) - set(col_orders)

    col_orders: kolom uit BESTELLINGEN (mag dubbelen hebben)
    col_all:    kolom uit CAD (uniek)
    """
    s_orders = normalize_series(col_orders)
    s_all = normalize_series(col_all)

    # lege waarden negeren
    s_orders = s_orders[s_orders.str.len() > 0]
    s_all = s_all[s_all.str.len() > 0]

    set_orders = set(s_orders.tolist())
    set_all = set(s_all.tolist())

    missing = list(set_all - set_orders)  # CAD \ BESTELLINGEN
    return sorted(missing, key=natural_key)


# =========================
# Main (GUI flow)
# =========================

def main():
    root = tk.Tk()
    root.withdraw()  # geen lege hoofdwindow

    messagebox.showinfo(
        "Excel/CSV Vergelijker",
        "Workflow:\n"
        "1) Kies het BESTELLINGEN-bestand (kan dubbelen bevatten) en de kolom met standnummers.\n"
        "2) Kies het CAD-bestand met ALLE STANDNUMMERS (uniek) en de kolom.\n\n"
        "De tool geeft als resultaat: standen die nog GEEN bestelling hebben (CAD minus Bestellingen)."
    )

    # 1e bestand = Bestellingen
    f1 = ask_for_file("Kies het BESTELLINGEN-bestand (CSV of Excel)")
    if f1 is None:
        return
    try:
        df1 = read_table(f1)
    except Exception as e:
        messagebox.showerror("Fout bij lezen BESTELLINGEN", f"{e}")
        return

    col1 = ask_for_column(root, df1, f"Kolom met standnummers (BESTELLINGEN) uit: {f1.name}")
    if col1 is None:
        return

    # 2e bestand = CAD / Alle standnummers
    f2 = ask_for_file("Kies het CAD-bestand met ALLE STANDNUMMERS (CSV of Excel)")
    if f2 is None:
        return
    try:
        df2 = read_table(f2)
    except Exception as e:
        messagebox.showerror("Fout bij lezen CAD/ALLE", f"{e}")
        return

    col2 = ask_for_column(root, df2, f"Kolom met standnummers (CAD/ALLE) uit: {f2.name}")
    if col2 is None:
        return

    # Vergelijken
    try:
        result = compute_stands_without_orders(df1[col1], df2[col2])
    except Exception as e:
        messagebox.showerror("Fout bij vergelijken", f"{e}")
        return

    # Output-pad kiezen
    out_path = filedialog.asksaveasfilename(
        title="Kies naam en map voor het output bestand",
        defaultextension=".xlsx",
        filetypes=[("Excel bestand", "*.xlsx")]
    )
    if not out_path:
        return

    # Wegschrijven
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
        # Laat een melding zien als er iets onverwachts fout gaat (bijv. buiten GUI context)
        try:
            messagebox.showerror("Onverwachte fout", f"{e}")
        except Exception:
            pass
        raise




