#!/usr/bin/env python3
"""
Excel/CSV Vergelijker – Beurs use-case

Doel:
- 1e bestand = BESTELLINGEN (kan dubbelen bevatten) → kies kolom met standnummers
- 2e bestand = CAD/ALLE STANDNUMMERS (uniek) → kies kolom met standnummers
- Output = alle standnummers die wél in CAD staan maar géén bestelling hebben
          (dus: CAD \ BESTELLINGEN), oplopend gesorteerd (natuurlijke sortering)
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
    if not path



