"""Load and clean stok and jual data from Excel files."""
from __future__ import annotations
from pathlib import Path
from typing import Tuple

import pandas as pd

from config import (
    STOK_SHEET, JUAL_SHOPEE_SHEET, JUAL_TIKTOK_SHEET, EXCLUDED_SKUS,
    COL_STOK_SKU, COL_STOK_QTY, COL_STOK_TOTAL_HPP, COL_STOK_TANGGAL_BAYAR,
    COL_JUAL_QTY, COL_JUAL_OMZET, COL_JUAL_KODE_UNIK,
    COL_JUAL_TAMBAHAN, COL_JUAL_TANGGAL, COL_JUAL_AKUN,
)


def validate_file(file_path: Path, required_sheets: list[str]) -> list[str]:
    """Validate file exists. Return list of sheet names actually present.

    Raises FileNotFoundError or ValueError if required sheets are missing.
    """
    if not file_path.exists():
        raise FileNotFoundError(f"File tidak ditemukan: {file_path}")
    xl = pd.ExcelFile(file_path)
    missing = set(required_sheets) - set(xl.sheet_names)
    if missing:
        raise ValueError(
            f"Sheet hilang di {file_path.name}: {missing}. "
            f"Sheet tersedia: {xl.sheet_names}"
        )
    return xl.sheet_names


def load_stok(file_path: Path) -> pd.DataFrame:
    """Load stok file. Headers on row 2 (header=1 in pandas)."""
    print(f"✓ Membaca file stok: {file_path.name}")
    validate_file(file_path, [STOK_SHEET])
    df = pd.read_excel(file_path, sheet_name=STOK_SHEET, header=1)

    required_cols = [COL_STOK_SKU, COL_STOK_QTY, COL_STOK_TOTAL_HPP]
    missing = set(required_cols) - set(df.columns)
    if missing:
        raise ValueError(f"Kolom hilang di sheet {STOK_SHEET}: {missing}")

    df = df.rename(columns={
        COL_STOK_QTY: "qty_beli",
        COL_STOK_TOTAL_HPP: "total_hpp",
        COL_STOK_TANGGAL_BAYAR: "tanggal_bayar",
    })

    before = len(df)
    df = df[df["SKU"].notna()].copy()
    df["qty_beli"] = pd.to_numeric(df["qty_beli"], errors="coerce")
    df["total_hpp"] = pd.to_numeric(df["total_hpp"], errors="coerce")
    if "tanggal_bayar" in df.columns:
        df["tanggal_bayar"] = pd.to_datetime(df["tanggal_bayar"], errors="coerce")
    df = df.dropna(subset=["qty_beli", "total_hpp"])

    print(f"  Stok valid rows: {len(df)} (dari {before} mentah)")
    return df


def load_jual(file_path: Path) -> pd.DataFrame:
    """Load jual file (Shopee required, Tiktok optional) and combine."""
    print(f"✓ Membaca file jual: {file_path.name}")
    available = validate_file(file_path, [JUAL_SHOPEE_SHEET])

    parts: list[pd.DataFrame] = []
    shopee = pd.read_excel(file_path, sheet_name=JUAL_SHOPEE_SHEET)
    shopee["_sheet_source"] = "Shopee"
    parts.append(shopee)
    print(f"  Loaded {JUAL_SHOPEE_SHEET}: {len(shopee)} rows")

    if JUAL_TIKTOK_SHEET in available:
        tiktok = pd.read_excel(file_path, sheet_name=JUAL_TIKTOK_SHEET)
        tiktok["_sheet_source"] = "Tiktok"
        parts.append(tiktok)
        print(f"  Loaded {JUAL_TIKTOK_SHEET}: {len(tiktok)} rows")
    else:
        print(f"  ⚠ Sheet {JUAL_TIKTOK_SHEET} tidak ada — lanjut tanpa data Tiktok/Tokopedia")

    df = pd.concat(parts, ignore_index=True)
    df = df.rename(columns={
        COL_JUAL_QTY: "qty_jual",
        COL_JUAL_OMZET: "omzet",
        COL_JUAL_KODE_UNIK: "kode_unik",
        COL_JUAL_TAMBAHAN: "tambahan",
        COL_JUAL_TANGGAL: "tanggal_pesan",
        COL_JUAL_AKUN: "akun_penjual",
    })
    return df


def clean_jual(df: pd.DataFrame) -> Tuple[pd.DataFrame, dict]:
    """Clean jual: coerce types, apply exclusion rules.

    Returns (cleaned_df, stats_dict) where stats has exclusion counts.
    """
    stats = {"total_raw": len(df)}

    # Coerce numeric columns
    for col in ["omzet", "qty_jual", "tambahan", "kode_unik"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    df["tanggal_pesan"] = pd.to_datetime(df["tanggal_pesan"], errors="coerce")

    # Fill admin fields
    df["tambahan"] = df.get("tambahan", 0).fillna(0)
    df["kode_unik"] = df.get("kode_unik", 0).fillna(0)

    # Exclusion 1: SKU null
    before = len(df)
    df = df[df["SKU"].notna()].copy()
    stats["dropped_sku_null"] = before - len(df)

    # Exclusion 2: Dummy invoices
    before = len(df)
    df = df[~df["Invoice"].astype(str).str.startswith("Dummy", na=False)].copy()
    stats["dropped_dummy"] = before - len(df)

    # Exclusion 3: Void
    before = len(df)
    if "Void" in df.columns:
        df = df[df["Void"] != True].copy()
    stats["dropped_void"] = before - len(df)

    # Exclusion 4: SKUs in exclusion list
    before = len(df)
    df = df[~df["SKU"].isin(EXCLUDED_SKUS)].copy()
    stats["dropped_excluded_sku"] = before - len(df)

    # Exclusion 5: rows with invalid numeric (e.g., scattered header rows)
    before = len(df)
    df = df.dropna(subset=["omzet", "qty_jual"]).copy()
    stats["dropped_invalid_numeric"] = before - len(df)

    stats["total_clean"] = len(df)
    print(f"✓ Cleaning jual:")
    print(f"  - Total mentah         : {stats['total_raw']:>6}")
    print(f"  - Dibuang (SKU null)   : {stats['dropped_sku_null']:>6}")
    print(f"  - Dibuang (Dummy)      : {stats['dropped_dummy']:>6}")
    print(f"  - Dibuang (Void)       : {stats['dropped_void']:>6}")
    print(f"  - Dibuang (SKU exclude): {stats['dropped_excluded_sku']:>6}")
    print(f"  - Dibuang (invalid)    : {stats['dropped_invalid_numeric']:>6}")
    print(f"  - Total bersih         : {stats['total_clean']:>6}")
    return df, stats
