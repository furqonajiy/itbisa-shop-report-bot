"""Load and clean stok and jual data."""
from __future__ import annotations
from pathlib import Path
from typing import Tuple

import numpy as np
import pandas as pd

from config import (
    COL_JUAL_AKUN, COL_JUAL_KODE_UNIK, COL_JUAL_OMZET, COL_JUAL_QTY,
    COL_JUAL_TAMBAHAN, COL_JUAL_TANGGAL, COL_STOK_LUAR_NEGERI, COL_STOK_QTY,
    COL_STOK_TANGGAL_BAYAR, COL_STOK_TOKO, COL_STOK_TOTAL_HPP,
    EXCLUDED_SKUS, JUAL_SHEETS, REQUIRED_JUAL_SHEET, STOK_SHEET,
)


def load_stok_files(file_paths: list[Path]) -> pd.DataFrame:
    """Load and combine stok files. Dedupe on (SKU, date, qty, total_hpp)."""
    if not file_paths:
        raise FileNotFoundError("Tidak ada file Stok ditemukan di data/")

    parts = []
    for fp in file_paths:
        print(f"✓ Membaca stok: {fp.name}")
        df = pd.read_excel(fp, sheet_name=STOK_SHEET, header=1)
        required = [COL_STOK_QTY, COL_STOK_TOTAL_HPP]
        missing = set(required) - set(df.columns)
        if missing:
            raise ValueError(f"Kolom hilang di {fp.name}: {missing}")
        parts.append(df)

    df = pd.concat(parts, ignore_index=True)
    df = df.rename(columns={
        COL_STOK_QTY: "qty_beli",
        COL_STOK_TOTAL_HPP: "total_hpp",
        COL_STOK_TANGGAL_BAYAR: "tanggal_bayar",
        COL_STOK_TOKO: "toko",
        COL_STOK_LUAR_NEGERI: "luar_negeri",
    })
    for col in ["toko", "luar_negeri", "tanggal_bayar"]:
        if col not in df.columns:
            df[col] = np.nan

    raw_count = len(df)
    df = df[df["SKU"].notna()].copy()
    df["qty_beli"] = pd.to_numeric(df["qty_beli"], errors="coerce")
    df["total_hpp"] = pd.to_numeric(df["total_hpp"], errors="coerce")
    df["tanggal_bayar"] = pd.to_datetime(df["tanggal_bayar"], errors="coerce")
    df["luar_negeri"] = pd.to_numeric(df["luar_negeri"], errors="coerce")
    df = df.dropna(subset=["qty_beli", "total_hpp"])

    before_dedup = len(df)
    df = df.drop_duplicates(
        subset=["SKU", "tanggal_bayar", "qty_beli", "total_hpp"], keep="first"
    )
    n_dup = before_dedup - len(df)
    print(f"  → Stok bersih: {len(df):,} baris (dari {raw_count:,} mentah, "
          f"hapus {n_dup:,} duplikat)")
    return df


def load_jual_files(file_paths: list[Path]) -> pd.DataFrame:
    """Load all available jual sheets from all files."""
    if not file_paths:
        raise FileNotFoundError("Tidak ada file Jual ditemukan di data/")

    parts = []
    for fp in file_paths:
        xl = pd.ExcelFile(fp)
        if REQUIRED_JUAL_SHEET not in xl.sheet_names:
            print(f"  ⚠ {fp.name} tidak punya sheet wajib '{REQUIRED_JUAL_SHEET}', dilewati")
            continue
        print(f"✓ Membaca jual: {fp.name}")
        for sheet in JUAL_SHEETS:
            if sheet in xl.sheet_names:
                df = pd.read_excel(fp, sheet_name=sheet)
                df["_sheet_source"] = sheet
                parts.append(df)
                print(f"  → {sheet}: {len(df):,} baris")

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


def clean_jual(df: pd.DataFrame, year: int | None = None) -> Tuple[pd.DataFrame, dict]:
    """Clean jual; if year provided, filter to that year by tanggal_pesan."""
    stats = {"total_raw": len(df)}

    for col in ["omzet", "qty_jual", "tambahan", "kode_unik"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    df["tanggal_pesan"] = pd.to_datetime(df["tanggal_pesan"], errors="coerce")
    df["tambahan"] = df["tambahan"].fillna(0) if "tambahan" in df.columns else 0
    df["kode_unik"] = df["kode_unik"].fillna(0) if "kode_unik" in df.columns else 0

    before = len(df)
    df = df[df["SKU"].notna()].copy()
    stats["dropped_sku_null"] = before - len(df)

    before = len(df)
    df = df[~df["Invoice"].astype(str).str.startswith("Dummy", na=False)].copy()
    stats["dropped_dummy"] = before - len(df)

    before = len(df)
    if "Void" in df.columns:
        df = df[df["Void"] != True].copy()
    stats["dropped_void"] = before - len(df)

    before = len(df)
    df = df[~df["SKU"].isin(EXCLUDED_SKUS)].copy()
    stats["dropped_excluded_sku"] = before - len(df)

    before = len(df)
    df = df.dropna(subset=["omzet", "qty_jual"]).copy()
    stats["dropped_invalid_numeric"] = before - len(df)

    if year is not None:
        before = len(df)
        df = df[df["tanggal_pesan"].dt.year == year].copy()
        stats["filtered_year"] = before - len(df)

    stats["total_clean"] = len(df)
    print(f"✓ Cleaning jual:")
    print(f"    Mentah               : {stats['total_raw']:>7,}")
    print(f"    Buang (SKU null)     : {stats['dropped_sku_null']:>7,}")
    print(f"    Buang (Dummy)        : {stats['dropped_dummy']:>7,}")
    print(f"    Buang (Void)         : {stats['dropped_void']:>7,}")
    print(f"    Buang (SKU exclude)  : {stats['dropped_excluded_sku']:>7,}")
    print(f"    Buang (invalid)      : {stats['dropped_invalid_numeric']:>7,}")
    if year is not None:
        print(f"    Filter tahun {year}   : buang {stats['filtered_year']:>5,}")
    print(f"    Bersih               : {stats['total_clean']:>7,}")
    return df, stats
