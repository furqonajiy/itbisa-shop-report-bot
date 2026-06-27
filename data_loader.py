"""Load and clean stok and jual data."""
from __future__ import annotations
from pathlib import Path

import numpy as np
import pandas as pd

from config import (
    COL_JUAL_AKUN, COL_JUAL_GUDANG, COL_JUAL_KODE_UNIK, COL_JUAL_OMZET, COL_JUAL_QTY,
    COL_JUAL_TAMBAHAN, COL_JUAL_TANGGAL, COL_STOK_ALAMAT, COL_STOK_LUAR_NEGERI,
    COL_STOK_QTY, COL_STOK_TANGGAL_BAYAR, COL_STOK_TANGGAL_SAMPAI, COL_STOK_TOKO,
    COL_STOK_TOKO_LEGACY, COL_STOK_TOTAL_HPP, COL_HILANG_GUDANG, COL_HILANG_HILANG, COL_HILANG_KETEMU,
    COL_HILANG_SKU, COL_PINDAH_KURANG, COL_PINDAH_QTY, COL_PINDAH_SKU,
    COL_PINDAH_TAMBAH, EXCLUDED_SKUS, HILANG_SHEET, JUAL_SHEETS, LEDGER_JUAL_PREFIX,
    MIGRASI_PREFIX, PINDAH_SHEET, REQUIRED_JUAL_SHEET, STOK_SHEET,
)


def resolve_sheet(fp, want: str) -> str:
    """Dual-name: return the present variant of sheet `want`, accepting both the
    de-branded name ('Stok', 'JualShopee', ...) and the legacy 'Bisa'-prefixed name
    ('BisaStok', 'BisaJualShopee', ...). Returns `want` if neither is present so the
    read raises a clear error."""
    if isinstance(fp, pd.ExcelFile):
        names = fp.sheet_names
    else:
        with pd.ExcelFile(fp) as _xl:   # close the handle (avoid ResourceWarning)
            names = _xl.sheet_names
    if want in names:
        return want
    alt = want[4:] if want.startswith("Bisa") else "Bisa" + want
    return alt if alt in names else want


def _normalize_sku(s: pd.Series) -> pd.Series:
    """Case/space-normalize SKU so case-only variants (e.g. PCB-5X7 vs 5x7) merge,
    matching the case-insensitive SUMIF in the Google Sheets rekap."""
    return s.astype(str).str.upper().str.strip()


def _is_migrasi(s: pd.Series) -> pd.Series:
    return s.astype(str).str.startswith(MIGRASI_PREFIX, na=False)


def _drop_duplicate_migrasi(df: pd.DataFrame) -> pd.DataFrame:
    """Drop Migrasi rows for SKUs that also have non-Migrasi (real purchase) data.
    Keep Migrasi rows where they're the only source of HPP info for the SKU."""
    is_mig = _is_migrasi(df["toko"])
    skus_with_real = set(df.loc[~is_mig, "SKU"].unique())
    to_drop = is_mig & df["SKU"].isin(skus_with_real)

    n_drop = int(to_drop.sum())
    n_keep_mig = int((is_mig & ~to_drop).sum())
    if n_drop or n_keep_mig:
        print(f"  → Migrasi: drop {n_drop:,} duplikat (SKU sudah ada di non-Migrasi), "
              f"keep {n_keep_mig:,} (satu-satunya sumber HPP)")
    return df[~to_drop].copy()


def load_stok_files(file_paths: list[Path]) -> pd.DataFrame:
    """Load and combine stok files. Dedupe on (SKU, date, qty, total_hpp)."""
    if not file_paths:
        raise FileNotFoundError("Tidak ada file Stok ditemukan di data/")

    parts = []
    for fp in file_paths:
        print(f"✓ Membaca stok: {fp.name}")
        df = pd.read_excel(fp, sheet_name=resolve_sheet(fp, STOK_SHEET), header=1)
        # Backward-compat: pre-standardization exports name the supplier column
        # "Toko[spasi]Akun Pemesan"; the standardized export uses "Toko".
        if COL_STOK_TOKO_LEGACY in df.columns and COL_STOK_TOKO not in df.columns:
            df = df.rename(columns={COL_STOK_TOKO_LEGACY: COL_STOK_TOKO})
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
        COL_STOK_TANGGAL_SAMPAI: "tanggal_sampai",
        COL_STOK_TOKO: "toko",
        COL_STOK_LUAR_NEGERI: "luar_negeri",
    })
    for col in ["toko", "luar_negeri", "tanggal_bayar", "tanggal_sampai"]:
        if col not in df.columns:
            df[col] = np.nan

    raw_count = len(df)
    df = df[df["SKU"].notna()].copy()
    df["SKU"] = _normalize_sku(df["SKU"])
    df["qty_beli"] = pd.to_numeric(df["qty_beli"], errors="coerce")
    df["total_hpp"] = pd.to_numeric(df["total_hpp"], errors="coerce")
    df["tanggal_bayar"] = pd.to_datetime(df["tanggal_bayar"], errors="coerce")
    df["tanggal_sampai"] = pd.to_datetime(df["tanggal_sampai"], errors="coerce")
    df["luar_negeri"] = pd.to_numeric(df["luar_negeri"], errors="coerce")
    df = df.dropna(subset=["qty_beli", "total_hpp"])
    print(f"  → Stok bersih: {len(df):,} baris (dari {raw_count:,} mentah)")

    # NOTE: the old drop_duplicates(SKU,tanggal_bayar,qty_beli,total_hpp) was REMOVED.
    # It deleted genuine repeated identical lots (e.g. 3× JUMPER-MF-20CM at the same
    # timestamp/qty/price = 6000 pcs → kept only 2000), undercounting beli by 14,258
    # pcs across 15 SKU and corrupting HPP weights. The Sheets rekap never dedups; we
    # only drop Migrasi rows that duplicate real purchases (below).
    df = _drop_duplicate_migrasi(df)
    return df


def load_jual_files(file_paths: list[Path]) -> pd.DataFrame:
    """Load all available jual sheets from all files."""
    if not file_paths:
        raise FileNotFoundError("Tidak ada file Jual ditemukan di data/")

    parts = []
    for fp in file_paths:
        with pd.ExcelFile(fp) as xl:   # context-managed so the handle always closes
            if resolve_sheet(xl, REQUIRED_JUAL_SHEET) not in xl.sheet_names:
                print(f"  ⚠ {fp.name} tidak punya sheet wajib '{REQUIRED_JUAL_SHEET}', dilewati")
                continue
            print(f"✓ Membaca jual: {fp.name}")
            for sheet in JUAL_SHEETS:
                actual = resolve_sheet(xl, sheet)
                if actual in xl.sheet_names:
                    df = xl.parse(sheet_name=actual)
                    df["_sheet_source"] = sheet   # canonical (new) name, so downstream filters match
                    parts.append(df)
                    print(f"  → {actual}: {len(df):,} baris")

    df = pd.concat(parts, ignore_index=True)
    df = df.rename(columns={
        COL_JUAL_QTY: "qty_jual",
        COL_JUAL_OMZET: "omzet",
        COL_JUAL_KODE_UNIK: "kode_unik",
        COL_JUAL_TAMBAHAN: "tambahan",
        COL_JUAL_TANGGAL: "tanggal_pesan",
        COL_JUAL_AKUN: "akun_penjual",
    })
    df = df[df["SKU"].notna()].copy()
    df["SKU"] = _normalize_sku(df["SKU"])
    return df


def clean_jual(df: pd.DataFrame, year: int | None = None) -> tuple[pd.DataFrame, dict]:
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


# ============================================================================
# Current-workbook loaders for the stock ledger (reconcile to RekapBarang).
# These read the LATEST stok+jual file by filename and reproduce the rekap formula.
# ============================================================================

def latest_file(file_paths: list[Path]) -> Path:
    """Return the latest file by filename (e.g. Stok_2026.xlsx > Stok_2018_2025.xlsx)."""
    if not file_paths:
        raise FileNotFoundError("Tidak ada file untuk menentukan workbook berjalan")
    return sorted(file_paths)[-1]


def load_current_stok_arrived(stok_file: Path) -> pd.DataFrame:
    """Arrived purchases (Tanggal Sampai filled) from the current stok workbook.
    Keeps Migrasi rows (opening balance). Returns columns: SKU, gudang, qty."""
    df = pd.read_excel(stok_file, sheet_name=resolve_sheet(stok_file, STOK_SHEET), header=1)
    df = df[df["SKU"].notna()].copy()
    df["SKU"] = _normalize_sku(df["SKU"])
    df["qty"] = pd.to_numeric(df[COL_STOK_QTY], errors="coerce").fillna(0)
    df["sampai"] = pd.to_datetime(df[COL_STOK_TANGGAL_SAMPAI], errors="coerce")
    df["gudang"] = df[COL_STOK_ALAMAT].astype(str).str.strip()
    arrived = df[df["sampai"].notna()].copy()   # filled = sudah sampai (in stock)
    return arrived[["SKU", "gudang", "qty"]]


def load_current_jual_nonvoid(jual_file: Path) -> pd.DataFrame:
    """Non-void sales from the current jual workbook, all Jual* sheets (matches
    rekap scope incl. Blibli/Investasi). Returns columns: SKU, gudang, qty."""
    parts = []
    with pd.ExcelFile(jual_file) as xl:   # context-managed so the handle always closes
        sheets = [s for s in xl.sheet_names
                  if s.startswith(LEDGER_JUAL_PREFIX) or s.startswith("Bisa" + LEDGER_JUAL_PREFIX)]
        for sh in sheets:
            d = xl.parse(sheet_name=sh)
            if "SKU" not in d.columns:
                continue
            parts.append(d)
    df = pd.concat(parts, ignore_index=True)
    df = df[df["SKU"].notna()].copy()
    df["SKU"] = _normalize_sku(df["SKU"])
    df["qty"] = pd.to_numeric(df[COL_JUAL_QTY], errors="coerce").fillna(0)
    if "Void" in df.columns:
        df = df[df["Void"] != True].copy()
    df["gudang"] = df[COL_JUAL_GUDANG].astype(str).str.strip() \
        if COL_JUAL_GUDANG in df.columns else ""
    return df[["SKU", "gudang", "qty"]]


def load_hilang(stok_file: Path) -> pd.DataFrame:
    """Hilang from the current stok workbook. Returns: SKU, gudang, ketemu, hilang."""
    try:
        df = pd.read_excel(stok_file, sheet_name=resolve_sheet(stok_file, HILANG_SHEET), header=0)
    except ValueError:
        return pd.DataFrame(columns=["SKU", "gudang", "ketemu", "hilang"])
    df = df[df[COL_HILANG_SKU].notna()].copy()
    df["SKU"] = _normalize_sku(df[COL_HILANG_SKU])
    df["ketemu"] = pd.to_numeric(df[COL_HILANG_KETEMU], errors="coerce").fillna(0)
    df["hilang"] = pd.to_numeric(df[COL_HILANG_HILANG], errors="coerce").fillna(0)
    df["gudang"] = df[COL_HILANG_GUDANG].astype(str).str.strip()
    return df[["SKU", "gudang", "ketemu", "hilang"]]


def load_pindah(stok_file: Path) -> pd.DataFrame:
    """PindahBarang from the current stok workbook (inter-warehouse transfers).
    Returns: SKU, gudang_in (receives +qty), gudang_out (loses −qty), qty."""
    try:
        df = pd.read_excel(stok_file, sheet_name=resolve_sheet(stok_file, PINDAH_SHEET), header=0)
    except ValueError:
        return pd.DataFrame(columns=["SKU", "gudang_in", "gudang_out", "qty"])
    if COL_PINDAH_SKU not in df.columns:
        return pd.DataFrame(columns=["SKU", "gudang_in", "gudang_out", "qty"])
    df = df[df[COL_PINDAH_SKU].notna()].copy()
    df["SKU"] = _normalize_sku(df[COL_PINDAH_SKU])
    df["qty"] = pd.to_numeric(df[COL_PINDAH_QTY], errors="coerce").fillna(0)
    df["gudang_in"] = df[COL_PINDAH_TAMBAH].astype(str).str.strip()
    df["gudang_out"] = df[COL_PINDAH_KURANG].astype(str).str.strip()
    df = df[df["qty"] > 0]
    return df[["SKU", "gudang_in", "gudang_out", "qty"]]
