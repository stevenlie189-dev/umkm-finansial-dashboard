#!/usr/bin/env python3
"""Validasi logika workbook: hitung ulang formula dengan LibreOffice (headless) dan bandingkan
dengan perhitungan manual independen (pandas) pada beberapa skenario.

Pakai:  python3 tools/validate.py            (membutuhkan LibreOffice 'soffice')
"""
import subprocess, shutil, sys, datetime as dt
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
WORK = ROOT / "tools/_work/val"; WORK.mkdir(parents=True, exist_ok=True)
SRC = ROOT / "ROTI_RONA_Pembukuan.xlsx"
CSV = pd.read_csv(ROOT / "sample_transactions.csv", parse_dates=["Tanggal"])
CSV.columns = ["tgl", "jenis", "kat", "ket", "metode", "nom"]
FIRST = 5
BULAN = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"]

def recalc(path: Path) -> Path:
    out = WORK / ("calc_" + path.stem); out.mkdir(exist_ok=True)
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(out), str(path)],
                   check=True, capture_output=True, timeout=300)
    return out / path.name

def expected(df, year, bulan, saldo):
    if bulan == "Semua Bulan": a, b = dt.datetime(year,1,1), dt.datetime(year,12,31,23,59)
    else:
        m = BULAN.index(bulan) + 1; a = dt.datetime(year,m,1)
        b = (pd.Timestamp(year, m, 1) + pd.offsets.MonthEnd(0)).to_pydatetime().replace(hour=23, minute=59)
    d = df.dropna(subset=["tgl"])
    inp = d[(d.tgl >= a) & (d.tgl <= b)]
    pend = inp[inp.jenis == "Pendapatan"].nom.sum(); peng = inp[inp.jenis == "Pengeluaran"].nom.sum()
    upto = d[d.tgl <= b]
    kas = saldo + upto[upto.jenis == "Pendapatan"].nom.sum() - upto[upto.jenis == "Pengeluaran"].nom.sum()
    mon = []
    for m in range(1, 13):
        x = d[(d.tgl.dt.year == year) & (d.tgl.dt.month == m)]
        mon.append((x[x.jenis == "Pendapatan"].nom.sum(), x[x.jenis == "Pengeluaran"].nom.sum()))
    cats = inp[inp.jenis == "Pengeluaran"].groupby("kat").nom.sum().to_dict()
    return dict(pend=pend, peng=peng, laba=pend - peng, kas=kas, mon=mon, cats=cats)

def read_dash(path):
    wb = load_workbook(path, data_only=True); ws = wb["DASHBOARD"]
    mon = [(ws.cell(r, 3).value or 0, ws.cell(r, 4).value or 0) for r in range(36, 48)]
    cats = {ws.cell(r, 8).value: ws.cell(r, 10).value for r in range(36, 48) if ws.cell(r, 8).value}
    return dict(pend=ws["B8"].value, peng=ws["E8"].value, laba=ws["H8"].value, kas=ws["K8"].value,
                status=ws["B13"].value, cek=ws["K13"].value, mon=mon, cats=cats,
                match=ws["L48"].value, tot_pend=ws["C48"].value, tot_peng=ws["D48"].value)

fails = 0
def check(nm, cond, detail=""):
    global fails
    print(("  PASS " if cond else "  FAIL ") + nm + (f"  [{detail}]" if detail and not cond else ""))
    fails += (not cond)

def scenario(name, edit, year, bulan, saldo=15_000_000, expect_cek=0):
    """edit(ws_trx, ws_dash, df) -> df baru (mirror pandas)."""
    print(f"\n== {name} ==")
    wb = load_workbook(SRC); trx, dash = wb["TRANSAKSI"], wb["DASHBOARD"]
    df = edit(trx, dash, CSV.copy())
    dash["C5"] = year; dash["E5"] = bulan; dash["I5"] = saldo
    p = WORK / (name.split(":")[0].replace(" ", "_") + ".xlsx"); wb.save(p)
    got = read_dash(recalc(p)); exp = expected(df, year, bulan, saldo)
    for k in ("pend", "peng", "laba", "kas"):
        check(f"{k}: {got[k]:,.0f} == {exp[k]:,.0f}", abs(got[k] - exp[k]) < 0.5, f"{got[k]} vs {exp[k]}")
    check("12 bulan (pend & peng) cocok", all(abs(g[0]-e[0]) < .5 and abs(g[1]-e[1]) < .5 for g, e in zip(got["mon"], exp["mon"])))
    check("komposisi kategori cocok", all(abs(got["cats"].get(k, 0) - v) < .5 for k, v in exp["cats"].items()) and
          abs(sum(got["cats"].values()) - exp["peng"]) < .5)
    check("laba = pendapatan - pengeluaran", abs(got["laba"] - (got["pend"] - got["peng"])) < .5)
    check(f"transaksi perlu dicek == {expect_cek}", got["cek"] == expect_cek, str(got["cek"]))
    check("status: " + str(got["status"]), got["status"] is not None and "#" not in str(got["status"]))
    return got

def baseline(trx, dash, df): return df
def write_row(trx, r, tgl, jenis, kat, ket, metode, nom):
    trx.cell(r, 1).value = tgl; trx.cell(r, 3).value = jenis; trx.cell(r, 4).value = kat
    trx.cell(r, 5).value = ket; trx.cell(r, 6).value = metode; trx.cell(r, 7).value = nom
def clear_row(trx, r):
    for c in (1, 3, 4, 5, 6, 7): trx.cell(r, c).value = None
N = len(CSV); NEXT = FIRST + N   # baris kosong pertama

def add_income(trx, dash, df):
    write_row(trx, NEXT, dt.datetime(2026,10,31), "Pendapatan", "Penjualan Produk", "uji", "QRIS", 1_000_000)
    return pd.concat([df, pd.DataFrame([[pd.Timestamp(2026,10,31), "Pendapatan", "Penjualan Produk", "uji", "QRIS", 1_000_000]], columns=df.columns)])
def add_expense(trx, dash, df):
    d = add_income(trx, dash, df)
    write_row(trx, NEXT + 1, dt.datetime(2026,10,30), "Pengeluaran", "Marketing", "uji", "Tunai", 2_500_000)
    return pd.concat([d, pd.DataFrame([[pd.Timestamp(2026,10,30), "Pengeluaran", "Marketing", "uji", "Tunai", 2_500_000]], columns=df.columns)])
def edit_nominal(trx, dash, df):
    trx.cell(FIRST, 7).value = 5_000_000; df = df.copy(); df.loc[0, "nom"] = 5_000_000; return df
def delete_trx(trx, dash, df):
    clear_row(trx, FIRST + 1); return df.drop(index=1)       # baris 6 = pembelian bahan baku 2025-01-01
def far_row(trx, dash, df):
    write_row(trx, 6000, dt.datetime(2026,10,15), "Pendapatan", "Penjualan Produk", "jauh", "Tunai", 700_000)
    return pd.concat([df, pd.DataFrame([[pd.Timestamp(2026,10,15), "Pendapatan", "Penjualan Produk", "jauh", "Tunai", 700_000]], columns=df.columns)])
def bad_category(trx, dash, df):
    write_row(trx, NEXT, dt.datetime(2026,10,31), "Pendapatan", "Sewa", "salah kategori", "Tunai", 100_000)
    return pd.concat([df, pd.DataFrame([[pd.Timestamp(2026,10,31), "Pendapatan", "Sewa", "x", "Tunai", 100_000]], columns=df.columns)])
def empty_all(trx, dash, df):
    for r in range(FIRST, FIRST + N): clear_row(trx, r)
    return df.iloc[0:0]

print("Memeriksa file:", SRC.name)
scenario("S0a: baseline 2026 semua bulan", baseline, 2026, "Semua Bulan")
scenario("S0b: baseline 2025 semua bulan", baseline, 2025, "Semua Bulan")
scenario("S1: ubah periode ke Maret 2026", baseline, 2026, "Maret")
scenario("S2: ubah periode ke Januari 2025 (bulan rugi)", baseline, 2025, "Januari")
scenario("S3: tambah pendapatan", add_income, 2026, "Oktober")
scenario("S4: tambah pengeluaran", add_expense, 2026, "Semua Bulan")
scenario("S5: ubah nominal", edit_nominal, 2025, "Januari")
scenario("S6: hapus transaksi", delete_trx, 2025, "Januari")
scenario("S7: baris di luar 5000 baris siap pakai (baris 6000)", far_row, 2026, "Oktober")
g = scenario("S8: kategori tidak cocok ditandai", bad_category, 2026, "Oktober", expect_cek=1)
g = scenario("S9: semua transaksi kosong (saldo = saldo awal)", empty_all, 2026, "Semua Bulan")
check("status BELUM ADA DATA", "BELUM ADA DATA" in str(g["status"]), str(g["status"]))
print("\nHASIL AKHIR:", "SEMUA LULUS" if fails == 0 else f"{fails} PEMERIKSAAN GAGAL")
sys.exit(1 if fails else 0)
