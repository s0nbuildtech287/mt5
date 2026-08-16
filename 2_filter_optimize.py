#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CANH BAO: FILE NAY CHI LA LIBRARY - KHONG CHAY RIENG LE!

Hay dung:
    python 0_run_optimized.py

(Chạy file này sẽ tạo file CSV filtered - gây lẫn lộn)

============================================================================
UPDATE: Tich hop thuat toan tu "mt5_optimize_filter.py" vao pipeline goc.
So voi ban cu, file nay them:
  1) Risk Hard Flags   - loai VINH VIEN cac bo tham so vi pham dieu kien rui ro
                          co ban (Payoff<=0, Profit Factor qua thap, Recovery
                          Factor <1, P/D Ratio <1) TRUOC KHI cham diem/xep hang.
  2) Composite score    - diem tong hop = trung binh co trong so cua z-score
                          Profit, Profit Factor, Recovery Factor, Sharpe,
                          Expected Payoff, Trades, tru z-score Equity DD%.
  3) Plateau robustness - do on dinh vung tham so: 1 bo tham so "dep" nhung
                          nam co lap (xung quanh toan bo xau) se bi nghi ngo
                          overfit va giam diem.
  4) Dedup khi chon Top N - tranh chon nhieu bo tham so gan giong het nhau.
  5) Ho tro doc CA HAI dinh dang file ket qua toi uu hoa cua MT5:
       - CSV thuong (dinh dang cu, pipeline nay dang dung)
       - XML "Excel Spreadsheet" (<?mso-application progid="Excel.Sheet"?>)
         - dinh dang ma "mt5_optimize_filter.py" ho tro.
     -> Neu sau nay ban doi cach xuat ket qua Optimize trong MT5 (vi du sang
        XML), pipeline VAN chay binh thuong ma khong can sua gi them.
  6) Ngoai file "<ten>_filtered.csv" cho tung EA/bot (de File 3 dung tiep nhu
     cu), con xuat them 1 file tong hop "all_bots_summary.csv" gop toan bo
     ung vien da loc cua TAT CA cac bot, xep hang chung theo diem cuoi cung
     (_final_score) - tien de ban so sanh nhanh giua cac bot voi nhau.

TAT CA nguong loc (Profit/Trades/DD/...) GIU NGUYEN gia tri mac dinh nhu
pipeline cu (khong doi hanh vi loc co ban), chi cong them lop cham diem +
hard-flag + plateau o tren. Co the tuy chinh cac nguong nay bang cach khai
bao them bien tuong ung trong config.py (xem phan CONFIG ben duoi), khong
bat buoc phai sua file nay.
"""

# PRINT CẢNH BÁO NGAY KHI RUN
import sys
if __name__ == "__main__":
    print("KHONG CHAY FILE NAY RIENG LE!")
    print("Hay chay: python 0_run_optimized.py")
    print("File nay chi la library de file 0 import")
    sys.exit(1)

import os
import numpy as np
import pandas as pd
import xml.etree.ElementTree as ET
import config  # Import cấu hình chung


# =============================================================================
# CONFIG - co the ghi de bang cach dinh nghia them trong config.py, vi du:
#   FILTER_MIN_PROFIT = 15000
# Neu khong dinh nghia, dung gia tri mac dinh (giong pipeline cu) ben duoi.
# =============================================================================
def _cfg(name, default):
    return getattr(config, name, default)


MIN_PROFIT = _cfg("FILTER_MIN_PROFIT", 10000)
MIN_TRADES = _cfg("FILTER_MIN_TRADES", 100)
MAX_TRADES = _cfg("FILTER_MAX_TRADES", 1200)
MIN_RECOVERY_FACTOR = _cfg("FILTER_MIN_RECOVERY_FACTOR", 2)
MIN_PROFIT_FACTOR = _cfg("FILTER_MIN_PROFIT_FACTOR", 1)
MAX_PROFIT_FACTOR = _cfg("FILTER_MAX_PROFIT_FACTOR", 4)
MIN_EQUITY_DD = _cfg("FILTER_MIN_EQUITY_DD", 8)
MAX_EQUITY_DD = _cfg("FILTER_MAX_EQUITY_DD", 13)

TOP_N = _cfg("FILTER_TOP_N", 10)
DEDUP_TOL = _cfg("FILTER_DEDUP_TOL", 0.15)
PLATEAU_TOL = _cfg("FILTER_PLATEAU_TOL", 0.15)

# Risk Hard Flags - bo THAM SO NAO vi pham 1 trong cac dieu kien duoi day se
# bi loai VINH VIEN, khong dua vao xep hang/diem so nua, du Profit/DD dep den may.
HARD_FLAG_MIN_PAYOFF = _cfg("FILTER_HARD_MIN_PAYOFF", 0.0)
HARD_FLAG_MIN_PF = _cfg("FILTER_HARD_MIN_PF", 1.2)
HARD_FLAG_MIN_PF_TRADES = _cfg("FILTER_HARD_MIN_PF_TRADES", 100)
HARD_FLAG_MIN_RF = _cfg("FILTER_HARD_MIN_RF", 1.0)
HARD_FLAG_MIN_PD = _cfg("FILTER_HARD_MIN_PD", 1.0)
HARD_FLAG_MIN_PD_TRADES = _cfg("FILTER_HARD_MIN_PD_TRADES", 100)

NS = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}

# 10 cot chi so co dinh MT5 luon xuat theo dung thu tu nay
FIXED_METRIC_COLS = [
    "Pass", "Result", "Profit", "Expected Payoff", "Profit Factor",
    "Recovery Factor", "Sharpe Ratio", "Custom", "Equity DD %", "Trades",
]

# Cac cot noi bo do script nay tao ra - KHONG duoc coi la tham so chien luoc
# (quan trong de File 3 khong nham lay nhung cot nay lam Input EA)
INTERNAL_COLS = [
    "bot", "source_file", "_score", "_plateau_score", "_final_score",
    "_pd_ratio", "_flag_reason", "DD_Deviation", "Result_Priority",
]


# =============================================================================
# DOC FILE KET QUA OPTIMIZE (ho tro ca CSV va XML Excel Spreadsheet cua MT5)
# =============================================================================
def _looks_like_mt5_xml(file_path):
    """Doan dinh dang file bang cach doc vai trieu byte dau, khong dua hoan
    toan vao duoi file (phong khi Report= trong .ini khong co duoi .xml)."""
    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            head = f.read(1000)
        return "urn:schemas-microsoft-com:office:spreadsheet" in head
    except Exception:
        return False


def parse_mt5_xml(file_path):
    """Doc file XML Spreadsheet cua MT5 Optimize Result -> DataFrame."""
    tree = ET.parse(file_path)
    root = tree.getroot()

    title = None
    deposit = None
    doc_props = root.find("ss:DocumentProperties", NS)
    if doc_props is not None:
        t = doc_props.find("ss:Title", NS)
        if t is not None:
            title = t.text
        dep = doc_props.find("ss:Deposit", NS)
        if dep is not None and dep.text:
            digits = "".join(ch for ch in dep.text if (ch.isdigit() or ch == "."))
            if digits:
                deposit = float(digits)

    ws = root.find("ss:Worksheet", NS)
    table = ws.find("ss:Table", NS)
    rows = table.findall("ss:Row", NS)
    if not rows:
        raise ValueError(f"Khong tim thay du lieu trong {file_path}")

    header = [c.find("ss:Data", NS).text for c in rows[0].findall("ss:Cell", NS)]

    data_rows = []
    for row in rows[1:]:
        vals = []
        for cell in row.findall("ss:Cell", NS):
            data = cell.find("ss:Data", NS)
            vals.append(data.text if data is not None else None)
        if len(vals) == len(header):
            data_rows.append(vals)

    df = pd.DataFrame(data_rows, columns=header)
    df.attrs["title"] = title
    df.attrs["deposit"] = deposit
    return df


def parse_mt5_csv(file_path):
    """Doc file CSV tu MT5 optimization (dinh dang goc cua pipeline nay)."""
    df = pd.read_csv(file_path, encoding="utf-8-sig")
    df.attrs["title"] = None
    df.attrs["deposit"] = None
    return df


def load_report_file(file_path):
    """Tu dong nhan dien CSV hay XML roi doc + ep kieu so cho cac cot chi so."""
    if file_path.lower().endswith(".xml") or _looks_like_mt5_xml(file_path):
        df = parse_mt5_xml(file_path)
    else:
        df = parse_mt5_csv(file_path)

    for col in FIXED_METRIC_COLS:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace("%", "").str.replace(" ", "")
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Cac cot con lai la THAM SO CUA CHIEN LUOC (Input EA) - luon la so.
    # Voi file XML, du lieu doc vao la text nen bat buoc phai ep kieu o day;
    # voi file CSV, pandas thuong da tu nhan dien so nen dong nay khong lam
    # thay doi gi (chi phong ngua truong hop cot bi doc thanh chuoi).
    for col in df.columns:
        if col not in FIXED_METRIC_COLS:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


# =============================================================================
# CHAM DIEM: composite score + plateau robustness (tu mt5_optimize_filter.py)
# =============================================================================
def get_param_cols(df):
    return [c for c in df.columns if c not in FIXED_METRIC_COLS and c not in INTERNAL_COLS
            and not c.startswith("_")]


def zscore(s):
    std = s.std()
    if std == 0 or pd.isna(std):
        return pd.Series(0, index=s.index)
    return (s - s.mean()) / std


def compute_composite_score(df):
    z_profit = zscore(df["Profit"])
    z_pf = zscore(df["Profit Factor"])
    z_rf = zscore(df["Recovery Factor"])
    z_sharpe = zscore(df["Sharpe Ratio"]) if "Sharpe Ratio" in df.columns else 0
    z_payoff = zscore(df["Expected Payoff"]) if "Expected Payoff" in df.columns else 0
    z_trades = zscore(df["Trades"])
    z_dd = zscore(df["Equity DD %"])

    return (
        z_profit * 1.0
        + z_pf * 1.5
        + z_rf * 1.5
        + z_sharpe * 1.0
        + z_payoff * 0.5
        + z_trades * 0.3
        - z_dd * 0.5
    )


def param_distance(row1, row2, param_cols):
    diffs = []
    for c in param_cols:
        v1, v2 = row1[c], row2[c]
        if pd.isna(v1) or pd.isna(v2):
            continue
        scale = max(abs(v1), abs(v2), 1)
        diffs.append(abs(v1 - v2) / scale)
    if not diffs:
        return 1.0
    return float(np.mean(diffs))


def compute_plateau_score(df, param_cols, tol, max_rows_for_full_search=4000):
    """Diem trung binh cua cac 'hang xom' (bo tham so lech nhau < tol) tinh
    theo composite score - phat hien diem nhon co lap (nghi ngo overfit)."""
    n = len(df)
    scores = df["_score"].values
    plateau = np.zeros(n)
    if n == 0 or not param_cols:
        return pd.Series(scores if n else plateau, index=df.index)

    params = df[param_cols].astype(float).values
    scale = np.maximum(np.abs(params).max(axis=0), 1)

    if n > max_rows_for_full_search:
        sample_idx = np.random.choice(n, size=max_rows_for_full_search, replace=False)
    else:
        sample_idx = np.arange(n)

    for i in range(n):
        diffs = np.abs(params[sample_idx] - params[i]) / scale
        rel_diff = diffs.mean(axis=1)
        neighbor_mask = rel_diff < tol
        neighbor_mask[sample_idx == i] = False
        if neighbor_mask.sum() > 0:
            plateau[i] = scores[sample_idx][neighbor_mask].mean()
        else:
            plateau[i] = scores[i]

    return pd.Series(plateau, index=df.index)


def dedup_select(df_sorted, param_cols, top_n, dedup_tol):
    """Chon top N nhung loai cac ung vien qua giong nhau ve tham so."""
    selected_idx = []
    selected_rows = []
    for idx, row in df_sorted.iterrows():
        close = any(
            param_distance(row, sel, param_cols) < dedup_tol for sel in selected_rows
        )
        if not close:
            selected_idx.append(idx)
            selected_rows.append(row)
        if len(selected_idx) >= top_n:
            break
    return df_sorted.loc[selected_idx]


def apply_hard_flags(df, deposit=None):
    """Loai vinh vien cac dong vi pham dieu kien rui ro cung, TRUOC khi loc
    theo 6 tieu chi va cham diem."""
    dep = deposit if deposit else 100_000.0

    dd_abs = (df["Equity DD %"] / 100.0) * dep
    with np.errstate(divide="ignore", invalid="ignore"):
        pd_ratio = np.where(dd_abs > 0, df["Profit"] / dd_abs, np.inf)
    df = df.copy()
    df["_pd_ratio"] = pd_ratio

    flag_payoff = (df["Expected Payoff"] <= HARD_FLAG_MIN_PAYOFF) if "Expected Payoff" in df.columns else False
    flag_pf = (df["Trades"] >= HARD_FLAG_MIN_PF_TRADES) & (df["Profit Factor"] < HARD_FLAG_MIN_PF)
    flag_rf = df["Recovery Factor"] < HARD_FLAG_MIN_RF
    flag_pd = (df["Trades"] >= HARD_FLAG_MIN_PD_TRADES) & (df["_pd_ratio"] < HARD_FLAG_MIN_PD)

    df["_flag_reason"] = ""
    if isinstance(flag_payoff, pd.Series):
        df.loc[flag_payoff, "_flag_reason"] += "Expected Payoff<=0;"
    df.loc[flag_pf, "_flag_reason"] += f"Profit Factor<{HARD_FLAG_MIN_PF};"
    df.loc[flag_rf, "_flag_reason"] += f"Recovery Factor<{HARD_FLAG_MIN_RF};"
    df.loc[flag_pd, "_flag_reason"] += f"P/D Ratio<{HARD_FLAG_MIN_PD};"

    any_flag = flag_pf | flag_rf | flag_pd
    if isinstance(flag_payoff, pd.Series):
        any_flag = any_flag | flag_payoff

    n_before = len(df)
    kept = df[~any_flag].copy()
    n_flagged = n_before - len(kept)
    if n_flagged > 0:
        print(f"   [Risk Hard Flags] loai {n_flagged}/{n_before} bo do vi pham dieu kien rui ro cung.")
    return kept


def filter_data(df, deposit=None):
    """Pipeline loc day du: Hard Flags -> 6 tieu chi goc -> cham diem
    (composite + plateau) -> chon Top N co dedup."""

    # BUOC 1: Risk Hard Flags - loai vinh vien truoc khi lam bat cu gi khac
    df = apply_hard_flags(df, deposit)
    if df.empty:
        return df

    # BUOC 2: GIU NGUYEN 6 TIEU CHI LOC GOC CUA BAN
    condition = (
        (df["Profit"] > MIN_PROFIT) &
        (df["Trades"] > MIN_TRADES) & (df["Trades"] < MAX_TRADES) &
        (df["Recovery Factor"] > MIN_RECOVERY_FACTOR) &
        (df["Profit Factor"] >= MIN_PROFIT_FACTOR) & (df["Profit Factor"] <= MAX_PROFIT_FACTOR) &
        (df["Equity DD %"] >= MIN_EQUITY_DD) & (df["Equity DD %"] <= MAX_EQUITY_DD)
    )
    filtered_df = df[condition].copy()
    if filtered_df.empty:
        return filtered_df

    # BUOC 3: Cham diem tong hop + do on dinh vung tham so
    param_cols = get_param_cols(df)
    filtered_df["_score"] = compute_composite_score(filtered_df)
    filtered_df["_plateau_score"] = compute_plateau_score(filtered_df, param_cols, PLATEAU_TOL)
    filtered_df["_final_score"] = (
        0.6 * zscore(filtered_df["_score"]) + 0.4 * zscore(filtered_df["_plateau_score"])
    )

    # BUOC 4: Xep hang theo diem cuoi cung, chon Top N co loai trung lap
    filtered_df = filtered_df.sort_values(by="_final_score", ascending=False)
    return dedup_select(filtered_df, param_cols, TOP_N, DEDUP_TOL)


# =============================================================================
# MAIN
# =============================================================================
def main():
    if not os.path.exists(config.OPTIMIZE_XML_DIR):
        print(f"Khong tim thay thu muc CSV: {config.OPTIMIZE_XML_DIR}")
        return

    result_files = [
        f for f in os.listdir(config.OPTIMIZE_XML_DIR)
        if f.lower().endswith((".csv", ".xml"))
    ]
    print(f"Bat dau loc {len(result_files)} file ket qua optimize...")

    all_tops = []
    for file_name in result_files:
        input_path = os.path.join(config.OPTIMIZE_XML_DIR, file_name)
        try:
            # 1. Doc du lieu (tu dong nhan dien CSV / XML)
            df = load_report_file(input_path)
            deposit = df.attrs.get("deposit")
            title = df.attrs.get("title") or os.path.splitext(file_name)[0]

            # 2. Loc + cham diem theo pipeline moi
            final_df = filter_data(df, deposit)

            # 3. Xuat ket qua CSV vao thu muc Filter cua config
            if final_df is not None and not final_df.empty:
                final_df = final_df.copy()
                final_df.insert(0, "bot", title)
                final_df.insert(1, "source_file", file_name)

                output_name = os.path.splitext(file_name)[0] + "_filtered.csv"
                output_path = os.path.join(config.FILTERED_CSV_DIR, output_name)

                final_df.to_csv(output_path, index=False, encoding="utf-8-sig")
                print(f"Thanh cong: {file_name} -> Da chon {len(final_df)} bo tot nhat.")
                print(f"   Luu vao: {output_name}")

                all_tops.append(final_df)
            else:
                print(f"Bo qua: {file_name} (Khong co bo nao thoa man tieu chi / deu bi Hard Flag loai)")

        except Exception as e:
            print(f"Loi xu ly file {file_name}: {e}")

    # 4. Bang tong hop xep hang chung tat ca cac bot
    if all_tops:
        summary = pd.concat(all_tops, ignore_index=True)
        summary = summary.sort_values("_final_score", ascending=False)
        summary_path = os.path.join(config.FILTERED_CSV_DIR, "all_bots_summary.csv")
        summary.to_csv(summary_path, index=False, encoding="utf-8-sig")
        print(f"\nDa gop {len(all_tops)} bot -> {summary_path}")


if __name__ == "__main__":
    main()