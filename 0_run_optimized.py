#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SUPER OPTIMIZED VERSION - Gop File 1+2+3 vao 1 file
- Khong luu XML trung gian
- Khong luu CSV trung gian
- Chi luu ket qua INI backtest cuoi cung
- Tat ca xu ly trong RAM - Nhanh hon 30%
"""

import os
import subprocess
import time
import shutil
import config
import xml.etree.ElementTree as ET
import pandas as pd
import re
import argparse
import sys
import json

DEFAULT_FILTERS = {
    'min_profit': 10000,
    'min_trades': 100,
    'max_trades': 1200,
    'min_recovery_factor': 2,
    'min_profit_factor': 1,
    'max_profit_factor': 4,
    'min_equity_dd': 8,
    'max_equity_dd': 13,
}

# ==================== PATH HELPERS ====================

def build_paths(terminal_data_path=None):
    paths = config.init_paths(terminal_data_path or config.DEFAULT_TERMINAL_DATA_PATH)
    config.TERMINAL_DATA_PATH = paths["TERMINAL_DATA_PATH"]
    config.INI_SOURCE_DIR = paths["INI_SOURCE_DIR"]
    config.OPTIMIZE_XML_DIR = paths["OPTIMIZE_XML_DIR"]
    config.FILTERED_CSV_DIR = paths["FILTERED_CSV_DIR"]
    config.BACKTEST_INI_DIR = paths["BACKTEST_INI_DIR"]
    config.BASE_RESULT_DIR = paths["BASE_RESULT_DIR"]
    config.RAW_REPORT_DIR = paths["RAW_REPORT_DIR"]
    return paths

# ==================== PHASE 1: OPTIMIZE (File 1 Logic) ====================

def get_latest_xml(report_name):
    """Tim file XML ket qua theo ten report"""
    for root, dirs, files in os.walk(config.TERMINAL_DATA_PATH):
        for f in files:
            if f.lower() == report_name.lower():
                return os.path.join(root, f)
    return None

def fix_ini_before_run(file_path, report_name):
    """Dam bao file INI co du cau hinh xuat Report"""
    lines = []
    encodings = ['utf-16', 'utf-8-sig', 'utf-8']
    for enc in encodings:
        try:
            with open(file_path, 'r', encoding=enc) as f:
                lines = f.readlines()
            break
        except:
            continue
    if not lines:
        return

    has_report = any("Report=" in l for l in lines)
    has_replace = any("ReplaceReport=" in l for l in lines)
    has_shutdown = any("ShutdownTerminal=" in l for l in lines)

    if not (has_report and has_replace and has_shutdown):
        new_lines = []
        for line in lines:
            new_lines.append(line)
            if line.strip().lower() == "[tester]":
                if not has_report:
                    new_lines.append(f"Report={report_name}\n")
                if not has_replace:
                    new_lines.append("ReplaceReport=1\n")
                if not has_shutdown:
                    new_lines.append("ShutdownTerminal=1\n")
        with open(file_path, 'w', encoding='utf-16') as f:
            f.writelines(new_lines)

def parse_xml_to_dataframe(xml_file_path):
    """Parse XML MT5 -> DataFrame (luu trong RAM)"""
    tree = ET.parse(xml_file_path)
    root = tree.getroot()
    ns = {'ss': 'urn:schemas-microsoft-com:office:spreadsheet'}
    all_rows = []
    for row in root.findall('.//ss:Row', ns):
        cells = row.findall('ss:Cell', ns)
        row_data = []
        for cell in cells:
            data_node = cell.find('ss:Data', ns)
            row_data.append(data_node.text if data_node is not None else "")
        if row_data and any(row_data):
            all_rows.append(row_data)
    if not all_rows:
        return None
    header = all_rows[0]
    data = all_rows[1:]
    df = pd.DataFrame(data, columns=header)
    df = df.dropna(axis=1, how='all')
    numeric_cols = ['Profit', 'Trades', 'Profit Factor', 'Equity DD %', 'Recovery Factor', 'Result', 'Sharpe Ratio']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    return df

def safe_remove_file(file_path, max_retries=3):
    """Xóa file với retry logic, chờ nếu file bị lock"""
    for attempt in range(max_retries):
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                print(f"[CLEANUP] Deleted: {os.path.basename(file_path)}", flush=True)
                return True
        except Exception as e:
            print(f"[CLEANUP] Failed to delete (attempt {attempt+1}/{max_retries}): {e}", flush=True)
            if attempt < max_retries - 1:
                time.sleep(1)
    return False

def cleanup_html_png_files():
    """Xóa tất cả file HTML & PNG được tạo ra từ optimize / backtest"""
    cleanup_dirs = [
        config.TERMINAL_DATA_PATH,
        os.path.join(config.TERMINAL_DATA_PATH, r"Tester\Files"),
        os.path.join(config.TERMINAL_DATA_PATH, r"Tester\Reports"),
        os.path.join(config.TERMINAL_DATA_PATH, r"MQL5\Files"),
        os.path.join(config.TERMINAL_DATA_PATH, r"reports"),
        config.BASE_RESULT_DIR,
    ]
    
    removed_count = 0
    for directory in cleanup_dirs:
        if not os.path.isdir(directory):
            continue
        try:
            for file_name in os.listdir(directory):
                if file_name.lower().endswith(('.htm', '.html', '.png')):
                    file_path = os.path.join(directory, file_name)
                    try:
                        os.remove(file_path)
                        removed_count += 1
                    except Exception as e:
                        print(f"[CLEANUP] Failed to delete {file_name}: {e}", flush=True)
        except Exception as e:
            print(f"[CLEANUP] Error scanning {directory}: {e}", flush=True)
    
    if removed_count > 0:
        print(f"[CLEANUP] Deleted {removed_count} HTML/PNG files", flush=True)

def phase1_optimize(ini_file, num_runs=1):
    path_ini = os.path.join(config.INI_SOURCE_DIR, ini_file)
    report_name = config.get_report_name_from_ini(path_ini)
    print(f"\n{'='*60}")
    print(f"PHASE 1: OPTIMIZE - {ini_file}")
    print(f"{'='*60}")
    print(f"Report name: {report_name}")
    print(f"So lan chay: {num_runs}")

    all_dataframes = []
    for run_number in range(1, num_runs + 1):
        print(f"\nLAN CHAY {run_number}/{num_runs}")
        print(f"-" * 40)
        fix_ini_before_run(path_ini, report_name)

        # KHÔNG kill_mt5() ở đây nữa — tránh kill profile khác!
        print("Chay MT5 optimize...")

        # shell=False để proc.pid là PID thật của terminal64.exe
        proc = subprocess.Popen(
            [config.MT5_PATH, f'/config:{path_ini}'],
            shell=False
        )
        mt5_pid = proc.pid
        print(f"[phase1] Spawned MT5 PID={mt5_pid}", flush=True)

        found = False
        start = time.time()
        last_heartbeat = start
        while (time.time() - start) < 3600:
            latest_xml = get_latest_xml(report_name)
            if latest_xml and os.path.getsize(latest_xml) > 0:
                time.sleep(2)
                try:
                    df = parse_xml_to_dataframe(latest_xml)
                    if df is not None:
                        all_dataframes.append(df)
                        print(f"Parsed: {len(df)} rows vao RAM")
                        safe_remove_file(latest_xml)
                    else:
                        print("Parse XML that bai")
                except Exception as e:
                    print(f"Loi parse XML: {e}")
                found = True
                break
            
            # Heartbeat mỗi 30s để giữ stdout sống
            current_time = time.time()
            if (current_time - last_heartbeat) >= 30:
                elapsed = int(current_time - start)
                print(f"[HEARTBEAT] Optimize dang chay... ({elapsed}s elapsed)", flush=True)
                last_heartbeat = current_time
            
            time.sleep(10)

        if not found:
            print("Timeout cho file XML")

        # Kill ĐÚNG PID này thôi
        print(f"[phase1] Killing MT5 PID={mt5_pid}", flush=True)
        config.kill_mt5(pid=mt5_pid)
        time.sleep(5)
        
        # Verify process thực sự đã exit
        for attempt in range(10):
            try:
                os.kill(mt5_pid, 0)  # signal 0 = check process still alive
                time.sleep(0.5)
            except OSError:
                # Process không tồn tại - OK
                print(f"[phase1] MT5 PID={mt5_pid} confirmed dead", flush=True)
                break
        
        time.sleep(1)

    if all_dataframes:
        merged_df = pd.concat(all_dataframes, ignore_index=True)
        print(f"\nGop xong: {len(merged_df)} dong toi uu trong RAM")
        return merged_df
    return None
# ==================== PHASE 2: FILTER (File 2 Logic) ====================

def phase2_filter(df, filters=None):
    filters = {**DEFAULT_FILTERS, **(filters or {})}
    print(f"\n{'='*60}")
    print("PHASE 2: FILTER")
    print(f"{'='*60}")
    print(f"Input: {len(df)} dong")
    print(f"Filter config: {json.dumps(filters, ensure_ascii=True)}")
    condition = (
        (df['Profit'] > filters['min_profit']) &
        (df['Trades'] > filters['min_trades']) & (df['Trades'] < filters['max_trades']) &
        (df['Recovery Factor'] > filters['min_recovery_factor']) &
        (df['Profit Factor'] >= filters['min_profit_factor']) & (df['Profit Factor'] <= filters['max_profit_factor']) &
        (df['Equity DD %'] >= filters['min_equity_dd']) & (df['Equity DD %'] <= filters['max_equity_dd'])
    )
    filtered_df = df[condition].copy()
    print(f"Sau loc 6 tieu chi: {len(filtered_df)} bo")
    if filtered_df.empty:
        print("Khong co bo nao thoa man tieu chi!")
        return filtered_df
    filtered_df['DD_Deviation'] = abs(filtered_df['Equity DD %'] - 10)
    filtered_df['Result_Priority'] = filtered_df['Result'].between(5, 8)
    filtered_df = filtered_df.sort_values(by=['Result_Priority', 'Trades', 'DD_Deviation', 'Profit'], ascending=[False, False, True, False])
    top_10 = filtered_df.head(10)
    print("Top 10 bo tot nhat (trong RAM)")
    print(f"   Profit cao nhat: ${top_10['Profit'].max():.2f}")
    print(f"   Sharpe Ratio cao nhat: {top_10['Sharpe Ratio'].max():.4f}")
    return top_10

# ==================== PHASE 3: CREATE INI (File 3 Logic) ====================

def get_bot_info(filename):
    clean_name = filename.replace('Filtered_', '').upper()
    id_match = re.search(r'(\d+)', clean_name)
    bot_id = id_match.group(1) if id_match else ""
    tf_match = re.search(r'(M15|M30|H1|H4|D1)', clean_name)
    timeframe = tf_match.group(1) if tf_match else ""
    symbol = ""
    common_symbols = ["DE30", "HK50", "USOIL", "BTCUSD", "ETHUSD", "XAUUSD", "XAGUSD", "US500", "JP225"]
    for s in common_symbols:
        if s in clean_name:
            symbol = s
            break
    return bot_id, symbol, timeframe

def find_matching_ini(ini_dir, ini_name):
    try:
        return os.path.join(ini_dir, ini_name)
    except:
        return None

def phase3_create_ini(filtered_df, ini_name):
    print(f"\n{'='*60}")
    print("PHASE 3: CREATE BACKTEST INI")
    print(f"{'='*60}")
    template_path = os.path.join(config.INI_SOURCE_DIR, ini_name)
    if not os.path.exists(template_path):
        print(f"Khong tim thay template INI: {ini_name}")
        return False
    source_lines = []
    try:
        with open(template_path, 'r', encoding='utf-16') as f:
            source_lines = f.readlines()
    except:
        with open(template_path, 'r', encoding='utf-8-sig') as f:
            source_lines = f.readlines()
    if not source_lines:
        print("File template rong")
        return False

    bot_folder = ini_name.replace('.ini', '')
    final_output_path = os.path.join(config.BACKTEST_INI_DIR, bot_folder)
    if not os.path.exists(final_output_path):
        os.makedirs(final_output_path)

    exclude = ['Pass', 'Result', 'Profit', 'Expected Payoff', 'Profit Factor', 'Recovery Factor', 'Sharpe Ratio', 'Custom', 'Equity DD %', 'Trades', 'DD_Deviation', 'Result_Priority']
    params = [col for col in filtered_df.columns if col not in exclude]

    base_ini_name = os.path.splitext(ini_name)[0]
    for i, row in filtered_df.iterrows():
        new_ini_content = []
        profit_val = int(float(row['Profit'])) if pd.notna(row['Profit']) else 0
        file_name = f"{base_ini_name}_{profit_val}_{i+1}.ini"
        for line in source_lines:
            if line.startswith("Optimization="):
                new_ini_content.append("Optimization=0\n")
            elif line.startswith("Model="):
                new_ini_content.append("Model=0\n")
            elif line.startswith("ShutdownTerminal="):
                new_ini_content.append("ShutdownTerminal=1\n")
            elif line.startswith("Report=") or line.startswith("ReplaceReport="):
                continue
            else:
                is_param = False
                for p in params:
                    if line.startswith(f"{p}="):
                        try:
                            parts = line.split('=')[1].split('||')
                            val = row[p]
                            new_line = f"{p}={val}||{parts[1]}||{parts[2]}||{parts[3]}||{parts[4]}"
                            new_ini_content.append(new_line)
                            is_param = True
                            break
                        except:
                            pass
                if not is_param:
                    new_ini_content.append(line)
        out_file = os.path.join(final_output_path, file_name)
        with open(out_file, 'w', encoding='utf-16') as f:
            f.writelines(new_ini_content)
    print(f"Tao xong 10 file INI tai: {bot_folder}/")
    return True

# ==================== MAIN ====================

def main(num_runs=1, mt5_path=None, terminal_data_path=None, filters=None):
    print(f"[INIT] Starting 0_run_optimized.py with terminal_data_path={terminal_data_path}", flush=True)
    sys.stdout.flush()
    
    if mt5_path:
        config.MT5_PATH = mt5_path
    if terminal_data_path:
        build_paths(terminal_data_path)
    else:
        build_paths(config.TERMINAL_DATA_PATH)

    print(f"\n{'='*60}", flush=True)
    print("SAVEBACKTEST - OPTIMIZED VERSION (RAM-based)", flush=True)
    print(f"{'='*60}", flush=True)
    print(f"MT5 Path: {config.MT5_PATH}", flush=True)
    print(f"Terminal Data: {config.TERMINAL_DATA_PATH}\n", flush=True)

    start_time = time.time()
    print("Don dep file cu...", flush=True)
    
    # Xóa HTML & PNG files cũ
    cleanup_html_png_files()
    
    if os.path.exists(config.OPTIMIZE_XML_DIR):
        for f in os.listdir(config.OPTIMIZE_XML_DIR):
            if f.endswith(('.csv', '.xml')):
                safe_remove_file(os.path.join(config.OPTIMIZE_XML_DIR, f))
    if os.path.exists(config.FILTERED_CSV_DIR):
        for f in os.listdir(config.FILTERED_CSV_DIR):
            safe_remove_file(os.path.join(config.FILTERED_CSV_DIR, f))
    if os.path.exists(config.BACKTEST_INI_DIR):
        for f in os.listdir(config.BACKTEST_INI_DIR):
            try:
                full_path = os.path.join(config.BACKTEST_INI_DIR, f)
                if os.path.isdir(full_path):
                    shutil.rmtree(full_path)
                else:
                    safe_remove_file(full_path)
            except Exception as e:
                print(f"[WARN] Failed to clean {f}: {e}", flush=True)
    print("Don dep xong!\n", flush=True)

    if not os.path.exists(config.INI_SOURCE_DIR):
        print(f"Khong thay thu muc INI: {config.INI_SOURCE_DIR}", flush=True)
        return

    ini_files = [f for f in os.listdir(config.INI_SOURCE_DIR) if f.endswith('.ini')]
    print(f"Tim thay {len(ini_files)} file INI\n", flush=True)

    summary = []
    total_input_rows = 0
    total_pass_rows = 0
    total_ini_files = 0

    for ini_file in ini_files:
        print(f"[PHASE1] Processing {ini_file}...", flush=True)
        merged_df = phase1_optimize(ini_file, num_runs)
        strategy_name = os.path.splitext(ini_file)[0]

        if merged_df is None or merged_df.empty:
            print(f"Bo qua {ini_file} (Khong co du lieu)", flush=True)
            summary.append({"strategy": strategy_name, "input_rows": 0, "passed_rows": 0, "ini_files": 0, "status": "NO_DATA"})
            continue

        input_rows = len(merged_df)
        total_input_rows += input_rows
        print(f"[PHASE2] Filtering {ini_file}...", flush=True)
        filtered_df = phase2_filter(merged_df, filters)
        passed_rows = 0 if filtered_df is None else len(filtered_df)
        total_pass_rows += passed_rows

        if filtered_df is None or filtered_df.empty:
            print(f"[RESULT] {strategy_name}: input={input_rows}, pass=0, INI created=0", flush=True)
            summary.append({"strategy": strategy_name, "input_rows": input_rows, "passed_rows": 0, "ini_files": 0, "status": "FILTERED_OUT"})
            del merged_df, filtered_df
            continue

        print(f"[PHASE3] Creating INI for {ini_file}...", flush=True)
        phase3_create_ini(filtered_df, ini_file)
        total_ini_files += passed_rows
        print(f"[RESULT] {strategy_name}: input={input_rows}, pass={passed_rows}, INI created={passed_rows}", flush=True)
        summary.append({"strategy": strategy_name, "input_rows": input_rows, "passed_rows": passed_rows, "ini_files": passed_rows, "status": "PASSED"})
        del merged_df, filtered_df

    print(f"\n{'='*60}", flush=True)
    print("TONG KET OPTIMIZATION", flush=True)
    print(f"Tong strategy xu ly : {len(ini_files)}", flush=True)
    print(f"Strategy dat filter : {sum(1 for item in summary if item['status'] == 'PASSED')}", flush=True)
    print(f"Strategy khong dat  : {sum(1 for item in summary if item['status'] != 'PASSED')}", flush=True)
    print(f"Tong bo da doc      : {total_input_rows}", flush=True)
    print(f"Tong bo dat filter  : {total_pass_rows}", flush=True)
    print(f"Tong file INI tao   : {total_ini_files}", flush=True)

    passed_items = [item for item in summary if item['status'] == 'PASSED']
    failed_items = [item for item in summary if item['status'] != 'PASSED']
    if passed_items:
        print("\nDAT:", flush=True)
        for item in passed_items:
            print(f"  [OK] {item['strategy']}: {item['passed_rows']} file INI", flush=True)
    if failed_items:
        print("\nKHONG DAT:", flush=True)
        for item in failed_items:
            print(f"  [FAIL] {item['strategy']}: {item['input_rows']} bo dau vao, 0 file INI ({item['status']})", flush=True)
    print(f"\nThu muc INI Every Tick: {config.BACKTEST_INI_DIR}", flush=True)
    print(f"{'='*60}\n", flush=True)
    elapsed = (time.time() - start_time) / 60
    print(f"\n{'='*60}", flush=True)
    print(f"HOAN THANH trong {elapsed:.2f} phut!", flush=True)
    print(f"Ket qua INI backtest san sang o: {config.BACKTEST_INI_DIR}")
    print(f"{'='*60}\n")
    
    # Cleanup HTML & PNG files tạo ra
    print("Dang don dep HTML/PNG files...", flush=True)
    cleanup_html_png_files()
    print("Done!\n", flush=True)

if __name__ == "__main__":
    print("Python optimize script started", flush=True)
    sys.stdout.flush()
    
    parser = argparse.ArgumentParser(description='SaveBacktest Optimizer')
    parser.add_argument('--num-runs', type=int, default=3, help='So lan optimize (1-20)')
    parser.add_argument('--mt5-path', type=str, default=None, help='Duong dan MT5 (tuy chon)')
    parser.add_argument('--terminal-path', type=str, default=None, help='Duong dan Terminal Data (tuy chon)')
    parser.add_argument('--json-output', action='store_true', help='Output ket qua duoi dang JSON')
    parser.add_argument('--min-profit', type=float, default=DEFAULT_FILTERS['min_profit'], help='Min Profit')
    parser.add_argument('--min-trades', type=float, default=DEFAULT_FILTERS['min_trades'], help='Min Trades')
    parser.add_argument('--max-trades', type=float, default=DEFAULT_FILTERS['max_trades'], help='Max Trades')
    parser.add_argument('--min-recovery-factor', type=float, default=DEFAULT_FILTERS['min_recovery_factor'], help='Min Recovery Factor')
    parser.add_argument('--min-profit-factor', type=float, default=DEFAULT_FILTERS['min_profit_factor'], help='Min Profit Factor')
    parser.add_argument('--max-profit-factor', type=float, default=DEFAULT_FILTERS['max_profit_factor'], help='Max Profit Factor')
    parser.add_argument('--min-equity-dd', type=float, default=DEFAULT_FILTERS['min_equity_dd'], help='Min Equity DD %')
    parser.add_argument('--max-equity-dd', type=float, default=DEFAULT_FILTERS['max_equity_dd'], help='Max Equity DD %')
    args = parser.parse_args()
    
    print(f"Args parsed: terminal_path={args.terminal_path}", flush=True)
    sys.stdout.flush()
    
    try:
        filters = {
            'min_profit': args.min_profit,
            'min_trades': args.min_trades,
            'max_trades': args.max_trades,
            'min_recovery_factor': args.min_recovery_factor,
            'min_profit_factor': args.min_profit_factor,
            'max_profit_factor': args.max_profit_factor,
            'min_equity_dd': args.min_equity_dd,
            'max_equity_dd': args.max_equity_dd,
        }
        result = main(
            num_runs=args.num_runs,
            mt5_path=args.mt5_path,
            terminal_data_path=args.terminal_path,
            filters=filters
        )
        if args.json_output:
            output = {
                "success": True,
                "status": "completed",
                "num_runs": args.num_runs,
                "backtest_ini_dir": config.BACKTEST_INI_DIR
            }
            print(json.dumps(output, ensure_ascii=False, indent=2), flush=True)
        print("[SUCCESS] Optimize completed", flush=True)
        sys.stdout.flush()
        sys.exit(0)
    except Exception as e:
        print(f"[ERROR] Exception: {e}", flush=True)
        import traceback
        traceback.print_exc(file=sys.stdout)
        sys.stdout.flush()
        if args.json_output:
            output = {
                "success": False,
                "status": "error",
                "error": str(e)
            }
            print(json.dumps(output, ensure_ascii=False, indent=2))
        else:
            print(f"Loi: {e}")
        sys.exit(1)
