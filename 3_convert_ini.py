#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CANH BAO: FILE NAY CHI LA LIBRARY - KHONG CHAY RIENG LE!

Hay dung:
    python 0_run_optimized.py

(Chạy file này sẽ tạo file INI backtest - gây lẫn lộn)
"""

# PRINT CẢNH BÁO NGAY KHI RUN
import sys
if __name__ == "__main__":
    print("KHONG CHAY FILE NAY RIENG LE!")
    print("Hay chay: python 0_run_optimized.py")
    print("File nay chi la library de file 0 import")
    sys.exit(1)

import pandas as pd
import os
import config 
import re

def get_bot_info(filename):
    """
    Bóc tách ID, Symbol và Timeframe từ tên file.
    Ví dụ: 'Filtered_122 - DE30H4_Buy' -> id='122', symbol='DE30', tf='H4'
    """
    clean_name = filename.replace('Filtered_', '').upper()
    
    # 1. Lấy ID (số đầu tiên trong tên file)
    id_match = re.search(r'(\d+)', clean_name)
    bot_id = id_match.group(1) if id_match else ""
    
    # 2. Lấy Timeframe (Tìm các cụm M15, M30, H1, H4, D1)
    tf_match = re.search(r'(M15|M30|H1|H4|D1)', clean_name)
    timeframe = tf_match.group(1) if tf_match else ""
    
    # 3. Lấy Symbol (Loại bỏ ID và TF để tìm Symbol còn lại)
    # Hoặc quét qua danh sách các Symbol phổ biến của bạn
    symbol = ""
    common_symbols = ["DE30", "HK50", "USOIL", "BTCUSD", "ETHUSD", "XAUUSD", "XAGUSD", "US500", "JP225"]
    for s in common_symbols:
        if s in clean_name:
            symbol = s
            break
            
    return bot_id, symbol, timeframe

def find_matching_ini(csv_name):
    """Tìm file INI gốc khớp cả 3 yếu tố: ID, Symbol và Timeframe"""
    csv_id, csv_symbol, csv_tf = get_bot_info(csv_name)
    is_buy = "BUY" in csv_name.upper()
    is_sell = "SELL" in csv_name.upper()
    
    if not os.path.exists(config.INI_SOURCE_DIR):
        return None
        
    all_inis = [f for f in os.listdir(config.INI_SOURCE_DIR) if f.endswith('.ini')]
    
    for ini in all_inis:
        ini_id, ini_symbol, ini_tf = get_bot_info(ini)
        
        # ĐIỀU KIỆN KHỚP NGHIÊM NGẶT: ID + SYMBOL + TIMEFRAME
        if ini_id == csv_id and ini_symbol == csv_symbol and ini_tf == csv_tf:
            # Kiểm tra thêm hướng Buy/Sell nếu có thể
            if (is_buy and "BUY" in ini.upper()) or (is_sell and "SELL" in ini.upper()):
                return os.path.join(config.INI_SOURCE_DIR, ini)
            # Trường hợp file INI gốc không ghi Buy/Sell thì vẫn chấp nhận vì đã khớp 3 yếu tố trên
            elif not ("BUY" in ini.upper() or "SELL" in ini.upper()):
                return os.path.join(config.INI_SOURCE_DIR, ini)
                
    return None

def main():
    if not os.path.exists(config.FILTERED_CSV_DIR):
        print(f"Thu muc CSV khong ton tai: {config.FILTERED_CSV_DIR}")
        return

    csv_files = [f for f in os.listdir(config.FILTERED_CSV_DIR) if f.endswith('.csv')]
    print(f"Dang xu ly {len(csv_files)} file CSV...")

    for csv_file in csv_files:
        template_path = find_matching_ini(csv_file)
        
        if not template_path:
            info = get_bot_info(csv_file)
            print(f"KHONG TIM THAY mau khop: ID={info[0]}, Symbol={info[1]}, TF={info[2]} cho {csv_file}")
            continue
        
        # Đọc nội dung file mẫu (Xử lý cả 2 loại mã hóa)
        source_lines = []
        try:
            with open(template_path, 'r', encoding='utf-16') as f:
                source_lines = f.readlines()
        except:
            with open(template_path, 'r', encoding='utf-8-sig') as f:
                source_lines = f.readlines()

        if not source_lines: continue

        # Đọc dữ liệu từ CSV đã lọc
        df = pd.read_csv(os.path.join(config.FILTERED_CSV_DIR, csv_file))
        
        # Tạo thư mục đầu ra cho Bot
        bot_folder = csv_file.replace('.csv', '').replace('Filtered_', '')
        final_output_path = os.path.join(config.BACKTEST_INI_DIR, bot_folder)
        if not os.path.exists(final_output_path): 
            os.makedirs(final_output_path)

        # Danh sách các cột không phải là thông số Input (Giữ nguyên của bạn)
        exclude = ['Pass', 'Result', 'Profit', 'Expected Payoff', 'Profit Factor', 
                   'Recovery Factor', 'Sharpe Ratio', 'Custom', 'Equity DD %', 
                   'Trades', 'DD_Deviation', 'Result_Priority',
                   # Cac cot moi do 2_filter_optimize.py (ban tich hop mt5_optimize_filter) tao ra
                   'bot', 'source_file', '_score', '_plateau_score', '_final_score',
                   '_pd_ratio', '_flag_reason']
        params = [col for col in df.columns if col not in exclude]

        for i, row in df.iterrows():
            new_ini_content = []
            # Đặt tên file theo đúng format bạn yêu cầu
            file_name = f"BT_Set_{i+1}_Profit_{int(row['Profit'])}.ini"
            
            for line in source_lines:
                # Ép cấu hình chạy Backtest chuẩn (Every Tick, No Optimize)
                if line.startswith("Optimization="):
                    new_ini_content.append("Optimization=0\n")
                elif line.startswith("Model="):
                    new_ini_content.append("Model=0\n") # 0 = Every Tick
                elif line.startswith("ShutdownTerminal="):
                    new_ini_content.append("ShutdownTerminal=1\n")
                elif line.startswith("Report=") or line.startswith("ReplaceReport="):
                    continue # Bỏ qua để File 4 tự điều khiển tên report
                else:
                    is_param = False
                    for p in params:
                        if line.startswith(f"{p}="):
                            try:
                                # Tách lấy phần đuôi ||Min||Step||Max||N để giữ nguyên cấu trúc MT5
                                parts = line.split('=')[1].split('||')
                                val = row[p]
                                # Ghi đè giá trị mới nhưng giữ nguyên format MT5
                                new_line = f"{p}={val}||{parts[1]}||{parts[2]}||{parts[3]}||{parts[4]}"
                                new_ini_content.append(new_line)
                                is_param = True
                                break
                            except: pass
                    
                    if not is_param:
                        new_ini_content.append(line)

            # Lưu file với định dạng UTF-16 LE (MT5 đọc tốt nhất)
            out_file = os.path.join(final_output_path, file_name)
            with open(out_file, 'w', encoding='utf-16') as f:
                f.writelines(new_ini_content)
                
    print(f"Da tao xong toan bo file .ini tai: {config.BACKTEST_INI_DIR}")

if __name__ == "__main__":
    main()