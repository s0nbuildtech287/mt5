#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CANH BAO: FILE NAY CHI LA LIBRARY - KHONG CHAY RIENG LE!

Hay dung:
    python 0_run_optimized.py

(Chạy file này sẽ tạo file CSV - gây lẫn lộn)
"""

# PRINT CẢNH BÁO NGAY KHI IMPORT
import sys
if __name__ == "__main__":
    print("KHONG CHAY FILE NAY RIENG LE!")
    print("Hay chay: python 0_run_optimized.py")
    print("File nay chi la library de file 0 import")
    sys.exit(1)

import os
import subprocess
import time
import shutil
import config

def fix_ini_before_run(file_path, report_name):
    """Đảm bảo file INI có đủ cấu hình xuất Report trước khi chạy"""
    lines = []
    encodings = ['utf-16', 'utf-8-sig', 'utf-8']
    for enc in encodings:
        try:
            with open(file_path, 'r', encoding=enc) as f:
                lines = f.readlines()
            break
        except: continue
    
    if not lines: return

    # Kiểm tra xem đã có các dòng cần thiết chưa
    has_report = any("Report=" in l for l in lines)
    has_replace = any("ReplaceReport=" in l for l in lines)
    has_shutdown = any("ShutdownTerminal=" in l for l in lines)

    if not (has_report and has_replace and has_shutdown):
        new_lines = []
        for line in lines:
            new_lines.append(line)
            if line.strip().lower() == "[tester]":
                if not has_report: new_lines.append(f"Report={report_name}\n")
                if not has_replace: new_lines.append("ReplaceReport=1\n")
                if not has_shutdown: new_lines.append("ShutdownTerminal=1\n")
        
        with open(file_path, 'w', encoding='utf-16') as f:
            f.writelines(new_lines)
        print(f"Da cap nhat cau hinh cho: {os.path.basename(file_path)}")

def get_latest_xml(report_name):
    """Tìm file XML kết quả theo tên report"""
    for root, dirs, files in os.walk(config.TERMINAL_DATA_PATH):
        for f in files:
            if f.lower() == report_name.lower():
                return os.path.join(root, f)
    return None

def main():
    if not os.path.exists(config.INI_SOURCE_DIR):
        print(f"Khong thay thu muc INI: {config.INI_SOURCE_DIR}")
        return

    ini_files = [f for f in os.listdir(config.INI_SOURCE_DIR) if f.endswith('.ini')]
    
    for ini in ini_files:
        path_ini = os.path.join(config.INI_SOURCE_DIR, ini)
        
        # Lấy tên file report từ file INI
        report_name = config.get_report_name_from_ini(path_ini)
        print(f"Report name se la: {report_name}", flush=True)
        
        # BƯỚC QUAN TRỌNG: Sửa file trước khi chạy
        fix_ini_before_run(path_ini, report_name)
        
        # Kill any leftover MT5 processes (all instances)
        config.kill_mt5()
        print(f"Optimizing: {ini}...", flush=True)
        
        # Chạy MT5 với file INI đã được chuẩn hóa - CAPTURE PID
        proc = subprocess.Popen(f'"{config.MT5_PATH}" /config:"{path_ini}"', shell=True)
        mt5_pid = proc.pid
        print(f"[1_save_optimize] Spawned MT5 with PID={mt5_pid}", flush=True)
        
        found = False
        start = time.time()
        # Timeout 1 tiếng cho mỗi file Optimize
        while (time.time() - start) < 3600: 
            latest = get_latest_xml(report_name)
            if latest and os.path.getsize(latest) > 0:
                time.sleep(2) # Đợi file ghi xong hẳn
                dst_path = os.path.join(config.OPTIMIZE_XML_DIR, report_name)
                shutil.copy2(latest, dst_path)
                os.remove(latest)
                found = True
                print(f"Da luu XML thanh cong: {os.path.basename(dst_path)}", flush=True)
                break
            time.sleep(10)
            
        if not found:
            print(f"Qua thoi gian cho (Timeout) cho file: {ini}", flush=True)
        
        # Kill ONLY this MT5 process by PID - not all terminal64.exe
        print(f"[1_save_optimize] Killing MT5 PID={mt5_pid}", flush=True)
        config.kill_mt5(pid=mt5_pid)
        time.sleep(3)

if __name__ == "__main__":
    main()