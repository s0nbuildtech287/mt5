import os

# --- THONG TIN CO DINH (Sua theo may cua ban) ---
MT5_PATH = r"C:\Program Files\MetaTrader 5\terminal64.exe"
DEFAULT_TERMINAL_DATA_PATH = r"C:\Users\XUAN SON\AppData\Roaming\MetaQuotes\Terminal\D0E8209F77C8CF37AD8BF550E51FF075"


def init_paths(terminal_data_path=None):
    terminal_data_path = terminal_data_path or DEFAULT_TERMINAL_DATA_PATH

    ini_source_dir = os.path.join(terminal_data_path, r"MQL5\Files\Files_ini")
    optimize_xml_dir = os.path.join(terminal_data_path, r"MQL5\Files\Optimize_xml")
    filtered_csv_dir = os.path.join(terminal_data_path, r"MQL5\Files\Optimize_filter")
    backtest_ini_dir = os.path.join(terminal_data_path, r"MQL5\Files\Backtest_ini")
    base_result_dir = os.path.join(terminal_data_path, r"MQL5\Files\Report_backtest")
    raw_report_dir = os.path.join(base_result_dir, "raw")

    for d in [ini_source_dir, optimize_xml_dir, filtered_csv_dir, backtest_ini_dir, base_result_dir, raw_report_dir]:
        if not os.path.exists(d):
            os.makedirs(d)

    return {
        "TERMINAL_DATA_PATH": terminal_data_path,
        "INI_SOURCE_DIR": ini_source_dir,
        "OPTIMIZE_XML_DIR": optimize_xml_dir,
        "FILTERED_CSV_DIR": filtered_csv_dir,
        "BACKTEST_INI_DIR": backtest_ini_dir,
        "BASE_RESULT_DIR": base_result_dir,
        "RAW_REPORT_DIR": raw_report_dir,
    }


# Backward-compatible defaults
_paths = init_paths(DEFAULT_TERMINAL_DATA_PATH)
TERMINAL_DATA_PATH = _paths["TERMINAL_DATA_PATH"]
INI_SOURCE_DIR = _paths["INI_SOURCE_DIR"]
OPTIMIZE_XML_DIR = _paths["OPTIMIZE_XML_DIR"]
FILTERED_CSV_DIR = _paths["FILTERED_CSV_DIR"]
BACKTEST_INI_DIR = _paths["BACKTEST_INI_DIR"]
BASE_RESULT_DIR = _paths["BASE_RESULT_DIR"]
RAW_REPORT_DIR = _paths["RAW_REPORT_DIR"]


def get_report_name_from_ini(ini_file_path):
    """
    Lay ten file report tu file INI
    Vi du: 001 - Strategy Supertrend HeikinashiXAGUSDH1_Buy.ini
         -> 001 - Strategy Supertrend HeikinashiXAGUSDH1_Buy.xml
    """
    try:
        encodings = ['utf-16', 'utf-8-sig', 'utf-8']
        lines = []
        for enc in encodings:
            try:
                with open(ini_file_path, 'r', encoding=enc) as f:
                    lines = f.readlines()
                break
            except:
                continue
        for line in lines:
            if line.startswith('Report='):
                report_name = line.split('=')[1].strip()
                return report_name if report_name else None
        ini_name = os.path.basename(ini_file_path)
        return ini_name.replace('.ini', '.xml')
    except Exception as e:
        print(f"Loi khi doc Report name tu INI: {e}")
        ini_name = os.path.basename(ini_file_path)
        return ini_name.replace('.ini', '.xml')

def kill_mt5(pid: int = None):
    if pid:
        # Kill specific PID and its child processes with /t flag
        os.system(f"taskkill /f /t /pid {pid} >nul 2>&1")
    else:
        # Kill all terminal64.exe instances (fallback for cleanup)
        os.system("taskkill /f /im terminal64.exe >nul 2>&1")