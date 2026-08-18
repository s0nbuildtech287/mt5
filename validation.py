import os
import sys
import re
import json
import time
import io
from glob import glob
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

# Tự động kiểm tra và cài đặt thư viện cần thiết nếu chưa có
try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
except ImportError:
    print("Đang cài đặt thư viện cần thiết (google-api-python-client, google-auth, openpyxl)...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "google-api-python-client", "google-auth", "openpyxl"])
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

SERVICE_ACCOUNT_DIR = r"C:\Users\SonBx\Desktop\Lotusquant\Optimize\backend\src\config\serviceAccounts"
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

VALID_SYMBOLS = [
    "BTCUSD", "ETHUSD", "XAUUSD", "XAGUSD", "US500", "DE30", "JP225", "HK50", 
    "USOIL", "STOXX50", "UK100", "US30", "USTECH", "NAS100", "UKOIL", "DE40",
    "EURUSD", "GBPUSD", "AUDUSD", "NZDUSD", "USDCAD", "USDCHF", "USDJPY"
]
VALID_TIMEFRAMES = ["H1", "H4"]
VALID_ORDER_TYPES = ["OB", "OS"]
MIN_TOTAL_FILES = 3600

class DriveAccountManager:
    """Quản lý và xoay vòng nhiều Service Account để tránh dính quota limit"""
    def __init__(self, sa_dir):
        self.sa_files = glob(os.path.join(sa_dir, "*.json"))
        if not self.sa_files:
            raise FileNotFoundError(f"Không tìm thấy file service account nào trong thư mục: {sa_dir}")
        self.current_idx = 0
        print(f"Đã tìm thấy {len(self.sa_files)} Service Accounts trong {sa_dir}")

    def get_service(self):
        sa_file = self.sa_files[self.current_idx]
        creds = service_account.Credentials.from_service_account_file(sa_file, scopes=SCOPES)
        return build('drive', 'v3', credentials=creds)

    def rotate_account(self):
        self.current_idx = (self.current_idx + 1) % len(self.sa_files)
        sa_name = os.path.basename(self.sa_files[self.current_idx])
        return self.get_service()


def extract_folder_id(url_or_id):
    """Trích xuất Folder ID từ đường dẫn Google Drive hoặc ID gốc"""
    url_or_id = url_or_id.strip()
    match = re.search(r'folders/([a-zA-Z0-9_-]+)', url_or_id)
    if match:
        return match.group(1)
    match_id = re.search(r'id=([a-zA-Z0-9_-]+)', url_or_id)
    if match_id:
        return match_id.group(1)
    return url_or_id


def list_items_in_folder(service_mgr, folder_id):
    """Lấy tất cả items (subfolders và files) trong 1 folder"""
    items = []
    page_token = None
    query = f"'{folder_id}' in parents and trashed = false"
    service = service_mgr.get_service()
    
    while True:
        try:
            response = service.files().list(
                q=query,
                spaces='drive',
                fields='nextPageToken, files(id, name, mimeType, size, modifiedTime)',
                pageToken=page_token,
                pageSize=1000
            ).execute()
            
            items.extend(response.get('files', []))
            page_token = response.get('nextPageToken')
            if not page_token:
                break
        except HttpError as error:
            if error.resp.status in [403, 429]:
                print(f"\n[QUOTA EXCEEDED] Đang xoay vòng Service Account...")
                service = service_mgr.rotate_account()
                time.sleep(1)
            else:
                print(f"\nLỗi API Google Drive: {error}")
                break
        except Exception as e:
            print(f"\nLỗi không xác định: {e}")
            break
            
    return items


def scan_asset_folder_recursive(service_mgr, folder_id, current_path=""):
    """Duyệt đệ quy tất cả các file trong folder tài sản"""
    all_files = []
    items = list_items_in_folder(service_mgr, folder_id)
    
    for item in items:
        mime_type = item.get('mimeType', '')
        item_name = item.get('name', '')
        
        if mime_type == 'application/vnd.google-apps.folder':
            sub_path = f"{current_path}/{item_name}" if current_path else item_name
            sub_files = scan_asset_folder_recursive(service_mgr, item.get('id'), sub_path)
            all_files.extend(sub_files)
        else:
            item['path'] = f"{current_path}/{item_name}" if current_path else item_name
            all_files.append(item)
            
    return all_files


# ==============================================================================
# 1. CHECK CẤU TRÚC FILE OPTIMIZATION (GROUP 1 & GROUP 2)
# ==============================================================================
def analyze_file_name(name):
    """
    Check Group 1:
    - Tên file phải viết HOA toàn bộ
    - Format chuẩn: {BotId}{Symbol}{Timeframe}{OrderType}
    - BotId đúng 3 chữ số
    - Symbol thuộc danh sách hợp lệ
    - Timeframe là H1 hoặc H4
    - OrderType là OB hoặc OS
    """
    base = re.sub(r'\.[^/.]+$', '', name)
    errors = []

    # Check hoa toàn bộ
    if base != base.upper():
        errors.append({"group": 1, "description": "Tên file phải viết hoa toàn bộ"})

    # Check BotId 3 chữ số
    bot_id_match = re.match(r'^(\d{3})', base)
    if not bot_id_match:
        errors.append({"group": 1, "description": "Thiếu hoặc sai định dạng BotId (phải là 3 chữ số)"})
        return {"valid": False, "errors": errors, "base": base}
    
    bot_id = bot_id_match.group(1)
    rest = base[3:]

    # Check Symbol
    symbol = None
    after_symbol = rest
    for sym in VALID_SYMBOLS:
        if rest.startswith(sym):
            symbol = sym
            after_symbol = rest[len(sym):]
            break
            
    if not symbol:
        guessed = re.sub(r'H1|H4|OB|OS', '', rest)[:8]
        errors.append({"group": 1, "description": f"Symbol không hợp lệ: {guessed or rest}"})
        return {"valid": False, "errors": errors, "base": base, "bot_id": bot_id}

    # Check Timeframe
    timeframe = None
    after_tf = after_symbol
    for tf in VALID_TIMEFRAMES:
        if after_symbol.startswith(tf):
            timeframe = tf
            after_tf = after_symbol[len(tf):]
            break

    if not timeframe:
        errors.append({"group": 1, "description": f"Timeframe không hợp lệ: {after_symbol[:2] or after_symbol}"})
        return {"valid": False, "errors": errors, "base": base, "bot_id": bot_id, "symbol": symbol}

    # Check OrderType
    order_type = None
    for ot in VALID_ORDER_TYPES:
        if after_tf == ot:
            order_type = ot
            break

    if not order_type:
        errors.append({"group": 1, "description": f"Thiếu hoặc sai OrderType (phải là OB hoặc OS): {after_tf}"})
        return {"valid": False, "errors": errors, "base": base, "bot_id": bot_id, "symbol": symbol, "timeframe": timeframe}

    return {
        "valid": len(errors) == 0,
        "errors": errors,
        "base": base,
        "bot_id": bot_id,
        "symbol": symbol,
        "timeframe": timeframe,
        "order_type": order_type
    }


def validate_drive_structure(file_list):
    """
    Check Group 2:
    - Tối thiểu 3600 file cho toàn bộ folder drive
    - Không có 2 file trùng tên trong cùng folder
    - Mỗi cặp BotId+Symbol phải có đủ H1OB, H1OS, H4OB, H4OS
    - Phát hiện file nào tên lạ ko đúng theo format tên
    """
    results = []

    # 1. Check tổng số file
    if len(file_list) < MIN_TOTAL_FILES:
        results.append({
            "file_name": "[Tổng số file]",
            "group": 2,
            "description": f"Tổng số file không đủ: {len(file_list)}/{MIN_TOTAL_FILES} (thiếu {MIN_TOTAL_FILES - len(file_list)} file)",
            "status": "LỖI"
        })

    # 2. Check trùng tên file
    name_counts = {}
    for item in file_list:
        n = item['name']
        name_counts[n] = name_counts.get(n, 0) + 1

    for name, count in name_counts.items():
        if count > 1:
            results.append({
                "file_name": name,
                "group": 2,
                "description": f"Tên file bị trùng trong folder Drive (xuất hiện {count} lần)",
                "status": "LỖI"
            })

    valid_by_pair = {}
    combo_counts = {}

    # 3. Phân tích từng file
    for item in file_list:
        fname = item['name']
        if name_counts.get(fname, 0) > 1:
            continue  # Đã flag trùng tên ở trên

        analyzed = analyze_file_name(fname)
        if not analyzed['valid']:
            for err in analyzed['errors']:
                results.append({
                    "file_name": fname,
                    "group": err['group'],
                    "description": err['description'],
                    "status": "LỖI"
                })
            continue

        bot_id = analyzed['bot_id']
        symbol = analyzed['symbol']
        tf = analyzed['timeframe']
        ot = analyzed['order_type']

        pair_key = f"{bot_id}{symbol}"
        combo_key = f"{bot_id}{symbol}{tf}{ot}"

        if pair_key not in valid_by_pair:
            valid_by_pair[pair_key] = []
        valid_by_pair[pair_key].append({"file_name": fname, "timeframe": tf, "order_type": ot})

        if combo_key not in combo_counts:
            combo_counts[combo_key] = []
        combo_counts[combo_key].append(fname)

    # Check trùng combo
    for combo_key, names in combo_counts.items():
        if len(names) > 1:
            results.append({
                "file_name": f"[Trùng Combo] {combo_key}",
                "group": 2,
                "description": f"Trùng file combo: {', '.join(names)}",
                "status": "LỖI"
            })

    # Check đủ bộ 4 file: H1OB, H1OS, H4OB, H4OS cho từng cặp BotId+Symbol
    for pair_key, files in valid_by_pair.items():
        required = ["H1OB", "H1OS", "H4OB", "H4OS"]
        present = [f"{f['timeframe']}{f['order_type']}" for f in files]
        missing = [r for r in required if r not in present]

        if missing:
            results.append({
                "file_name": f"[Bộ {pair_key}]",
                "group": 2,
                "description": f"Thiếu file: {', '.join(missing)}",
                "status": "LỖI"
            })

    return results


# ==============================================================================
# 2. CHECK DỮ LIỆU TRONG FILE EXCEL (GROUP 3 & GROUP 4)
# ==============================================================================
def parse_excel_content(file_bytes):
    """Phân tích nội dung file Excel thành Settings, Results, và Deals"""
    wb = openpyxl.load_workbook(filename=io.BytesIO(file_bytes), data_only=True)
    sheet = wb.active

    rows = list(sheet.iter_rows(values_only=True))

    settings = {}
    results = {}
    deals = []

    current_section = None
    header_deals = None

    for idx, row in enumerate(rows):
        if not row or not any(row):
            continue

        first_val = str(row[0]).strip() if row[0] is not None else ""

        if first_val.lower() == "settings":
            current_section = "settings"
            continue
        elif first_val.lower() == "results":
            current_section = "results"
            continue
        elif first_val.lower() == "deals":
            current_section = "deals"
            if idx + 1 < len(rows):
                header_deals = [str(c).strip().lower() if c is not None else "" for c in rows[idx + 1]]
            continue

        if current_section == "settings":
            if len(row) >= 2 and row[0] is not None:
                key = str(row[0]).strip()
                val = row[1]
                settings[key] = val
        elif current_section == "results":
            if len(row) >= 2 and row[0] is not None:
                key = str(row[0]).strip()
                val = row[1]
                results[key] = val
        elif current_section == "deals" and header_deals:
            if first_val.lower() in ["results", "summary", "settings", "orders", "graph"]:
                break
            if row == rows[rows.index(row)] and any(h in first_val.lower() for h in ["time", "deal", "type"]):
                continue  # Header row

            deal_dict = {}
            for col_i, col_name in enumerate(header_deals):
                if col_i < len(row):
                    deal_dict[col_name] = row[col_i]
            if deal_dict.get("time") and str(deal_dict.get("type", "")).lower() in ["buy", "sell"]:
                deals.append(deal_dict)

    return settings, results, deals


def check_file_content_and_deals(file_name, file_bytes, min_profit=0, min_trades=100, max_loss=-20000):
    """Check Group 3 & Group 4 cho 1 file Excel"""
    errors = []
    analyzed = analyze_file_name(file_name)

    try:
        settings, results, deals = parse_excel_content(file_bytes)
    except Exception as e:
        errors.append({"file_name": file_name, "group": 3, "description": f"Không đọc được file Excel: {e}", "status": "LỖI"})
        return errors

    # --- GROUP 3: THÔNG TIN CƠ BẢN ---
    if analyzed['valid']:
        content_symbol = str(settings.get("Symbol", "")).strip().upper()
        if content_symbol and content_symbol != analyzed['symbol']:
            errors.append({
                "file_name": file_name, "group": 3,
                "description": f"Symbol mâu thuẫn: tên file={analyzed['symbol']}, nội dung={content_symbol}",
                "status": "LỖI"
            })

        content_tf = str(settings.get("Period", "")).strip().upper()
        if content_tf and content_tf != analyzed['timeframe']:
            errors.append({
                "file_name": file_name, "group": 3,
                "description": f"Timeframe mâu thuẫn: tên file={analyzed['timeframe']}, nội dung={content_tf}",
                "status": "LỖI"
            })

        expert_str = str(settings.get("Expert", "")).strip()
        expert_match = re.match(r'^(\d{3})', expert_str)
        if expert_match and expert_match.group(1) != analyzed['bot_id']:
            errors.append({
                "file_name": file_name, "group": 3,
                "description": f"BotId mâu thuẫn: tên file={analyzed['bot_id']}, Expert ID={expert_match.group(1)}",
                "status": "LỖI"
            })

    def safe_float(val):
        try: return float(val)
        except: return None

    net_profit = safe_float(results.get("Total Net Profit"))
    total_trades = safe_float(results.get("Total Trades"))
    largest_loss = safe_float(results.get("Largest loss trade"))
    profit_factor = safe_float(results.get("Profit Factor"))

    if net_profit is not None and net_profit < min_profit:
        errors.append({"file_name": file_name, "group": 3, "description": f"Total Net Profit thấp: {net_profit} (yêu cầu ≥ {min_profit})", "status": "LỖI"})
    if total_trades is not None and total_trades < min_trades:
        errors.append({"file_name": file_name, "group": 3, "description": f"Total Trades thấp: {total_trades} (yêu cầu ≥ {min_trades})", "status": "LỖI"})
    if largest_loss is not None and largest_loss < max_loss:
        errors.append({"file_name": file_name, "group": 3, "description": f"Largest Loss Trade vượt ngưỡng: {largest_loss} (yêu cầu ≥ {max_loss})", "status": "LỖI"})
    if profit_factor is not None and profit_factor <= 1.0:
        errors.append({"file_name": file_name, "group": 3, "description": f"Profit Factor ≤ 1.0: {profit_factor} (yêu cầu > 1.0)", "status": "LỖI"})

    # --- GROUP 4: LOGIC GIAO DỊCH (DEALS) ---
    if deals:
        active_opens = []
        last_deal = None

        for idx, deal in enumerate(deals):
            d_time = str(deal.get("time", "")).strip()
            d_type = str(deal.get("type", "")).strip().lower()
            d_direction = str(deal.get("direction", deal.get("entry", ""))).strip().lower()
            d_comment = str(deal.get("comment", "")).strip()
            d_swap = safe_float(deal.get("swap"))
            d_vol = safe_float(deal.get("volume"))

            if d_swap is not None and d_swap != 0:
                errors.append({"file_name": file_name, "group": 4, "description": f"Swap ≠ 0: {d_swap} (deal #{idx+1}, {d_time})", "status": "LỖI"})

            if d_vol is not None and d_vol <= 0:
                errors.append({"file_name": file_name, "group": 4, "description": f"Volume ≤ 0: {d_vol} (deal #{idx+1}, {d_time})", "status": "LỖI"})

            if not d_comment:
                errors.append({"file_name": file_name, "group": 3, "description": f"Comment trống (deal #{idx+1}, {d_time})", "status": "LỖI"})

            if last_deal and last_deal.get("type") == d_type and d_direction == "in" and last_deal.get("direction") == "in":
                if not (d_comment and last_deal.get("comment") and d_comment.split("_")[0] == last_deal.get("comment").split("_")[0]):
                    errors.append({
                        "file_name": file_name, "group": 4,
                        "description": f"Cặp lệnh {d_type.upper()}/{d_type.upper()} liền nhau không qua đóng lệnh (deal #{idx} → #{idx+1})",
                        "status": "LỖI"
                    })

            if d_direction == "in":
                if active_opens:
                    errors.append({"file_name": file_name, "group": 4, "description": f"Mở lệnh mới khi chưa đóng lệnh cũ (deal #{idx+1})", "status": "LỖI"})
                active_opens.append({"deal_idx": idx+1, "time": d_time, "vol": d_vol, "comment": d_comment})

            elif d_direction == "out":
                if active_opens:
                    matched = active_opens.pop(0)
                    if matched['time'] == d_time and not (d_comment.lower().startswith("sl ") or d_comment.lower().startswith("tp ")):
                        errors.append({"file_name": file_name, "group": 4, "description": f"Đóng lệnh ngay sau khi mở cùng timestamp: {d_time}", "status": "LỖI"})
                    
                    try:
                        t1 = datetime.strptime(matched['time'].replace(".", "-"), "%Y-%m-%dT%H:%M:%S" if "T" in matched['time'] else "%Y-%m-%d %H:%M:%S")
                        t2 = datetime.strptime(d_time.replace(".", "-"), "%Y-%m-%dT%H:%M:%S" if "T" in d_time else "%Y-%m-%d %H:%M:%S")
                        diff_sec = (t2 - t1).total_seconds()
                        if 0 < diff_sec < 60 and ("_C" in d_comment or "_CO" in d_comment):
                            errors.append({"file_name": file_name, "group": 4, "description": f"Giữ lệnh < 1 phút ({diff_sec}s) và đóng bằng _C/_CO", "status": "LỖI"})
                    except:
                        pass

            last_deal = deal

    return errors


def fetch_and_check_single_file(service_mgr, item):
    file_id = item['id']
    file_name = item['name']
    
    for attempt in range(5):
        try:
            service = service_mgr.get_service()
            try:
                res = service.files().export(
                    fileId=file_id,
                    mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                ).execute()
                file_bytes = res
            except Exception:
                res = service.files().get_media(fileId=file_id).execute()
                file_bytes = res

            return check_file_content_and_deals(file_name, file_bytes)
        except Exception as e:
            if any(k in str(e).lower() for k in ["429", "403", "quota", "rate limit"]):
                service_mgr.rotate_account()
                time.sleep(0.5)
            else:
                break
    return [{"file_name": file_name, "group": 3, "description": "Không thể tải/đọc nội dung file từ Drive", "status": "LỖI"}]


# ==============================================================================
# 3. XUẤT FILE EXCEL BÁO CÁO KẾT QUẢ
# ==============================================================================
def export_validation_report_to_excel(output_filename, report_data, all_errors, all_files):
    """Xuất toàn bộ dữ liệu thống kê và danh sách lỗi ra file Excel định dạng đẹp"""
    wb = openpyxl.Workbook()

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    err_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    err_font = Font(name="Calibri", size=11, color="C00000", bold=True)
    align_center = Alignment(horizontal="center", vertical="center")

    # Sheet 1: Thống kê số lượng file theo folder tài sản
    ws1 = wb.active
    ws1.title = "Thong_Ke_Tai_San"
    ws1.append(["STT", "Tên Folder Tài Sản", "Tổng Số File", "Số File .xlsx"])
    for cell in ws1[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center

    for idx, row in enumerate(report_data, 1):
        ws1.append([idx, row['folder'], row['total_files'], row['xlsx_files']])
        ws1.cell(row=idx+1, column=1).alignment = align_center
        ws1.cell(row=idx+1, column=3).alignment = align_center
        ws1.cell(row=idx+1, column=4).alignment = align_center

    # Sheet 2: Danh sách lỗi Validation chi tiết
    ws2 = wb.create_sheet(title="Danh_Sach_Loi_Validation")
    ws2.append(["STT", "Trạng Thái", "Nhóm Lỗi", "File / Mục Bị Lỗi", "Mô Tả Chi Tiết Lỗi"])
    for cell in ws2[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center

    if not all_errors:
        ws2.append([1, "OK", "-", "-", "Toàn bộ cấu trúc folder, tên file và nội dung dữ liệu HOÀN TOÀN HỢP LỆ!"])
    else:
        for idx, err in enumerate(all_errors, 1):
            ws2.append([idx, err.get('status', 'LỖI'), f"Group {err.get('group', 2)}", err.get('file_name', ''), err.get('description', '')])
            row_idx = idx + 1
            ws2.cell(row=row_idx, column=1).alignment = align_center
            ws2.cell(row=row_idx, column=2).alignment = align_center
            ws2.cell(row=row_idx, column=2).font = err_font
            ws2.cell(row=row_idx, column=2).fill = err_fill
            ws2.cell(row=row_idx, column=3).alignment = align_center

    # Sheet 3: Danh sách toàn bộ file đã quét
    ws3 = wb.create_sheet(title="Danh_Sach_Toan_Bo_File")
    ws3.append(["STT", "Tên File", "Đường Dẫn Trong Drive", "File ID"])
    for cell in ws3[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center

    for idx, f in enumerate(all_files, 1):
        ws3.append([idx, f.get('name', ''), f.get('path', ''), f.get('id', '')])
        ws3.cell(row=idx+1, column=1).alignment = align_center

    # Tự động căn chỉnh độ rộng cột
    for ws in [ws1, ws2, ws3]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

    wb.save(output_filename)
    print(f"\n[EXCEL EXPORT] Đã tự động xuất báo cáo đầy đủ ({len(all_errors)} lỗi) ra file Excel:")
    print(f" 📂 File Excel: {os.path.abspath(output_filename)}")


# ==============================================================================
# 4. QUẢN LÝ QUÉT & BÁO CÁO
# ==============================================================================
def process_drive_validation(url_or_id, service_mgr):
    folder_id = extract_folder_id(url_or_id)
    
    print("\n" + "="*85)
    print(f" 🚀 QUÉT VÀ KIỂM TRA ĐIỀU KIỆN VALIDATION LINK DRIVE (ID: {folder_id})")
    print("="*85)

    print("Step 1: Đang tải danh sách toàn bộ thư mục và file...")
    top_items = list_items_in_folder(service_mgr, folder_id)
    asset_folders = [item for item in top_items if item.get('mimeType') == 'application/vnd.google-apps.folder']
    
    all_files = []
    report_data = []

    for idx, asset in enumerate(asset_folders, 1):
        asset_id = asset.get('id')
        asset_name = asset.get('name')
        print(f"  [{idx}/{len(asset_folders)}] Đang thu thập file trong folder: {asset_name} ...", end="", flush=True)
        files_in_asset = scan_asset_folder_recursive(service_mgr, asset_id)
        all_files.extend(files_in_asset)
        
        xlsx_count = sum(1 for f in files_in_asset if f.get('name', '').endswith('.xlsx'))
        report_data.append({
            "folder": asset_name,
            "total_files": len(files_in_asset),
            "xlsx_files": xlsx_count
        })
        print(f" -> Done ({len(files_in_asset)} files)")

    print(f"\nStep 2: Thực hiện kiểm tra Validation Cấu Trúc (Group 1 & Group 2) cho {len(all_files)} file...")
    struct_errors = validate_drive_structure(all_files)

    print("\n" + "="*85)
    print(" 📊 BẢNG THỐNG KÊ SỐ LƯỢNG FILE THEO TÀI SẢN")
    print("="*85)
    print(f"{'STT':<5} | {'Tên Folder Tài Sản':<25} | {'Tổng số File':<15} | {'Số file .xlsx':<15}")
    print("-" * 68)
    for i, row in enumerate(report_data, 1):
        print(f"{i:<5} | {row['folder']:<25} | {row['total_files']:<15} | {row['xlsx_files']:<15}")
    print("-" * 68)
    print(f"TỔNG CỘNG: {len(asset_folders)} folder tài sản | {len(all_files)} tổng file")
    print("="*85)

    print(f"\nStep 3: Thực hiện kiểm tra Nội dung File Excel (Group 3 & Group 4) cho {len(all_files)} file...")
    content_errors = []
    
    user_choice = input("Bạn có muốn tải & kiểm tra chi tiết nội dung dữ liệu bên trong file Excel (Group 3 & Group 4) không? (y/n, mặc định y): ").strip().lower()
    if user_choice in ["", "y", "yes"]:
        print("  Đang chạy kiểm tra đa luồng (multi-thread 12 workers) bằng 8 Service Accounts...")
        completed = 0
        total = len(all_files)
        with ThreadPoolExecutor(max_workers=12) as executor:
            future_map = {executor.submit(fetch_and_check_single_file, service_mgr, f): f for f in all_files}
            for future in as_completed(future_map):
                completed += 1
                if completed % 100 == 0 or completed == total:
                    print(f"  [TIẾN ĐỘ] Đã kiểm tra nội dung {completed}/{total} file...", flush=True)
                res = future.result()
                if res:
                    for err in res:
                        if err.get('status') == 'LỖI':
                            content_errors.append(err)

    all_errors = struct_errors + content_errors

    print("\n" + "="*85)
    print(f" 📋 TỔNG KẾT VALIDATION (Group 1-4): {len(all_errors)} LỖI BỊ PHÁT HIỆN")
    print("="*85)
    if not all_errors:
        print("  🎉 XIN CHÚC MỪNG! Toàn bộ file và dữ liệu HOÀN TOÀN HỢP LỆ!")
    else:
        for err in all_errors[:20]:
            print(f"  ❌ [{err['status']}] [Group {err['group']}] File/Mục: {err['file_name']} -> {err['description']}")
        if len(all_errors) > 20:
            print(f"  ... và còn {len(all_errors) - 20} lỗi khác (Xem chi tiết trong file Excel export).")
    print("="*85)

    # Tự động xuất ra file Excel ngay trong thư mục chứa file script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_filename = os.path.join(script_dir, f"Bao_Cao_Validation_Drive_{timestamp_str}.xlsx")
    export_validation_report_to_excel(excel_filename, report_data, all_errors, all_files)

    return {
        "total_files": len(all_files),
        "all_errors": all_errors
    }


def main():
    print("==========================================================================")
    print("  TOOL KIỂM TRA VALIDATION FILE OPTIMIZATION & GOOGLE DRIVE  ")
    print("==========================================================================")

    try:
        service_mgr = DriveAccountManager(SERVICE_ACCOUNT_DIR)
    except Exception as e:
        print(f"Lỗi khởi tạo Service Account: {e}")
        return

    if len(sys.argv) > 1:
        drive_links = sys.argv[1:]
    else:
        user_input = input("\nNhập vào Link hoặc Folder ID Google Drive:\n> ")
        drive_links = [link.strip() for link in user_input.split(",") if link.strip()]

    if not drive_links:
        print("Không có link Google Drive nào được nhập.")
        return

    for link in drive_links:
        try:
            process_drive_validation(link, service_mgr)
        except Exception as e:
            print(f"Lỗi trong quá trình kiểm tra link '{link}': {e}")

    print("\nDONE! Đã hoàn thành toàn bộ kiểm tra Validation.")


if __name__ == "__main__":
    main()
