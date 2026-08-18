import os
import sys
import re
import time
from glob import glob
from datetime import datetime

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
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

SERVICE_ACCOUNT_DIR = r"C:\Users\XUAN SON\Desktop\Xuan Son Version\Lotusquant\Optimazation Process Data\backend\src\config\serviceAccounts"
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

VALID_SYMBOLS = [
    "BTCUSD", "ETHUSD", "XAUUSD", "XAGUSD", "US500", "DE30", "JP225", "HK50", 
    "USOIL", "STOXX50", "UK100", "US30", "USTECH", "NAS100", "UKOIL", "DE40",
    "EURUSD", "GBPUSD", "AUDUSD", "NZDUSD", "USDCAD", "USDCHF", "USDJPY"
]
VALID_TIMEFRAMES = ["H1", "H4"]
VALID_ORDER_TYPES = ["OB", "OS"]


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


def scan_asset_folder(service_mgr, folder_id, current_path=""):
    """
    Duyệt đệ quy các item trong folder tài sản.
    """
    all_files = []
    unexpected_items = []
    
    items = list_items_in_folder(service_mgr, folder_id)
    
    for item in items:
        mime_type = item.get('mimeType', '')
        item_name = item.get('name', '')
        item_id = item.get('id', '')
        item_path = f"{current_path}/{item_name}" if current_path else item_name
        
        if mime_type == 'application/vnd.google-apps.folder':
            unexpected_items.append({
                "id": item_id,
                "name": item_name,
                "type": "Folder lạ trong Folder tài sản",
                "reason": "Phát hiện thư mục con lạ nằm bên trong folder tài sản",
                "path": item_path
            })
            sub_files, sub_unexpected = scan_asset_folder(service_mgr, item_id, item_path)
            all_files.extend(sub_files)
            unexpected_items.extend(sub_unexpected)
        else:
            item['path'] = item_path
            all_files.append(item)
            
            EXCEL_MIME_TYPES = {
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "application/vnd.ms-excel",
                "application/vnd.google-apps.spreadsheet"
            }
            is_excel = item_name.lower().endswith(('.xlsx', '.xls')) or mime_type in EXCEL_MIME_TYPES
            if not is_excel:
                unexpected_items.append({
                    "id": item_id,
                    "name": item_name,
                    "type": "File lạ (Không phải file Excel)",
                    "reason": f"File không phải định dạng Excel/Google Sheet (MIME: {mime_type})",
                    "path": item_path
                })
            
    return all_files, unexpected_items


def check_filename_format(name):
    """
    Kiểm tra tên file đúng chuẩn định dạng: {BotId}{Symbol}{Timeframe}{OrderType}
    Ví dụ hợp lệ: 001BTCUSDH1OB (.xlsx)
    """
    base = re.sub(r'\.[^/.]+$', '', name)
    
    if base != base.upper():
        return False, "Tên file phải viết HOA toàn bộ", None, None

    bot_id_match = re.match(r'^(\d{3})', base)
    if not bot_id_match:
        return False, "BotId không hợp lệ (phải là 3 chữ số đầu tiên)", None, None
    
    bot_id = bot_id_match.group(1)
    rest = base[3:]

    symbol = None
    after_symbol = rest
    for sym in VALID_SYMBOLS:
        if rest.startswith(sym):
            symbol = sym
            after_symbol = rest[len(sym):]
            break
            
    if not symbol:
        return False, "Symbol không nằm trong danh sách hợp lệ", bot_id, None

    timeframe = None
    after_tf = after_symbol
    for tf in VALID_TIMEFRAMES:
        if after_symbol.startswith(tf):
            timeframe = tf
            after_tf = after_symbol[len(tf):]
            break

    if not timeframe:
        return False, "Khung thời gian không hợp lệ (phải là H1 hoặc H4)", bot_id, symbol

    order_type = None
    for ot in VALID_ORDER_TYPES:
        if after_tf == ot:
            order_type = ot
            break

    if not order_type:
        return False, "Kiểu lệnh không hợp lệ (phải là OB hoặc OS)", bot_id, symbol

    return True, "Hợp lệ", bot_id, symbol


def export_count_report_to_excel(output_filename, report_data, bot_summary, unexpected_items, all_files, is_detailed_checked):
    """Xuất thống kê ra Excel dựa trên kết quả các bước đã chạy"""
    wb = openpyxl.Workbook()

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    err_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    err_font = Font(name="Calibri", size=11, color="C00000", bold=True)
    align_center = Alignment(horizontal="center", vertical="center")

    # Sheet 1: Thống kê số lượng file theo folder tài sản
    ws1 = wb.active
    ws1.title = "Thong_Ke_Tai_San"
    if is_detailed_checked:
        ws1.append(["STT", "Tên Folder Tài Sản", "Tổng Số File", "Số File .xlsx", "Tên Hợp Lệ", "Tên SAI Định Dạng", "Số Mục Lạ"])
    else:
        ws1.append(["STT", "Tên Folder Tài Sản", "Tổng Số File", "Số File .xlsx"])
    
    for cell in ws1[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center

    for idx, row in enumerate(report_data, 1):
        if is_detailed_checked:
            ws1.append([idx, row['folder'], row['total_files'], row['xlsx_files'], row.get('valid_name_files', 0), row.get('invalid_name_files', 0), row.get('unexpected_count', 0)])
            ws1.cell(row=idx+1, column=5).alignment = align_center
            ws1.cell(row=idx+1, column=6).alignment = align_center
            ws1.cell(row=idx+1, column=7).alignment = align_center
            if row.get('invalid_name_files', 0) > 0 or row.get('unexpected_count', 0) > 0:
                ws1.cell(row=idx+1, column=6).font = err_font
                ws1.cell(row=idx+1, column=7).font = err_font
                ws1.cell(row=idx+1, column=6).fill = err_fill
        else:
            ws1.append([idx, row['folder'], row['total_files'], row['xlsx_files']])
            
        ws1.cell(row=idx+1, column=1).alignment = align_center
        ws1.cell(row=idx+1, column=3).alignment = align_center
        ws1.cell(row=idx+1, column=4).alignment = align_center

    # Sheet 2: Thống kê danh sách các con Bot (Bot ID)
    ws_bot = wb.create_sheet(title="Thong_Ke_Danh_Sach_Bot")
    ws_bot.append(["STT", "Mã Con Bot (Bot ID)", "Số Lượng Chiến Lược (File)", "Danh Sách Symbol Tài Sản", "Trạng Thái"])
    for cell in ws_bot[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center

    for idx, bot in enumerate(bot_summary, 1):
        ws_bot.append([idx, f"Bot {bot['bot_id']}", bot['file_count'], ", ".join(bot['symbols']), bot['status']])
        ws_bot.cell(row=idx+1, column=1).alignment = align_center
        ws_bot.cell(row=idx+1, column=2).alignment = align_center
        ws_bot.cell(row=idx+1, column=3).alignment = align_center
        ws_bot.cell(row=idx+1, column=5).alignment = align_center

    # Sheet 3 (Nếu đã check chi tiết): Danh sách các File/Folder LẠ hoặc SAI định dạng
    if is_detailed_checked:
        ws2 = wb.create_sheet(title="Danh_Sach_Muc_La_Va_Loi")
        ws2.append(["STT", "Tên Mục / File", "Loại Mục Lạ / Lỗi", "Mô Tả Chi Tiết / Lý Do", "Đường Dẫn Trong Drive", "File ID"])
        for cell in ws2[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = align_center

        if not unexpected_items:
            ws2.append([1, "-", "OK", "TẤT CẢ FILE VÀ FOLDER ĐỀU CHUẨN ĐỊNH DẠNG VÀ CẤU TRÚC!", "-", "-"])
        else:
            for idx, item in enumerate(unexpected_items, 1):
                ws2.append([idx, item.get('name', ''), item.get('type', ''), item.get('reason', ''), item.get('path', ''), item.get('id', '')])
                row_idx = idx + 1
                ws2.cell(row=row_idx, column=1).alignment = align_center
                ws2.cell(row=row_idx, column=2).font = err_font
                ws2.cell(row=row_idx, column=2).fill = err_fill
                ws2.cell(row=row_idx, column=3).alignment = align_center

    # Sheet 4: Danh sách toàn bộ file đã quét
    ws3 = wb.create_sheet(title="Danh_Sach_Toan_Bo_File")
    ws3.append(["STT", "Tên File", "Mã Bot", "Trạng Thái Tên", "Chi Tiết Định Dạng", "Đường Dẫn Trong Drive", "File ID"])
    for cell in ws3[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center

    for idx, f in enumerate(all_files, 1):
        status_text = "Hợp lệ" if f.get('is_valid_name', True) else "SAI định dạng / Lạ"
        bot_id_str = f.get('bot_id', '-')
        ws3.append([idx, f.get('name', ''), bot_id_str, status_text, f.get('name_reason', '-'), f.get('path', ''), f.get('id', '')])
        row_idx = idx + 1
        ws3.cell(row=row_idx, column=1).alignment = align_center
        ws3.cell(row=row_idx, column=3).alignment = align_center
        ws3.cell(row=row_idx, column=4).alignment = align_center
        if is_detailed_checked and not f.get('is_valid_name', True):
            ws3.cell(row=row_idx, column=4).font = err_font
            ws3.cell(row=row_idx, column=4).fill = err_fill

    # Tự động căn chỉnh độ rộng cột
    sheets = [ws1, ws_bot, ws3]
    if is_detailed_checked:
        sheets.append(ws2)

    for ws in sheets:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

    wb.save(output_filename)
    print(f"\n[EXCEL EXPORT] Đã xuất thành công báo cáo ra file Excel:")
    print(f" 📂 File Excel: {os.path.abspath(output_filename)}")


def process_drive_validation(url_or_id, service_mgr):
    folder_id = extract_folder_id(url_or_id)
    
    print("\n" + "="*105)
    print(f" 🚀 BƯỚC 1: ĐẾM SỐ LƯỢNG FILE VÀ THỐNG KÊ CON BOT (LINK DRIVE ID: {folder_id})")
    print("="*105)

    print("Đang quét cấu trúc danh sách thư mục và file trên Drive...")
    top_items = list_items_in_folder(service_mgr, folder_id)
    
    root_files = [item for item in top_items if item.get('mimeType') != 'application/vnd.google-apps.folder']
    asset_folders = [item for item in top_items if item.get('mimeType') == 'application/vnd.google-apps.folder']
    
    all_files = []
    raw_unexpected_items = []
    report_data = []

    if root_files:
        for rf in root_files:
            rf_name = rf.get('name', '')
            raw_unexpected_items.append({
                "id": rf.get('id'),
                "name": rf_name,
                "type": "File lạ tại Drive Root",
                "reason": "File nằm trực tiếp ở thư mục gốc Drive (không nằm trong folder tài sản nào)",
                "path": f"/[Drive Root]/{rf_name}"
            })
            all_files.append(rf)

    bot_map = {}

    for idx, asset in enumerate(asset_folders, 1):
        asset_id = asset.get('id')
        asset_name = asset.get('name', '').strip()
        asset_path = asset_name
        
        asset_upper = asset_name.upper()
        is_asset_known = any(sym in asset_upper for sym in VALID_SYMBOLS)
        if not is_asset_known:
            raw_unexpected_items.append({
                "id": asset_id,
                "name": asset_name,
                "type": "Folder lạ tại Drive Root",
                "reason": "Tên folder tài sản không chứa mã Symbol chuẩn hợp lệ (Ví dụ: BTCUSD, XAUUSD,...)",
                "path": f"/[Drive Root]/{asset_name}"
            })

        print(f"  [{idx}/{len(asset_folders)}] Đang đếm file trong folder: {asset_name} ...", end="", flush=True)
        files_in_asset, sub_unexpected = scan_asset_folder(service_mgr, asset_id, asset_path)
        raw_unexpected_items.extend(sub_unexpected)
        
        EXCEL_MIME_TYPES = {"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "application/vnd.ms-excel", "application/vnd.google-apps.spreadsheet"}
        xlsx_count = sum(1 for f in files_in_asset if f.get('name', '').lower().endswith(('.xlsx', '.xls')) or f.get('mimeType') in EXCEL_MIME_TYPES)

        for f in files_in_asset:
            fname = f.get('name', '')
            bot_id_match = re.match(r'^(\d{3})', fname)
            bot_id = bot_id_match.group(1) if bot_id_match else None
            f['bot_id'] = bot_id or "-"
            
            # Trích xuất sơ bộ Symbol
            extracted_sym = "-"
            fname_upper = fname.upper()
            for sym in VALID_SYMBOLS:
                if sym in fname_upper:
                    extracted_sym = sym
                    break
            f['symbol'] = extracted_sym

            if bot_id:
                if bot_id not in bot_map:
                    bot_map[bot_id] = []
                bot_map[bot_id].append(f)

        all_files.extend(files_in_asset)

        report_data.append({
            "folder": asset_name,
            "total_files": len(files_in_asset),
            "xlsx_files": xlsx_count,
            "raw_files_in_asset": files_in_asset
        })
        print(f" -> Done ({len(files_in_asset)} files)")

    # 1. Thống kê danh sách con Bot
    bot_summary = []
    sorted_bot_ids = sorted(bot_map.keys())
    for b_id in sorted_bot_ids:
        b_files = bot_map[b_id]
        symbols = sorted(list(set([f['symbol'] for f in b_files if f['symbol'] != "-"])))
        bot_summary.append({
            "bot_id": b_id,
            "file_count": len(b_files),
            "symbols": symbols if symbols else ["Không xác định"],
            "status": "Đủ chiến lược" if len(b_files) >= 4 else "Ít chiến lược"
        })

    # --- IN TỔNG SỐ LƯỢNG FILE BƯỚC 1 ---
    total_xlsx_count = sum(row['xlsx_files'] for row in report_data)
    print("\n" + "="*105)
    print(f" 📊 TỔNG SỐ LƯỢNG FILE: {len(all_files)} TỔNG FILE ({total_xlsx_count} file .xlsx) | Thuộc {len(asset_folders)} folder tài sản")
    print("="*105)

    print("\n" + "="*105)
    print(" 🤖 2. BẢNG THỐNG KÊ SỐ LƯỢNG CON BOT (BOT ID) VÀ SỐ CHIẾN LƯỢC")
    print("="*105)
    print(f"{'STT':<5} | {'Mã Con Bot':<15} | {'Số Lượng Chiến Lược (File)':<30} | {'Danh Sách Tài Sản (Symbol)':<35}")
    print("-" * 105)
    for i, bot in enumerate(bot_summary, 1):
        sym_str = ", ".join(bot['symbols'])
        print(f"{i:<5} | Bot {bot['bot_id']:<11} | {bot['file_count']:<30} | {sym_str:<35}")
    print("-" * 105)
    print(f"TỔNG CỘNG: {len(bot_summary)} CON BOT duy nhất được phát hiện!")
    print("="*105)

    # --- BƯỚC HỎI USER CÓ KIỂM TRA ĐỊNH DẠNG TÊN / MỤC LẠ KHÔNG ---
    print("\n" + "❓ " * 15)
    choice_check = input("👉 Đã thống kê xong số lượng. Bạn có muốn tiếp tục kiểm tra CHI TIẾT ĐỊNH DẠNG TÊN FILE & PHÁT HIỆN FILE/FOLDER LẠ không? (y/n, mặc định y): ").strip().lower()
    
    is_detailed_checked = False
    unexpected_items = []

    if choice_check in ["", "y", "yes"]:
        is_detailed_checked = True
        print("\n" + "="*105)
        print(" 🚀 BƯỚC 2: KIỂM TRA ĐỊNH DẠNG TÊN FILE & PHÁT HIỆN FILE / FOLDER LẠ")
        print("="*105)

        unexpected_items = list(raw_unexpected_items)

        for row in report_data:
            valid_in_asset = 0
            invalid_in_asset = 0
            unexpected_in_asset = 0

            for f in row['raw_files_in_asset']:
                fname = f.get('name', '')
                is_valid, reason, b_id, sym = check_filename_format(fname)
                f['is_valid_name'] = is_valid
                f['name_reason'] = reason

                if is_valid:
                    valid_in_asset += 1
                else:
                    invalid_in_asset += 1
                    unexpected_items.append({
                        "id": f.get('id'),
                        "name": fname,
                        "type": "Tên file sai định dạng",
                        "reason": reason,
                        "path": f.get('path')
                    })

            row['valid_name_files'] = valid_in_asset
            row['invalid_name_files'] = invalid_in_asset
            
            # Đếm mục lạ thuộc asset folder này
            asset_prefix = f"/{row['folder']}/"
            asset_unexp = sum(1 for item in unexpected_items if item.get('path', '').startswith(row['folder']))
            row['unexpected_count'] = asset_unexp

        total_unexpected = len(unexpected_items)
        
        print("\n" + "="*105)
        print(f" 📊 TỔNG KẾT BƯỚC 2: PHÁT HIỆN {total_unexpected} MỤC LẠ & LỖI ĐỊNH DẠNG TÊN!")
        print("="*105)

        if unexpected_items:
            print("\n" + "⚠️ " * 20)
            print(f" ❌ DANH SÁCH MỤC LẠ / FILE SAI ĐỊNH DẠNG (Tối đa 25 mục đầu tiên):")
            print("="*105)
            for idx, item in enumerate(unexpected_items[:25], 1):
                print(f"  {idx}. [{item['type']}] {item['name']} -> {item['reason']} (Đường dẫn: {item['path']})")
            if total_unexpected > 25:
                print(f"  ... và còn {total_unexpected - 25} mục khác (Xem chi tiết trong file Excel).")
            print("="*105)
        else:
            print("\n  🎉 CHÚC MỪNG! CHÍNH XÁC TOÀN BỘ FILE VÀ CẤU TRÚC DRIVE ĐỀU CHUẨN HỢP LỆ!")
    else:
        print("\n  ⏩ Đã bỏ qua bước kiểm tra chi tiết định dạng tên file.")

    # --- BƯỚC HỎI USER CÓ XUẤT FILE EXCEL KHÔNG ---
    print("\n" + "❓ " * 15)
    choice_excel = input("👉 Bạn có muốn xuất toàn bộ báo cáo kết quả ra FILE EXCEL (.xlsx) không? (y/n, mặc định y): ").strip().lower()
    
    if choice_excel in ["", "y", "yes"]:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_filename = os.path.join(script_dir, f"Bao_Cao_Thong_Ke_Drive_{timestamp_str}.xlsx")
        export_count_report_to_excel(excel_filename, report_data, bot_summary, unexpected_items, all_files, is_detailed_checked)
    else:
        print("\n  ⏩ Đã bỏ qua bước xuất file Excel.")

    return {
        "total_bots": len(bot_summary),
        "total_files": len(all_files),
        "report_data": report_data
    }


def main():
    print("==========================================================================")
    print("  TOOL THỐNG KÊ SỐ LƯỢNG FILE, CON BOT & CHECK VALIDATION DRIVE  ")
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

    print("\nDONE! Hoàn thành quy trình kiểm tra.")


if __name__ == "__main__":
    main()
