# MT5 Optimize Local Pipeline

## 1. Mục đích

Đây là bộ script Python chạy MT5 trực tiếp trên máy local, không đi qua web, backend hoặc MT5 Worker. Pipeline có hai phần chính:

```text
Open Price Optimization
→ lọc kết quả và chọn bộ tham số
→ tạo file INI Every Tick
→ chạy Every Tick
→ lưu report XLSX
```

## 2. Cấu trúc thư mục MT5

Các thư mục được tạo bên trong:

```text
C:\Users\XUAN SON\AppData\Roaming\MetaQuotes\Terminal\D0E8209F77C8CF37AD8BF550E51FF075\MQL5\Files
```

### `Files_ini`

Thư mục đầu vào. Đặt các file `.ini` dùng để chạy Open Price Optimization vào đây.

### `Optimize_xml`

Nơi MT5/Python lưu tạm kết quả Open Price dạng XML hoặc CSV. Các file này chứa thông số đầu vào và các chỉ số như Profit, Trades, Profit Factor, Recovery Factor, Drawdown và Expected Payoff.

### `Optimize_filter`

Nơi lưu các file CSV sau khi đã lọc kết quả Open Price. Folder này được sử dụng bởi pipeline filter nâng cao.

### `Backtest_ini`

Nơi lưu các file `.ini` Every Tick được tạo từ các bộ tham số đã chọn.

### `Report_backtest`

Nơi lưu report thô sau khi MT5 chạy Every Tick.

### `Report_backtest\raw`

Nơi giữ các report thô trước hoặc trong quá trình chuyển đổi sang Excel.

### `Report_backtest\xlsx`

Nơi lưu file XLSX cuối cùng sau khi Every Tick hoàn tất.

## 3. Các file Python

### `config.py`

File cấu hình đường dẫn MT5:

- `MT5_PATH`: đường dẫn tới `terminal64.exe`.
- `DEFAULT_TERMINAL_DATA_PATH`: đường dẫn MT5 Data Folder.

Khi import file này, các folder pipeline còn thiếu sẽ được tự động tạo.

### `0_run_optimized.py`

Pipeline gộp dùng để chạy nhanh:

```text
Open Price
→ lọc cơ bản
→ tạo file INI Every Tick
```

File này chạy trực tiếp MT5, đọc kết quả optimization và tạo các file INI trong `Backtest_ini`.

Lưu ý: filter trong file này là filter cơ bản và lấy Top 10 cố định. Nó chưa sử dụng đầy đủ pipeline Stable/Hard Flags/Dedup của `2_filter_optimize.py`.

### `1_save_optimize.py`

Module dùng để chạy Open Price và lưu XML. File này hiện được thiết kế như thư viện cho `0_run_optimized.py`, không chạy độc lập trực tiếp.

### `2_filter_optimize.py`

Bộ lọc nâng cao. Đọc XML/CSV trong `Optimize_xml`, sau đó tạo CSV đã lọc trong `Optimize_filter`.

### `3_convert_ini.py`

Đọc CSV đã lọc trong `Optimize_filter` và tạo các file INI Every Tick trong `Backtest_ini`.

### `4_run_save_backtest.py`

Chạy Every Tick cho tất cả các file INI trong `Backtest_ini`, đọc report MT5 và tạo XLSX trong `Report_backtest\xlsx`.

## 4. Filter cơ bản

Filter cơ bản nằm trong `0_run_optimized.py`.

Luồng xử lý:

```text
Lọc 6 điều kiện
→ ưu tiên Result trong khoảng 5–8
→ ưu tiên Trades cao
→ ưu tiên Equity DD gần 10%
→ ưu tiên Profit cao
→ lấy Top 10
```

Các điều kiện gồm:

- Min Profit.
- Min/Max Trades.
- Min Recovery Factor.
- Min/Max Profit Factor.
- Min/Max Equity Drawdown.

Filter cơ bản không có:

- Hard Flags.
- Stable Score.
- Plateau Score.
- Loại các bộ tham số gần nhau.
- Top N tùy chỉnh đầy đủ.

## 5. Filter nâng cao

Filter nâng cao nằm trong `2_filter_optimize.py`.

Luồng xử lý:

```text
Hard Flags
→ filter gốc
→ Composite Score
→ Plateau Score
→ Final Score
→ loại bộ tham số quá gần nhau
→ lấy Top N
```

### Hard Flags

Loại các kết quả có rủi ro rõ ràng trước khi xếp hạng:

- Expected Payoff không đạt.
- Profit Factor thấp.
- Recovery Factor thấp.
- P/D Ratio thấp.

### Composite Score

Chuẩn hóa các chỉ số bằng Z-Score rồi kết hợp:

- Profit.
- Profit Factor.
- Recovery Factor.
- Sharpe Ratio.
- Expected Payoff.
- Total Trades.
- Drawdown.

### Plateau Score

Kiểm tra vùng lân cận của một bộ tham số. Nếu các bộ gần đó cũng có điểm tốt thì bộ đang xét được đánh giá ổn định hơn.

### Dedup và khoảng cách tham số

Sau khi xếp hạng, script bỏ các bộ có tham số quá gần nhau để tránh chọn nhiều bộ gần như giống hệt.

### Top N

Mặc định:

```python
TOP_N = 10
```

Có thể đặt bằng PowerShell:

```powershell
$env:FILTER_TOP_N=5
```

## 6. Các lệnh chạy

Mở PowerShell tại thư mục source:

```powershell
cd "C:\Users\XUAN SON\Downloads\Optimize"
```

### Pipeline hiện chạy được ngay

Lệnh 1: Open Price, filter cơ bản và tạo INI Every Tick:

```powershell
python 0_run_optimized.py --num-runs 1
```

Chạy Open Price 3 lần:

```powershell
python 0_run_optimized.py --num-runs 3
```

Lệnh 2: chạy Every Tick và tạo XLSX:

```powershell
python 4_run_save_backtest.py
```

Có thể chạy nối tiếp trong một lệnh PowerShell:

```powershell
python 0_run_optimized.py --num-runs 1; if ($LASTEXITCODE -eq 0) { python 4_run_save_backtest.py }
```

### Pipeline filter nâng cao

Về thiết kế, pipeline nâng cao gồm:

```powershell
python 1_save_optimize.py
python 2_filter_optimize.py
python 3_convert_ini.py
python 4_run_save_backtest.py
```

Tuy nhiên trong phiên bản hiện tại, `1_save_optimize.py` bị khóa chạy độc lập và chỉ được dùng như module cho `0_run_optimized.py`. Vì vậy bốn lệnh trên chưa thể chạy liền thành pipeline nâng cao hoàn chỉnh nếu chưa chỉnh lại source.

## 7. Tham số filter của pipeline gộp

Có thể truyền trực tiếp cho `0_run_optimized.py`:

```powershell
python 0_run_optimized.py `
  --num-runs 1 `
  --min-profit 100000 `
  --min-trades 100 `
  --max-trades 1200 `
  --min-recovery-factor 2 `
  --min-profit-factor 1 `
  --max-profit-factor 100 `
  --min-equity-dd 6 `
  --max-equity-dd 13
```

## 8. Kết quả đầu ra

Sau Open Price:

```text
MQL5\Files\Backtest_ini
```

Sau Every Tick:

```text
MQL5\Files\Report_backtest\xlsx
```

File XLSX cuối cùng nằm tại:

```text
C:\Users\XUAN SON\AppData\Roaming\MetaQuotes\Terminal\D0E8209F77C8CF37AD8BF550E51FF075\MQL5\Files\Report_backtest\xlsx
```

## 9. Cảnh báo

- `0_run_optimized.py` sẽ dọn các file cũ trong `Optimize_xml`, `Optimize_filter` và `Backtest_ini` trước khi chạy.
- Không xóa folder hoặc file trong lúc MT5 đang chạy.
- `4_run_save_backtest.py` sẽ chạy tất cả file INI có trong `Backtest_ini`.
- Muốn Every Tick chạy ít bộ hơn thì phải giảm số file INI được tạo ra.
- Source local này không lưu dữ liệu vào database Lotusquant.
- Source local này không dùng web, API, worker, phân quyền user hoặc batch history.

python 0_run_optimized.py --num-runs 1
python 4_run_save_backtest.py
