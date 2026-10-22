# SheetReason — LLM-Powered Excel Question Answering

Hệ thống hỏi đáp bằng ngôn ngữ tự nhiên trên các workbook Excel, kết hợp suy luận của LLM với thực thi Python và bằng chứng từ ô dữ liệu nguồn. SheetReason phân tích cấu trúc bảng, ô gộp, công thức và định dạng ô; tạo nhiều lời giải rồi đối chiếu kết quả bằng bỏ phiếu và trọng tài LLM khi cần.

Phiên bản này là bài nộp **AI2026 Bảng C — AI26C0024**. Chương trình dùng `qwen/qwen3.7-flash` qua OpenRouter theo cấu hình bản nộp; model và API endpoint được truyền qua tham số dòng lệnh. Mã đội và tên gói nộp được giữ nguyên.

## Lấy mã nguồn

```bash
git clone https://github.com/VPhunggggg/sheetreason-excel-qa.git
cd sheetreason-excel-qa
```

## Luồng xử lý

1. Đọc workbook bằng `openpyxl`, lập hồ sơ các sheet và vùng bảng theo câu hỏi.
2. Suy luận kiểu đáp án cần trả về, sinh hàm Python `solve(file_path)` và chạy trong tiến trình con có timeout.
3. Kiểm tra kết quả và bằng chứng ô nguồn, sửa mã khi lỗi; chọn đáp án từ tối đa ba ứng viên bằng đa số hoặc trọng tài LLM.
4. Định dạng đáp án, ghi CSV nguyên tử, kiểm tra ID và lưu log chẩn đoán riêng.

Công nghệ chính: Python, OpenRouter, `openpyxl`, pandas, NumPy và Docker. Chương trình chạy trên CPU.

## 1. Cấu trúc bài nộp

Tên gói: `AI26C0024.zip`, giải nén trực tiếp được cấu trúc sau:

```text
Dockerfile
.dockerignore
requirements.txt
README.md
main.py
excelqa.py
scripts/__init__.py
scripts/validate_submission.py
scripts/package_submission.py
```

`Dockerfile` nằm ngay ở gốc ZIP. Gói chỉ chứa source và cấu hình cần thiết; không chứa API key, workbook mẫu, đáp án chuẩn, submission cũ, log, virtualenv hay Docker image.

## 2. Build và chạy Docker

Yêu cầu Docker với Linux containers. Ảnh dùng Python 3.11, Linux amd64, CPU; không cần GPU.

Build tại thư mục vừa giải nén:

```bash
docker build --platform linux/amd64 -t ai26c0024:submission .
```

Thư mục input do hệ thống cung cấp có `queries.json` và thư mục `xlsx_files/`. Tạo một thư mục output có quyền ghi và thay các đường dẫn mount/API key trong lệnh sau:

```bash
docker run --rm --platform linux/amd64 --cpus 6 --memory 12g \
  -v "/absolute/path/input:/input:ro" \
  -v "/absolute/path/output:/output:rw" \
  ai26c0024:submission \
  --xlsx-dir /input/xlsx_files \
  --queries /input/queries.json \
  --output /output/submission.csv \
  --openrouter-key "<OPENROUTER_KEY>" \
  --openrouter-model "qwen/qwen3.7-flash" \
  --openrouter-base-url "https://openrouter.ai/api/v1"
```

API key là giá trị hệ thống chấm truyền vào lúc chạy. Không cần tạo file `.env`, sửa source hoặc đặt key trong Dockerfile. Sau khi build, ảnh đã có toàn bộ source và dependency; khi chạy cần mạng để gọi OpenRouter.

`ENTRYPOINT ["python", "main.py"]` nhận đúng sáu tham số:

| Tham số | Ý nghĩa |
|---|---|
| `--xlsx-dir` | Thư mục workbook đầu vào |
| `--queries` | Đường dẫn JSON câu hỏi |
| `--output` | Đường dẫn CSV cần ghi |
| `--openrouter-key` | API key được cấp bên ngoài |
| `--openrouter-model` | Model do hệ thống chỉ định |
| `--openrouter-base-url` | Base URL API do hệ thống chỉ định |

Chương trình không yêu cầu tên thư mục input cố định và không tìm dữ liệu ngoài `--xlsx-dir`. Không ghi workbook hoặc sửa dữ liệu đầu vào. Mọi kết quả và chẩn đoán dạng file được ghi cạnh đường dẫn `--output`; log tiến độ ghi ra stderr.

## 3. Chạy trực tiếp bằng Python 3.11

Linux:

```bash
python3.11 -m venv .venv
. .venv/bin/activate
python -m pip install -r requirements.txt
```

Windows PowerShell:

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Sau khi cài dependency, gọi `python -X utf8 main.py` với cùng sáu tham số ở trên. Với PowerShell, có thể dùng trực tiếp `.\.venv\Scripts\python.exe -X utf8 main.py`; dùng dấu backtick để xuống dòng thay cho dấu `\` trong ví dụ Bash.

## 4. Dữ liệu và kết quả

`queries.json` là một mảng các object có `id`, `file_name`, `query`. Ví dụ minh họa cấu trúc, không phải dữ liệu được đóng sẵn trong bài:

```json
[{"id": 1, "file_name": "input.xlsx", "query": "What is the total for 2023?"}]
```

CSV UTF-8 có đúng hai cột `id,answer`. Mỗi ID xuất hiện đúng một lần; answer chỉ chứa đáp án cuối, không có giải thích hoặc code. Các ID được sắp tăng dần. Dấu phẩy trong đáp án được quote theo chuẩn CSV; không tự cắt tên hoặc làm mất chữ số 0 tận cùng được yêu cầu.

CSV được ghi nguyên tử sau mỗi câu hoàn tất và được kiểm tra độc lập trước khi chương trình thoát 0. Lỗi đầu vào/ghi file/CSV không hợp lệ trả mã khác 0. Câu chưa giải được xuất `UNVALUE` và ghi nguyên nhân trong `<output>.errors.json`; code và bằng chứng nằm trong `<output>.debug.log`. Các file phụ không được trộn vào CSV.

Kiểm tra cấu trúc CSV, không gọi API:

```bash
python scripts/validate_submission.py --queries /input/queries.json --submission /output/submission.csv
```

## 5. Phương pháp và ngân sách xử lý

- Một lượt LLM suy luận schema định dạng từ câu hỏi, reasoning `low`; solver và trọng tài dùng reasoning `medium`.
- Tối đa ba ứng viên. Hai ứng viên đầu chạy song song; đa số dựa trên chuỗi đáp án đã định dạng chính xác. Nếu bất đồng, trọng tài xác minh code và bằng chứng ô nguồn. Không đoán từ ID/file_name hoặc tra đáp án mẫu.
- Python đọc được toàn bộ workbook; profile 16000 ký tự là ngân sách ngữ cảnh mềm. Header, bảng con, ô gộp, tổng, indent và style là gợi ý để suy luận, không tự ép phép tính.
- Mỗi ứng viên được sửa code tối đa hai lần; mỗi lần thực thi có timeout 60 giây. HTTP tối đa năm lần thử, timeout 60 giây/lần. Completion bắt đầu 4096 token, tối đa 8192 khi bị cắt; bước định dạng bắt đầu 1500 token.
- Tối đa sáu câu song song, đồng thời tôn trọng CPU affinity/quota nhỏ hơn. Các thư viện số chạy một luồng. Không dùng GPU.
- Giá trị số/ngày mơ hồ cần solver chọn cách đọc theo ngữ cảnh. Công thức chưa tính được giữ lại cho solver; không tự thay dữ liệu thiếu bằng 0. Định dạng số cuối dùng Decimal/ROUND_HALF_UP.
- Phí được cộng từ `usage.cost`, gồm các lượt retry/truncated có usage. Thiếu thông tin phí được ghi `unknown`, không giả định bằng 0. Reasoning tính vào token đầu ra.

Các giới hạn trên kiểm soát retry và tài nguyên từng bước, không phải cơ chế hard cap chi phí API hoặc deadline toàn chương trình. Chi phí và thời gian trên tập ẩn phụ thuộc độ khó, kích thước workbook và dịch vụ OpenRouter.

## 6. Kiểm chứng được ghi nhận cho bản nộp

Các kết quả dưới đây được giữ từ tài liệu trong ZIP gốc. Bộ test hồi quy và dữ liệu đánh giá không có trong gói mã nguồn này.

- Build và chạy được Linux amd64; image khoảng 439 MB, nhỏ hơn 1 GB.
- 20/20 test hồi quy chạy trong Docker với 6 CPU và 12 GB RAM, không cần API.
- Bản reasoning đã đạt 30/30 câu Public data 2 ở một lượt chạy Windows. Đây là kiểm tra công khai, không bảo đảm điểm hidden set.
- Kết quả kiểm tra API thật trên container được ghi trong báo cáo bàn giao riêng, không đóng dữ liệu test/log vào ZIP.

## 7. Đóng gói

```bash
python scripts/package_submission.py
```

Gói được tạo tại `dist/AI26C0024.zip`. Script chỉ đưa các file source cho phép vào gói, kiểm tra cú pháp/import, quét credential, kiểm tra CRC ZIP và giới hạn 100 MB. Không nộp image Docker hoặc CSV public thay cho ZIP source.
