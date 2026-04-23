# Trộn Câu Hỏi - Chương Trình Trộn Đề Thi Tự Động

Chương trình Python đọc file DOCX đề thi gốc, tự động nhận diện câu hỏi và đáp án đúng (qua formatting màu đỏ hoặc gạch chân), cho phép người dùng phân nhóm câu hỏi qua console, sau đó trộn đề theo nhiều chế độ khác nhau và xuất ra các file DOCX đề trộn + bảng đáp án Excel.

---

## Mục Lục

- [Tính Năng Chính](#tính-năng-chính)
- [Cấu Trúc Thư Mục](#cấu-trúc-thư-mục)
- [Yêu Cầu Hệ Thống](#yêu-cầu-hệ-thống)
- [Cài Đặt](#cài-đặt)
- [Hướng Dẫn Sử Dụng](#hướng-dẫn-sử-dụng)
  - [Chạy Full Pipeline](#chạy-full-pipeline)
  - [Cú Pháp Phân Nhóm](#cú-pháp-phân-nhóm)
  - [Ví Dụ Sử Dụng](#ví-dụ-sử-dụng)
- [Mô Tả Các Module](#mô-tả-các-module)
  - [main.py](#mainpy)
  - [parser.py](#parserpy)
  - [models.py](#modelspy)
  - [shuffler.py](#shufflerpy)
  - [writer.py](#writerpy)
  - [console_ui.py](#console_uipy)
- [Debug Từng Module](#debug-từng-module)
- [Cấu Trúc Dữ Liệu](#cấu-trúc-dữ-liệu)
- [Thuật Toán Trộn](#thuật-toán-trộn)
- [Các Trường Hợp Đặc Biệt](#các-trường-hợp-đặc-biệt)
- [Định Dạng Đầu Ra](#định-dạng-đầu-ra)

---

## Tính Năng Chính

- Tự động parse file DOCX đề thi, nhận diện câu hỏi theo nhiều định dạng (Câu X, Question X, số thứ tự).
- Nhận diện đáp án đúng qua formatting: màu đỏ (RGB) hoặc gạch chân.
- Hỗ trợ 3 kiểu bố trí đáp án: mỗi đáp án 1 dòng, tất cả trên 1 dòng, 2 đáp án/dòng.
- Phát hiện tự động các section (PART, SECTION, PHAN) trong đề thi.
- Phát hiện và xử lý đoạn văn đọc hiểu (context block): các câu trong block di chuyển cùng nhau khi trộn.
- 4 chế độ trộn linh hoạt: không trộn (g0), trộn câu hỏi (g1), trộn đáp án (g2), trộn cả hai (g3).
- Hỗ trợ cố định vị trí nhóm (dùng dấu `#` trước nhóm).
- Đảm bảo không trùng đề (fingerprint-based).
- Xuất N file DOCX đề trộn giữ nguyên 100% formatting gốc (clone XML).
- Xuất bảng đáp án Excel với 2 sheet: đáp án tổng hợp và chi tiết mapping.
- Cho phép nhập thủ công đáp án đúng cho các câu không phát hiện được.

---

## Cấu Trúc Thư Mục

```
TronCauHoi/
|-- main.py                 # Entry point CLI, chạy toàn bộ pipeline
|-- models.py               # Các data class: ExamDocument, Question, OptionData, GroupConfig, QuestionBlock
|-- parser.py               # Parse file DOCX, trích xuất câu hỏi + đáp án + nhận diện đáp án đúng
|-- shuffler.py             # Thuật toán trộn đề (Fisher-Yates shuffle)
|-- writer.py               # Xuất file DOCX đề trộn + Excel bảng đáp án
|-- console_ui.py           # Giao diện console: hiển thị câu hỏi, nhận input phân nhóm từ user
|-- requirements.txt        # Danh sách thư viện Python cần thiết
|-- implementation_plan.md  # Tài liệu thiết kế chi tiết
|-- debug/                  # Các script debug
|   |-- debug_formatting.py      # Debug nhận diện formatting đáp án đúng
|   |-- debug_shuffle.py         # Debug thuật toán trộn
|-- output/                 # Thư mục chứa đề trộn và bảng đáp án (tự động tạo)
|-- venv/                   # Virtual environment (nếu có)
|-- __pycache__/            # Python cache (tự động tạo)
```

---

## Yêu Cầu Hệ Thống

- Python 3.8 trở lên
- Pip (trình quản lý gói Python)

---

## Cài Đặt

### Bước 1: Clone hoặc tải về dự án

```bash
cd duong/dan/toi/thu/muc
git clone <repository-url> TronCauHoi
cd TronCauHoi
```

Hoặc giải nén file zip vào thư mục TronCauHoi.

### Bước 2: Tạo virtual environment (khuyến nghị)

```bash
python -m venv venv
```

Kích hoạt virtual environment:

- Windows (PowerShell):
  ```powershell
  .\venv\Scripts\Activate.ps1
  ```
- Windows (CMD):
  ```cmd
  .\venv\Scripts\activate.bat
  ```
- Linux/macOS:
  ```bash
  source venv/bin/activate
  ```

### Bước 3: Cài đặt thư viện

```bash
pip install -r requirements.txt
```

Danh sách thư viện:

| Thư viện            | Phiên bản  | Mục đích                            |
|---------------------|------------|-------------------------------------|
| python-docx         | >=1.2.0    | Đọc/ghi file DOCX                   |
| openpyxl            | >=3.1.5    | Xuất file Excel bảng đáp án         |
| lxml                | >=6.0.4    | Thao tác XML (đi kèm python-docx)   |
| colorama            | >=0.4.6    | Hiển thị màu sắc trên console       |
| pillow              | >=12.2.0   | Xử lý hình ảnh (nếu có trong đề)    |

---

## Hướng Dẫn Sử Dụng

### Chạy Full Pipeline

Lệnh cơ bản:

```bash
python main.py <duong-dan-file-docx>
```

Với tùy chọn:

```bash
python main.py <file.docx> --output <thu-muc-output>
python main.py <file.docx> --output <thu-muc-output> --detail 1-5
```

Tham số:

| Tham số        | Mô tả                                                        | Mặc định                         |
|----------------|---------------------------------------------------------------|----------------------------------|
| `input`        | Đường dẫn file DOCX đề gốc (bắt buộc)                       | -                                |
| `--output, -o` | Thư mục xuất kết quả                                         | `output/` cùng thư mục với input |
| `--detail, -d` | Hiển thị chi tiết câu hỏi (VD: `1-5`, `10`, `15-20`)         | Không hiển thị                   |

### Cú Pháp Phân Nhóm

Khi chương trình hiển thị danh sách câu hỏi và yêu cầu nhập phân nhóm, sử dụng cú pháp sau:

```
<nhóm>: <dải câu>
```

Các loại nhóm:

| Nhóm  | Ý nghĩa                                     |
|-------|----------------------------------------------|
| `g0`  | Không trộn gì (giữ nguyên vị trí và đáp án) |
| `g1`  | Chỉ trộn thứ tự câu hỏi                     |
| `g2`  | Chỉ trộn thứ tự đáp án A/B/C/D              |
| `g3`  | Trộn cả thứ tự câu hỏi lẫn đáp án           |

Thêm dấu `#` trước nhóm để cố định vị trí nhóm đó (không bị hoán đổi với nhóm khác):

```
#g0: 1-10       <-- Cố định vị trí, không trộn gì
g3: 11-30        <-- Trộn cả hai, có thể đổi vị trí với nhóm khác
g3: 31-50        <-- Trộn cả hai
```

Nhấn Enter trong khi chưa nhập gì để mặc định tất cả câu là g3 (trộn hết).

### Ví Dụ Sử Dụng

**Ví dụ 1: Đề 50 câu, trộn tất cả**

```bash
python main.py de_goc.docx
```

Sau đó nhập:

```
> [Enter]
```

Tất cả câu 1-50 sẽ được trộn cả câu hỏi lẫn đáp án.

**Ví dụ 2: Đề có phần Listening + Reading + Writing**

```bash
python main.py de_thi_tieng_anh.docx --output ket_qua/
```

Sau đó nhập:

```
> g0: 1-10
> #g3: 11-30
> g3: 31-50
> [Enter]
```

Giải thích:
- Câu 1-10 (Listening): Không trộn gì.
- Câu 11-30 (Reading): Trộn cả hai, cố định vị trí.
- Câu 31-50 (Writing): Trộn cả hai.

Sau đó nhập số đề cần tạo (hoặc Enter để mặc định 4 đề).

**Ví dụ 3: Xem chi tiết câu hỏi 1-5 trước khi trộn**

```bash
python main.py de_goc.docx --detail 1-5,10
```

---

## Mô Tả Các Module

### main.py

Entry point của chương trình. Thực hiện pipeline theo thứ tự:

1. Parse arguments dòng lệnh (argparse).
2. Gọi `parse_docx()` để đọc và phân tích file DOCX.
3. Gọi `detect_sections()` để phát hiện section tự động.
4. Hiển thị tóm tắt đề thi, danh sách câu hỏi, sections.
5. Xử lý câu chưa có đáp án (cho phép nhập thủ công).
6. Nhận input phân nhóm từ user.
7. Gọi `generate_unique_exams()` để trộn đề.
8. Gọi `write_all_outputs()` để xuất file DOCX + Excel.

### parser.py

Module parse file DOCX. Class chính: `DocxParser`.

Chức năng:
- Thu thập tất cả paragraphs từ file DOCX.
- Nhận diện câu hỏi theo pattern: `Câu X:`, `Question X.`, `X. `.
- Nhận diện đáp án A/B/C/D theo 3 kiểu bố trí (dòng riêng, cùng dòng, 2 dòng).
- Nhận diện đáp án đúng qua formatting:
  - Màu đỏ (RGB): R > 180, G < 100, B < 100.
  - Gạch chân (underline).
  - Ưu tiên: Red > Underline.
- Phát hiện context blocks (đoạn văn đọc hiểu) và gán vào nhóm câu hỏi tương ứng.
- Phát hiện section headers (PART, SECTION, PHAN).

### models.py

Các data class:

- `OptionData`: Dữ liệu 1 đáp án (A/B/C/D) gồm chữ cái gốc, chữ cái hiện tại, nội dung, trạng thái đúng/sai, XML elements.
- `Question`: Dữ liệu 1 câu hỏi gồm số thứ tự gốc, số thứ tự mới, nội dung stem, đáp án đúng, các options, context, thông tin trộn.
- `QuestionBlock`: Khối câu hỏi đi kèm ngữ cảnh chung (đọc hiểu). Khi trộn, cả block di chuyển cùng nhau.
- `GroupConfig`: Cấu hình 1 nhóm câu hỏi do user nhập, gồm loại nhóm (0-3), cố định hay không, dải câu hỏi.
- `ExamDocument`: Dữ liệu toàn bộ đề thi gồm filepath, header elements, danh sách câu hỏi, blocks.

### shuffler.py

Thuật toán trộn đề thi. Sử dụng Fisher-Yates Shuffle (qua `random.shuffle`).

Chức năng:
- `shuffle_options()`: Hoán vị đáp án A/B/C/D của 1 câu hỏi.
- `shuffle_questions_in_group()`: Hoán vị thứ tự câu hỏi trong 1 nhóm, giữ câu hỏi thuộc cùng block đi cùng nhau.
- `shuffle_exam()`: Tạo 1 bản trộn từ đề gốc theo cấu hình phân nhóm.
- `generate_unique_exams()`: Sinh N đề trộn đảm bảo không trùng nhau (dựa trên fingerprint).
- `compute_signature()`: Tạo fingerprint của 1 đề trộn.

### writer.py

Module xuất file.

Chức năng:
- `write_shuffled_docx()`: Xuất 1 file DOCX đề trộn. Strategy: clone toàn bộ document gốc (giữ styles, page setup, fonts), xóa body, clone lại paragraphs theo thứ tự mới. Xử lý thay số câu hỏi, thay chữ cái đáp án, cập nhật formatting đáp án đúng.
- `write_answer_key_excel()`: Xuất bảng đáp án Excel với 2 sheet:
  - Sheet 1 "Đáp án": Câu | Đề gốc | Mã 001 | Mã 002 | ...
  - Sheet 2 "Chi tiết mapping": Đề | Câu mới | Câu gốc | Đáp án gốc | Đáp án mới | Trộn đáp án
- `write_all_outputs()`: Xuất tất cả đề trộn + bảng đáp án vào thư mục output.

### console_ui.py

Giao diện console tương tác.

Chức năng:
- `display_exam_summary()`: Hiển thị tóm tắt đề thi (số câu, số đáp án phát hiện).
- `display_questions_list()`: Hiển thị danh sách câu hỏi dạng bảng ngắn gọn.
- `display_questions_detail()`: Hiển thị chi tiết câu hỏi (cho debug).
- `display_sections()`: Hiển thị sections phát hiện tự động.
- `get_group_input()`: Nhận input phân nhóm từ user.
- `validate_groups()`: Validate phân nhóm (kiểm tra overlap, coverage).
- `get_num_exams()`: Nhận số lượng đề trộn cần tạo.
- `get_manual_answers()`: Cho phép user nhập thủ công đáp án đúng.
- `apply_manual_answers()`: Áp dụng đáp án thủ công vào exam.

---

## Debug Từng Module

Mỗi file đều có `if __name__ == "__main__"` để chạy độc lập:

```bash
# Test parse DOCX
python parser.py de_goc.docx

# Test hiển thị console + nhập nhóm
python console_ui.py de_goc.docx

# Test shuffle với dữ liệu mẫu
python shuffler.py --demo

# Test shuffle với file DOCX thật
python shuffler.py de_goc.docx "g3: 1-50" --num 4

# Test xuất DOCX + Excel
python writer.py de_goc.docx --test

# Test data models
python models.py

# Chạy full pipeline
python main.py de_goc.docx
```

Các script debug bổ sung nằm trong thư mục `debug/`:
- `debug_formatting.py`: Kiểm tra nhận diện formatting đáp án đúng.
- `debug_shuffle.py`: Kiểm tra thuật toán trộn chi tiết.

---

## Cấu Trúc Dữ Liệu

```
ExamDocument
|-- filepath: str                    # Đường dẫn file DOCX
|-- header_elements: list            # XML elements phần đầu đề thi
|-- header_text: str                 # Text phần đầu
|-- questions: List[Question]        # Danh sách câu hỏi
|-- blocks: List[QuestionBlock]      # Danh sách blocks đọc hiểu
|-- total_questions: int             # Tổng số câu hỏi

Question
|-- original_number: int             # Số thứ tự gốc
|-- new_number: int                  # Số thứ tự sau trộn
|-- stem_text: str                   # Nội dung câu hỏi (text thuần)
|-- stem_elements: list              # XML elements thân câu hỏi
|-- correct_answer: str              # Đáp án đúng hiện tại (A/B/C/D)
|-- original_correct_answer: str     # Đáp án đúng gốc
|-- options: Dict[str, OptionData]   # 4 đáp án {'A': ..., 'B': ..., 'C': ..., 'D': ...}
|-- context_elements: list           # XML elements ngữ cảnh (đọc hiểu)
|-- context_text: str                # Text ngữ cảnh
|-- options_were_shuffled: bool      # Đáp án có bị trộn không
|-- option_mapping: Dict[str, str]   # Mapping đáp án mới -> gốc
|-- block_id: int (optional)         # ID block đọc hiểu (nếu có)
|-- is_inline: bool                  # Đáp án nằm trên cùng dòng với stem

OptionData
|-- original_letter: str             # Chữ cái gốc (A/B/C/D)
|-- current_letter: str              # Chữ cái hiện tại (sau trộn)
|-- text: str                        # Nội dung đáp án (text thuần)
|-- is_correct: bool                 # Có phải đáp án đúng không
|-- elements: list                   # XML paragraph elements gốc

GroupConfig
|-- group_type: int                  # 0/1/2/3
|-- fixed: bool                      # Cố định vị trí (dấu #)
|-- question_ranges: List[Tuple]     # Dải câu hỏi [(start, end), ...]

QuestionBlock
|-- block_id: int                    # ID block
|-- context_elements: list           # Paragraph elements đoạn văn
|-- context_text: str                # Text đoạn văn
|-- question_numbers: List[int]      # Các câu thuộc block
```

---

## Thuật Toán Trộn

Chương trình sử dụng thuật toán Fisher-Yates Shuffle (đã được implement sẵn trong `random.shuffle()` của Python).

Quy trình trộn 1 đề:

1. **Bước A - Trộn bên trong mỗi nhóm:**
   - g0: Không làm gì.
   - g1: Hoán vị thứ tự câu hỏi (giữ đáp án nguyên).
   - g2: Hoán vị đáp án A/B/C/D (giữ thứ tự câu nguyên).
   - g3: Hoán vị cả thứ tự câu lẫn đáp án.
   - Câu hỏi thuộc cùng QuestionBlock sẽ di chuyển cùng nhau.

2. **Bước B - Trộn vị trí giữa các nhóm:**
   - Nhóm có dấu `#` (fixed) giữ nguyên vị trí.
   - Các nhóm không có `#` có thể bị hoán đổi vị trí với nhau.

3. **Bước C - Ghép câu hỏi theo thứ tự nhóm mới.**

4. **Bước D - Đánh số lại 1 -> N.**

Đảm bảo không trùng đề: Mỗi đề được tính fingerprint (signature) dựa trên thứ tự câu gốc + thứ tự đáp án. Nếu trùng với đề đã tạo, sẽ thử lại với seed khác (tối đa N * 50 lần thử).

---

## Các Trường Hợp Đặc Biệt

### Câu hỏi gắn đoạn văn đọc hiểu (Reading Comprehension)

Khi phát hiện context paragraph chứa pattern "questions X-Y" hoặc "câu X đến Y", chương trình sẽ tạo QuestionBlock gồm context + các câu X->Y. Khi trộn, cả block di chuyển cùng nhau như 1 đơn vị.

### Đáp án trên cùng 1 dòng

Chương trình hỗ trợ 3 kiểu bố trí đáp án:
- Mỗi đáp án 1 paragraph (dòng riêng).
- Tất cả 4 đáp án trên 1 paragraph (cùng dòng).
- 2 đáp án/dòng, 2 dòng liên tiếp.

Khi trộn đáp án inline, chương trình xử lý lại các runs trong paragraph để đổi vị trí nội dung + cập nhật formatting.

### Câu không có đáp án đúng

Nếu không phát hiện formatting đỏ/gạch chân ở bất kỳ đáp án nào, chương trình sẽ cảnh báo và cho phép user nhập thủ công đáp án đúng qua console (cú pháp: `12A 35C`).

### Hình ảnh trong câu hỏi

Hình ảnh nằm trong paragraph element. Khi deep copy XML, hình ảnh sẽ được copy theo vì nó là element con (`<w:drawing>`).

---

## Định Dạng Đầu Ra

### File DOCX đề trộn

- Tên file: `<tên_đề_gốc>_Ma001.docx`, `<tên_đề_gốc>_Ma002.docx`, ...
- Giữ nguyên 100% formatting gốc (fonts, styles, page setup, hình ảnh).
- Số câu hỏi được đánh lại 1->N.
- Formatting đáp án đúng (màu đỏ/gạch chân) được cập nhật theo vị trí mới.

### File Excel bảng đáp án

- Tên file: `<tên_đề_gốc>_DapAn.xlsx`
- Sheet 1 "Đáp án": Bảng tổng hợp đáp án đúng của đề gốc và tất cả đề trộn.
- Sheet 2 "Chi tiết mapping": Bảng chi tiết mapping câu mới <-> câu gốc, đáp án gốc <-> đáp án mới, trạng thái trộn đáp án.

---

## Ghi Chú

- File DOCX đầu vào không cần tag hay định dạng đặc biệt. Câu hỏi chỉ cần đánh số liên tục (1, 2, 3, ...) và đáp án đúng được tô màu đỏ hoặc gạch chân.
- Chương trình chỉ hỗ trợ đề thi trắc nghiệm 4 đáp án A/B/C/D.
- Nên tạo virtual environment riêng để tránh xung đột thư viện.
- Thư mục `output/` sẽ tự động được tạo nếu chưa tồn tại.
