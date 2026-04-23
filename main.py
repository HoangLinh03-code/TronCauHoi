# -*- coding: utf-8 -*-
"""
Main — Entry point cho chương trình Trộn Câu Hỏi.

Chạy:
    python main.py input/de_goc.docx
    python main.py input/de_goc.docx --output output/
"""

import os
import sys
import argparse

# Đảm bảo encoding UTF-8
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

from parser import parse_docx, DocxParser
from console_ui import (
    display_exam_summary,
    display_questions_list,
    display_questions_detail,
    display_sections,
    get_group_input,
    get_num_exams,
    get_manual_answers,
    apply_manual_answers,
    validate_groups,
)
from shuffler import generate_unique_exams, print_exam_comparison
from writer import write_all_outputs


def main():
    """Pipeline chính."""
    # ═══ Parse arguments ═══
    arg_parser = argparse.ArgumentParser(
        description="Trộn Câu Hỏi — Trộn đề thi từ file DOCX",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ví dụ:
  python main.py de_goc.docx
  python main.py de_goc.docx --output output/
  python main.py de_goc.docx --output output/ --detail 1-5
        """
    )
    arg_parser.add_argument("input", help="File DOCX đề gốc")
    arg_parser.add_argument("--output", "-o", default=None,
                           help="Thư mục output (mặc định: output/ cùng thư mục input)")
    arg_parser.add_argument("--detail", "-d", default=None,
                           help="Hiển thị chi tiết câu hỏi (VD: 1-5, 10, 15-20)")
    args = arg_parser.parse_args()

    # Validate input
    if not os.path.exists(args.input):
        print(f"❌ File không tồn tại: {args.input}")
        sys.exit(1)

    if not args.input.lower().endswith('.docx'):
        print(f"❌ File phải là .docx: {args.input}")
        sys.exit(1)

    # Output dir
    if args.output:
        output_dir = args.output
    else:
        output_dir = os.path.join(os.path.dirname(args.input) or '.', "output")

    print()
    print("╔══════════════════════════════════════════════════════════╗")
    print("║           🔀 CHƯƠNG TRÌNH TRỘN CÂU HỎI                 ║")
    print("║                    v1.0                                  ║")
    print("╚══════════════════════════════════════════════════════════╝")

    # ═══ Bước 1: Parse DOCX ═══
    print(f"\n{'─' * 58}")
    print(f"  BƯỚC 1: Parse đề gốc")
    print(f"{'─' * 58}")

    exam = parse_docx(args.input)

    if exam.total_questions == 0:
        print("❌ Không tìm thấy câu hỏi nào trong file!")
        sys.exit(1)

    # ═══ Bước 2: Auto-detect sections ═══
    parser = DocxParser(args.input)
    parser._collect_paragraphs()
    sections = parser.detect_sections()

    # ═══ Bước 3: Hiển thị ═══
    print(f"\n{'─' * 58}")
    print(f"  BƯỚC 2: Kiểm tra kết quả parse")
    print(f"{'─' * 58}")

    display_exam_summary(exam)
    display_questions_list(exam)
    display_sections(sections)

    # Hiển thị chi tiết nếu yêu cầu
    if args.detail:
        detail_nums = _parse_detail_range(args.detail)
        if detail_nums:
            display_questions_detail(exam, detail_nums)

    # ═══ Bước 4: Xử lý câu chưa có đáp án ═══
    no_correct = [q.original_number for q in exam.questions if not q.correct_answer]
    if no_correct:
        manual = get_manual_answers(exam)
        if manual:
            apply_manual_answers(exam, manual)

    # ═══ Bước 5: Nhập phân nhóm ═══
    print(f"\n{'─' * 58}")
    print(f"  BƯỚC 3: Phân nhóm câu hỏi")
    print(f"{'─' * 58}")

    groups = get_group_input(exam.total_questions)
    if not validate_groups(groups, exam.total_questions):
        print("❌ Phân nhóm không hợp lệ!")
        sys.exit(1)

    print(f"\n  Tóm tắt nhóm:")
    for g in groups:
        type_desc = {0: "Không trộn", 1: "Trộn câu hỏi", 2: "Trộn đáp án", 3: "Trộn cả hai"}
        fixed_desc = " (cố định vị trí)" if g.fixed else ""
        print(f"    {g} → {type_desc[g.group_type]}{fixed_desc}")

    # ═══ Bước 6: Số lượng đề ═══
    num_exams = get_num_exams()

    # ═══ Bước 7: Trộn ═══
    print(f"\n{'─' * 58}")
    print(f"  BƯỚC 4: Trộn đề")
    print(f"{'─' * 58}")

    shuffled_exams = generate_unique_exams(exam, groups, num_exams)

    if not shuffled_exams:
        print("❌ Không tạo được đề trộn nào!")
        sys.exit(1)

    # Hiển thị so sánh (tóm tắt)
    for i, sh_exam in enumerate(shuffled_exams):
        print(f"\n  📋 Đề Mã {i + 1:03d}:")
        # In 5 câu đầu
        for q in sh_exam.questions[:5]:
            ans_info = f"({q.original_correct_answer}→{q.correct_answer})" if q.options_were_shuffled else f"({q.correct_answer})"
            print(f"    Câu {q.new_number:>2} ← gốc {q.original_number:>2} {ans_info}")
        if len(sh_exam.questions) > 5:
            print(f"    ... ({len(sh_exam.questions) - 5} câu nữa)")

    # ═══ Bước 8: Xuất files ═══
    print(f"\n{'─' * 58}")
    print(f"  BƯỚC 5: Xuất file")
    print(f"{'─' * 58}")

    write_all_outputs(exam, shuffled_exams, args.input, output_dir)

    print(f"\n{'═' * 58}")
    print(f"  🎉 HOÀN TẤT!")
    print(f"  📁 Output: {os.path.abspath(output_dir)}")
    print(f"  📄 {len(shuffled_exams)} đề trộn + 1 file đáp án")
    print(f"{'═' * 58}")


def _parse_detail_range(range_str: str):
    """Parse range string cho --detail."""
    import re
    nums = []
    parts = range_str.split(',')
    for part in parts:
        part = part.strip()
        m = re.match(r'(\d+)\s*[-–]\s*(\d+)', part)
        if m:
            nums.extend(range(int(m.group(1)), int(m.group(2)) + 1))
        elif part.isdigit():
            nums.append(int(part))
    return nums


if __name__ == "__main__":
    main()
