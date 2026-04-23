# -*- coding: utf-8 -*-
"""
Console UI — Hiển thị câu hỏi + nhận input phân nhóm từ user.

Chạy debug:
    python console_ui.py input/de_goc.docx
"""

import os
import sys
import re
from typing import List, Optional, Tuple

from models import ExamDocument, Question, GroupConfig

# Đảm bảo encoding UTF-8
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# Thử import colorama cho console đẹp hơn
try:
    from colorama import init, Fore, Style
    init(autoreset=True)
    HAS_COLOR = True
except ImportError:
    HAS_COLOR = False
    # Fallback: tạo dummy objects
    class _Dummy:
        def __getattr__(self, name):
            return ''
    Fore = _Dummy()
    Style = _Dummy()


# ============================================================
# Display Functions
# ============================================================

def display_exam_summary(exam: ExamDocument):
    """Hiển thị tóm tắt đề thi."""
    correct_count = sum(1 for q in exam.questions if q.correct_answer)
    no_correct = [q.original_number for q in exam.questions if not q.correct_answer]

    print()
    print(f"{Fore.CYAN}╔══════════════════════════════════════════════════════════╗")
    print(f"║  📄 ĐỀ GỐC: {os.path.basename(exam.filepath):<43}║")
    print(f"║  📊 Tổng số câu hỏi: {exam.total_questions:<35}║")
    print(f"║  ✅ Đáp án đúng phát hiện: {correct_count}/{exam.total_questions:<29}║")
    if exam.blocks:
        print(f"║  📖 Blocks đọc hiểu: {len(exam.blocks):<35}║")
    print(f"╚══════════════════════════════════════════════════════════╝{Style.RESET_ALL}")

    if no_correct:
        print(f"\n{Fore.YELLOW}  ⚠️  Câu chưa phát hiện đáp án đúng: {no_correct}{Style.RESET_ALL}")


def display_questions_list(exam: ExamDocument):
    """Hiển thị danh sách câu hỏi dạng bảng ngắn gọn."""
    print(f"\n{Fore.WHITE}--- Danh sách câu hỏi ---{Style.RESET_ALL}")

    for q in exam.questions:
        # Tạo hiển thị đáp án
        opts_display = ""
        for letter in 'ABCD':
            opt = q.options.get(letter)
            if opt and opt.is_correct:
                opts_display += f"{Fore.RED}[{letter}]{Style.RESET_ALL} "
            elif opt:
                opts_display += f" {letter}  "
            else:
                opts_display += " ?  "

        # Cắt stem cho vừa dòng
        stem_short = q.stem_text[:50]
        if len(q.stem_text) > 50:
            stem_short += "..."

        # Block info
        block_info = f"  [block {q.block_id}]" if q.block_id is not None else ""

        # Đáp án đúng
        if q.correct_answer:
            ans_info = f"← Đáp án: {Fore.GREEN}{q.correct_answer}{Style.RESET_ALL}"
        else:
            ans_info = f"← {Fore.YELLOW}⚠️ Chưa xác định{Style.RESET_ALL}"

        print(f"  Câu {q.original_number:>3}: {stem_short:<55} {opts_display} {ans_info}{block_info}")


def display_questions_detail(exam: ExamDocument, question_numbers: List[int] = None):
    """Hiển thị chi tiết các câu hỏi (cho debug)."""
    questions = exam.questions
    if question_numbers:
        questions = [q for q in exam.questions if q.original_number in question_numbers]

    for q in questions:
        print(f"\n{Fore.CYAN}--- Câu {q.original_number} ---{Style.RESET_ALL}")
        print(f"  {q.stem_text}")
        for letter in 'ABCD':
            opt = q.options.get(letter)
            if opt:
                if opt.is_correct:
                    print(f"  {Fore.RED}→ {letter}. {opt.text}{Style.RESET_ALL}")
                else:
                    print(f"    {letter}. {opt.text}")


def display_sections(sections: List[dict]):
    """Hiển thị sections phát hiện tự động."""
    if sections:
        print(f"\n{Fore.CYAN}--- Sections phát hiện tự động ---{Style.RESET_ALL}")
        for s in sections:
            print(f"  [Gợi ý] {s['text']}")
    else:
        print(f"\n{Fore.YELLOW}  ℹ️  Không phát hiện section headers (câu hỏi rải liên tục){Style.RESET_ALL}")


# ============================================================
# Input Functions
# ============================================================

GROUP_INPUT_PATTERN = re.compile(
    r'^(#?)g([0-3])\s*:\s*(.+)$',
    re.IGNORECASE
)

RANGE_PATTERN = re.compile(r'(\d+)\s*[-–]\s*(\d+)')
SINGLE_PATTERN = re.compile(r'^(\d+)$')


def parse_range_string(range_str: str) -> List[Tuple[int, int]]:
    """
    Parse chuỗi dải câu hỏi.
    Ví dụ: "1-10, 15, 20-25" → [(1,10), (15,15), (20,25)]
    """
    ranges = []
    parts = range_str.split(',')
    for part in parts:
        part = part.strip()
        m_range = RANGE_PATTERN.match(part)
        m_single = SINGLE_PATTERN.match(part)
        if m_range:
            start = int(m_range.group(1))
            end = int(m_range.group(2))
            ranges.append((start, end))
        elif m_single:
            num = int(m_single.group(1))
            ranges.append((num, num))
        else:
            print(f"  {Fore.YELLOW}⚠️  Bỏ qua phần không hợp lệ: '{part}'{Style.RESET_ALL}")
    return ranges


def get_group_input(total_questions: int) -> List[GroupConfig]:
    """Nhận input phân nhóm từ user qua console."""
    print(f"\n{Fore.WHITE}{'═' * 58}")
    print(f"📝 Nhập phân nhóm câu hỏi")
    print(f"{'═' * 58}{Style.RESET_ALL}")
    print(f"""
{Fore.CYAN}Cú pháp:{Style.RESET_ALL}  <nhóm>: <dải câu>
{Fore.CYAN}Ví dụ:{Style.RESET_ALL}
  g0: 1-10           ← Không trộn gì (nghe)
  g1: 11-20          ← Chỉ trộn thứ tự câu
  g2: 21-30          ← Chỉ trộn đáp án
  g3: 31-50          ← Trộn cả câu lẫn đáp án
  #g3: 11-20         ← Trộn cả hai + cố định vị trí nhóm

{Fore.YELLOW}Nhấn Enter trống = tất cả câu là g3 (trộn hết){Style.RESET_ALL}
""")

    groups = []

    while True:
        try:
            line = input(f"{Fore.GREEN}> {Style.RESET_ALL}").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            break

        if not line:
            if not groups:
                # Default: tất cả g3
                groups.append(GroupConfig(
                    group_type=3,
                    fixed=False,
                    question_ranges=[(1, total_questions)]
                ))
                print(f"  → Mặc định: g3: 1-{total_questions} (trộn tất cả)")
            break

        m = GROUP_INPUT_PATTERN.match(line)
        if not m:
            print(f"  {Fore.RED}❌ Cú pháp sai! Ví dụ: g3: 1-50{Style.RESET_ALL}")
            continue

        fixed = m.group(1) == '#'
        group_type = int(m.group(2))
        range_str = m.group(3)
        ranges = parse_range_string(range_str)

        if not ranges:
            print(f"  {Fore.RED}❌ Dải câu không hợp lệ!{Style.RESET_ALL}")
            continue

        group = GroupConfig(
            group_type=group_type,
            fixed=fixed,
            question_ranges=ranges
        )

        # Validate ranges
        valid = True
        for start, end in ranges:
            if start < 1 or end > total_questions or start > end:
                print(f"  {Fore.RED}❌ Dải {start}-{end} ngoài phạm vi (1-{total_questions})!{Style.RESET_ALL}")
                valid = False
                break

        if valid:
            groups.append(group)
            print(f"  ✅ {group}")

    return groups


def validate_groups(groups: List[GroupConfig], total_questions: int) -> bool:
    """Validate phân nhóm: kiểm tra overlap, coverage."""
    all_nums = set()
    errors = False

    for group in groups:
        for num in group.question_numbers:
            if num in all_nums:
                print(f"  {Fore.RED}❌ Câu {num} bị trùng nhóm!{Style.RESET_ALL}")
                errors = True
            all_nums.add(num)

    # Kiểm tra coverage
    expected = set(range(1, total_questions + 1))
    missing = expected - all_nums
    extra = all_nums - expected

    if missing:
        print(f"  {Fore.YELLOW}⚠️  Câu chưa gán nhóm: {sorted(missing)} → tự động gán g3{Style.RESET_ALL}")
        # Tự động thêm vào g3
        missing_ranges = _numbers_to_ranges(sorted(missing))
        groups.append(GroupConfig(group_type=3, fixed=False, question_ranges=missing_ranges))

    if extra:
        print(f"  {Fore.RED}❌ Câu ngoài phạm vi: {sorted(extra)}{Style.RESET_ALL}")
        errors = True

    return not errors


def _numbers_to_ranges(nums: List[int]) -> List[Tuple[int, int]]:
    """Chuyển danh sách số thành dải. [1,2,3,5,6,8] → [(1,3), (5,6), (8,8)]"""
    if not nums:
        return []
    ranges = []
    start = nums[0]
    prev = nums[0]
    for n in nums[1:]:
        if n == prev + 1:
            prev = n
        else:
            ranges.append((start, prev))
            start = n
            prev = n
    ranges.append((start, prev))
    return ranges


def get_num_exams() -> int:
    """Nhận số lượng đề trộn cần tạo."""
    while True:
        try:
            line = input(f"\n{Fore.GREEN}📝 Số lượng đề trộn cần tạo: {Style.RESET_ALL}").strip()
            if not line:
                print("  → Mặc định: 4 đề")
                return 4
            num = int(line)
            if num < 1:
                print(f"  {Fore.RED}❌ Phải ít nhất 1!{Style.RESET_ALL}")
                continue
            if num > 50:
                print(f"  {Fore.YELLOW}⚠️  Số lượng lớn ({num}), có thể trùng đề!{Style.RESET_ALL}")
            return num
        except ValueError:
            print(f"  {Fore.RED}❌ Nhập số!{Style.RESET_ALL}")
        except (EOFError, KeyboardInterrupt):
            print("\n  → Mặc định: 4 đề")
            return 4


def get_manual_answers(exam: ExamDocument) -> dict:
    """
    Cho phép user nhập thủ công đáp án đúng cho các câu chưa phát hiện.
    Returns: {q_num: 'A'/'B'/'C'/'D'}
    """
    no_correct = [q.original_number for q in exam.questions if not q.correct_answer]
    if not no_correct:
        return {}

    print(f"\n{Fore.YELLOW}⚠️  {len(no_correct)} câu chưa có đáp án đúng: {no_correct}{Style.RESET_ALL}")
    print(f"Nhập đáp án (VD: 12A 35C) hoặc Enter để bỏ qua:")

    try:
        line = input(f"{Fore.GREEN}> {Style.RESET_ALL}").strip()
    except (EOFError, KeyboardInterrupt):
        return {}

    if not line:
        return {}

    answers = {}
    parts = line.split()
    for part in parts:
        m = re.match(r'(\d+)([A-Da-d])', part)
        if m:
            q_num = int(m.group(1))
            letter = m.group(2).upper()
            answers[q_num] = letter
            print(f"  ✅ Câu {q_num}: đáp án {letter}")
        else:
            print(f"  {Fore.YELLOW}⚠️  Bỏ qua: '{part}'{Style.RESET_ALL}")

    return answers


def apply_manual_answers(exam: ExamDocument, answers: dict):
    """Áp dụng đáp án thủ công vào exam."""
    for q_num, letter in answers.items():
        q = exam.get_question(q_num)
        if q and letter in q.options:
            # Reset tất cả
            for opt in q.options.values():
                opt.is_correct = False
            # Set đáp án đúng
            q.options[letter].is_correct = True
            q.correct_answer = letter
            q.original_correct_answer = letter


# ============================================================
# Debug: chạy python console_ui.py <file.docx>
# ============================================================
if __name__ == "__main__":
    print("=" * 60)
    print("  DEBUG: console_ui.py — Hiển thị + Input")
    print("=" * 60)

    if len(sys.argv) < 2:
        print("\nCách dùng: python console_ui.py <file.docx>")
        print("Ví dụ:     python console_ui.py input/de_goc.docx")

        # Demo với dữ liệu mẫu
        print("\n--- Demo với dữ liệu mẫu ---")
        from models import OptionData
        exam = ExamDocument(filepath="demo_exam.docx", total_questions=5)
        for i in range(1, 6):
            q = Question(
                original_number=i,
                stem_text=f"This is sample question number {i} about English grammar.",
                correct_answer='B' if i % 2 == 0 else 'A',
            )
            for j, letter in enumerate('ABCD'):
                q.options[letter] = OptionData(
                    original_letter=letter,
                    current_letter=letter,
                    text=f"Option {letter} for question {i}",
                    is_correct=(letter == q.correct_answer),
                )
            exam.questions.append(q)

        display_exam_summary(exam)
        display_questions_list(exam)
        display_sections([])

        # Test input
        print("\n--- Test nhập nhóm ---")
        groups = get_group_input(5)
        print(f"\nGroups nhập được: {groups}")
        validate_groups(groups, 5)

        num = get_num_exams()
        print(f"Số đề: {num}")

        print("\n✅ console_ui.py OK")
        sys.exit(0)

    # Dùng file DOCX thật
    from parser import parse_docx, DocxParser

    filepath = sys.argv[1]
    if not os.path.exists(filepath):
        print(f"\n❌ File không tồn tại: {filepath}")
        sys.exit(1)

    exam = parse_docx(filepath)

    # Detect sections
    parser = DocxParser(filepath)
    parser._collect_paragraphs()
    sections = parser.detect_sections()

    # Hiển thị
    display_exam_summary(exam)
    display_questions_list(exam)
    display_sections(sections)

    # Xử lý câu chưa có đáp án
    manual = get_manual_answers(exam)
    if manual:
        apply_manual_answers(exam, manual)

    # Nhập nhóm
    groups = get_group_input(exam.total_questions)
    validate_groups(groups, exam.total_questions)

    num = get_num_exams()
    print(f"\nSố đề trộn: {num}")

    # Hiển thị tóm tắt
    print(f"\n--- Tóm tắt phân nhóm ---")
    for g in groups:
        print(f"  {g}")

    print("\n✅ console_ui.py OK")
