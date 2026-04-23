# -*- coding: utf-8 -*-
"""
Shuffler — Thuật toán trộn đề thi (Fisher-Yates shuffle).

Chạy debug:
    python shuffler.py input/de_goc.docx "g3: 1-50" --num 4
    python shuffler.py --demo
"""

import sys
import random
from copy import deepcopy
from typing import List, Set, Tuple, Optional

from models import ExamDocument, Question, OptionData, GroupConfig, QuestionBlock

# Đảm bảo encoding UTF-8
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')


# ============================================================
# Core Shuffle Functions
# ============================================================

def shuffle_options(question: Question, rng: random.Random):
    """
    Hoán vị đáp án A/B/C/D của 1 câu hỏi.
    Sử dụng Fisher-Yates shuffle (qua rng.shuffle).
    Cập nhật correct_answer và option_mapping.
    """
    letters = ['A', 'B', 'C', 'D']
    # Lấy danh sách OptionData theo thứ tự A, B, C, D
    option_values = [question.options[l] for l in letters if l in question.options]

    if len(option_values) < 4:
        return  # Ko đủ đáp án, bỏ qua

    # Fisher-Yates shuffle
    rng.shuffle(option_values)

    # Gán lại letter mới + lưu mapping
    question.option_mapping = {}
    for i, letter in enumerate(letters):
        opt = option_values[i]
        old_letter = opt.original_letter
        opt.current_letter = letter
        question.options[letter] = opt
        question.option_mapping[letter] = old_letter  # mới → gốc

    # Cập nhật đáp án đúng
    question.correct_answer = question.get_correct_letter() or ""
    question.options_were_shuffled = True


def shuffle_questions_in_group(questions: List[Question], rng: random.Random,
                                blocks: List[QuestionBlock] = None):
    """
    Hoán vị thứ tự câu hỏi trong 1 nhóm.
    Câu hỏi thuộc cùng block (đọc hiểu) sẽ di chuyển cùng nhau.

    Args:
        questions: Danh sách câu hỏi CẦN trộn (sẽ bị modified in-place)
        rng: Random generator
        blocks: Danh sách blocks đọc hiểu (nếu có)
    """
    if not questions:
        return

    if not blocks:
        # Không có block → shuffle bình thường
        rng.shuffle(questions)
        return

    # Có blocks → cần giữ câu hỏi trong block đi cùng nhau
    # Tạo danh sách "shuffle units": mỗi unit là 1 câu hoặc 1 block
    units = []  # List of (type, data): ('single', Question) hoặc ('block', [Question])
    in_block = set()

    # Tìm block chứa câu hỏi thuộc nhóm này
    q_nums = {q.original_number for q in questions}
    relevant_blocks = []
    if blocks:
        for block in blocks:
            block_q_in_group = [n for n in block.question_numbers if n in q_nums]
            if block_q_in_group:
                relevant_blocks.append(block)
                in_block.update(block_q_in_group)

    # Tạo units
    for block in relevant_blocks:
        block_questions = [q for q in questions if q.original_number in block.question_numbers]
        if block_questions:
            units.append(('block', block_questions))

    for q in questions:
        if q.original_number not in in_block:
            units.append(('single', [q]))

    # Shuffle units
    rng.shuffle(units)

    # Flatten lại
    result = []
    for unit_type, unit_questions in units:
        result.extend(unit_questions)

    # Cập nhật in-place
    questions[:] = result


def shuffle_exam(exam: ExamDocument, group_configs: List[GroupConfig],
                 seed: int) -> ExamDocument:
    """
    Tạo 1 bản trộn từ đề gốc.

    Args:
        exam: Đề gốc đã parse
        group_configs: Phân nhóm do user chỉ định
        seed: Random seed cho reproducibility

    Returns:
        ExamDocument mới đã trộn
    """
    rng = random.Random(seed)
    shuffled = deepcopy(exam)

    # Tạo mapping: original_number → group_index để theo dõi nhóm
    num_to_group = {}
    for gi, group in enumerate(group_configs):
        for n in group.question_numbers:
            num_to_group[n] = gi

    # ═══ BƯỚC A: Trộn bên trong mỗi nhóm ═══
    # Lưu lại kết quả trộn cho mỗi nhóm (thứ tự mới)
    group_shuffled_questions = {}  # {group_index: [Question, ...]}

    for gi, group in enumerate(group_configs):
        q_nums = group.question_numbers
        questions = [shuffled.get_question(n) for n in q_nums]
        questions = [q for q in questions if q is not None]

        if not questions:
            group_shuffled_questions[gi] = []
            continue

        if group.group_type == 0:
            # g0: Không trộn gì
            group_shuffled_questions[gi] = questions

        elif group.group_type == 1:
            # g1: Chỉ trộn thứ tự câu hỏi
            shuffle_questions_in_group(questions, rng, shuffled.blocks)
            group_shuffled_questions[gi] = questions

        elif group.group_type == 2:
            # g2: Chỉ trộn đáp án
            for q in questions:
                shuffle_options(q, rng)
            group_shuffled_questions[gi] = questions

        elif group.group_type == 3:
            # g3: Trộn cả hai
            shuffle_questions_in_group(questions, rng, shuffled.blocks)
            for q in questions:
                shuffle_options(q, rng)
            group_shuffled_questions[gi] = questions

    # ═══ BƯỚC B: Trộn vị trí giữa các nhóm (nếu có nhóm không cố định) ═══
    group_order = list(range(len(group_configs)))
    movable_indices = [i for i in group_order if not group_configs[i].fixed]
    if len(movable_indices) > 1:
        movable_order = [group_order[i] for i in movable_indices]
        rng.shuffle(movable_order)
        for j, idx in enumerate(movable_indices):
            group_order[idx] = movable_order[j]

    # ═══ BƯỚC C: Ghép câu hỏi theo thứ tự nhóm mới ═══
    final_questions = []
    for gi in group_order:
        questions = group_shuffled_questions.get(gi, [])
        final_questions.extend(questions)

    # ═══ BƯỚC D: Đánh số lại 1→N ═══
    for i, q in enumerate(final_questions):
        q.new_number = i + 1

    shuffled.questions = final_questions
    shuffled.total_questions = len(final_questions)

    return shuffled


# ============================================================
# Unique Exam Generation
# ============================================================

def compute_signature(exam: ExamDocument) -> tuple:
    """
    Tạo fingerprint của 1 đề trộn.
    Dựa trên thứ tự câu gốc + thứ tự đáp án.
    """
    parts = []
    for q in exam.questions:
        # Thứ tự câu gốc
        q_id = q.original_number
        # Thứ tự đáp án (nếu bị trộn)
        if q.options_were_shuffled:
            opt_order = tuple(
                q.options[l].original_letter for l in 'ABCD'
                if l in q.options
            )
        else:
            opt_order = ('A', 'B', 'C', 'D')
        parts.append((q_id, opt_order))
    return tuple(parts)


def generate_unique_exams(exam: ExamDocument, group_configs: List[GroupConfig],
                          num_exams: int) -> List[ExamDocument]:
    """
    Sinh N đề trộn, đảm bảo không trùng nhau.

    Args:
        exam: Đề gốc
        group_configs: Phân nhóm
        num_exams: Số đề cần tạo

    Returns:
        List[ExamDocument]: Danh sách đề trộn
    """
    results = []
    seen: Set[tuple] = set()
    max_attempts = num_exams * 50  # Tránh infinite loop

    print(f"\n🔄 Đang trộn {num_exams} đề...")

    for attempt in range(max_attempts):
        if len(results) >= num_exams:
            break

        seed = random.randint(0, 2**32)
        # Deep copy group_configs vì shuffle_exam có thể thay đổi thứ tự
        gc_copy = deepcopy(group_configs)
        shuffled = shuffle_exam(exam, gc_copy, seed)

        sig = compute_signature(shuffled)

        if sig not in seen:
            seen.add(sig)
            results.append(shuffled)
            print(f"  ✅ Đề {len(results)}/{num_exams} (seed={seed})")

    if len(results) < num_exams:
        print(f"  ⚠️  Chỉ tạo được {len(results)}/{num_exams} đề khác nhau "
              f"(sau {max_attempts} lần thử)")

    return results


# ============================================================
# Debug helpers
# ============================================================

def print_exam_comparison(original: ExamDocument, shuffled: ExamDocument):
    """In so sánh đề gốc vs đề trộn."""
    print("\n--- So sánh Gốc vs Trộn ---")
    print(f"{'Câu mới':>8} | {'Câu gốc':>8} | {'Đáp án gốc':>10} | {'Đáp án mới':>10} | Câu hỏi")
    print("-" * 80)

    for q in shuffled.questions:
        orig_ans = q.original_correct_answer
        new_ans = q.correct_answer

        # Tình trạng đáp án
        if q.options_were_shuffled:
            ans_status = f"{orig_ans} → {new_ans}"
        else:
            ans_status = f"{orig_ans} (giữ nguyên)"

        stem_short = q.stem_text[:35]
        if len(q.stem_text) > 35:
            stem_short += "..."

        print(f"  {q.new_number:>5} | {q.original_number:>8} | {orig_ans:>10} | {new_ans:>10} | {stem_short}")


# ============================================================
# Debug: chạy python shuffler.py
# ============================================================
if __name__ == "__main__":
    print("=" * 60)
    print("  DEBUG: shuffler.py — Trộn đề thi")
    print("=" * 60)

    # Demo với dữ liệu mẫu
    if len(sys.argv) < 2 or sys.argv[1] == '--demo':
        print("\n--- Demo trộn dữ liệu mẫu ---")

        # Tạo exam mẫu
        exam = ExamDocument(filepath="demo.docx", total_questions=10)
        for i in range(1, 11):
            q = Question(
                original_number=i,
                stem_text=f"Question {i}: What is the answer?",
                correct_answer='A',
                original_correct_answer='A',
            )
            for j, letter in enumerate('ABCD'):
                q.options[letter] = OptionData(
                    original_letter=letter,
                    current_letter=letter,
                    text=f"Option {letter}{i}",
                    is_correct=(letter == 'A'),
                )
            exam.questions.append(q)

        # Config: g0 cho 1-3, g3 cho 4-10
        configs = [
            GroupConfig(group_type=0, fixed=True, question_ranges=[(1, 3)]),
            GroupConfig(group_type=3, fixed=False, question_ranges=[(4, 10)]),
        ]

        print(f"\nGroups: {configs}")
        print(f"Exam: {exam.total_questions} câu")

        # Trộn 3 đề
        exams = generate_unique_exams(exam, configs, 3)

        for idx, sh in enumerate(exams):
            print(f"\n{'=' * 40}")
            print(f"  ĐỀ TRỘN {idx + 1}")
            print(f"{'=' * 40}")
            print_exam_comparison(exam, sh)

            # Verify g0 questions unchanged
            print(f"\n  Kiểm tra g0 (câu 1-3 giữ nguyên):")
            for q in sh.questions[:3]:
                unchanged = (q.original_number == q.new_number and
                             not q.options_were_shuffled)
                status = "✅ OK" if unchanged else "❌ SAI"
                print(f"    Câu {q.new_number}: gốc={q.original_number}, "
                      f"đổi đáp án={q.options_were_shuffled} → {status}")

        # Verify uniqueness
        sigs = [compute_signature(sh) for sh in exams]
        unique = len(set(sigs))
        print(f"\n📊 Unique đề: {unique}/{len(exams)}")
        if unique == len(exams):
            print("✅ Tất cả các đề đều khác nhau!")
        else:
            print("❌ Có đề trùng nhau!")

        print("\n✅ shuffler.py OK")
        sys.exit(0)

    # Dùng file DOCX thật
    from parser import parse_docx
    import re

    filepath = sys.argv[1]
    group_str = sys.argv[2] if len(sys.argv) > 2 else "g3: 1-50"
    num_exams = 4

    for arg in sys.argv:
        if arg.startswith('--num'):
            num_exams = int(arg.split('=')[-1]) if '=' in arg else int(sys.argv[sys.argv.index(arg) + 1])

    print(f"\nFile: {filepath}")
    print(f"Groups: {group_str}")
    print(f"Num exams: {num_exams}")

    import os
    if not os.path.exists(filepath):
        print(f"❌ File không tồn tại: {filepath}")
        sys.exit(1)

    exam = parse_docx(filepath)

    # Parse group string
    from console_ui import parse_range_string
    m = re.match(r'(#?)g([0-3]):\s*(.+)', group_str, re.IGNORECASE)
    if m:
        configs = [GroupConfig(
            group_type=int(m.group(2)),
            fixed=m.group(1) == '#',
            question_ranges=parse_range_string(m.group(3))
        )]
    else:
        configs = [GroupConfig(group_type=3, question_ranges=[(1, exam.total_questions)])]

    exams = generate_unique_exams(exam, configs, num_exams)

    for idx, sh in enumerate(exams):
        print(f"\n--- ĐỀ {idx + 1} ---")
        print_exam_comparison(exam, sh)

    print("\n✅ shuffler.py OK")
