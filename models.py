# -*- coding: utf-8 -*-
"""
Models — Cấu trúc dữ liệu cho chương trình Trộn Câu Hỏi.
Chạy: python models.py   → In cấu trúc mẫu để kiểm tra.
"""

import sys
from copy import deepcopy
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple


@dataclass
class OptionData:
    """Dữ liệu 1 đáp án (A/B/C/D)."""
    original_letter: str                # Chữ cái gốc ('A', 'B', 'C', 'D')
    current_letter: str                 # Chữ cái hiện tại (sau trộn)
    text: str                           # Nội dung text thuần
    is_correct: bool = False            # Đáp án đúng?
    elements: list = field(default_factory=list)  # lxml paragraph elements gốc

    def __repr__(self):
        mark = "✓" if self.is_correct else " "
        return f"[{mark}] {self.current_letter}. {self.text[:60]}"


@dataclass
class Question:
    """Dữ liệu 1 câu hỏi."""
    original_number: int                     # Số thứ tự gốc
    new_number: int = 0                      # Số thứ tự sau trộn (0 = chưa gán)
    stem_text: str = ""                      # Nội dung câu hỏi (text thuần)
    correct_answer: str = ""                 # Đáp án đúng hiện tại ('A'/'B'/'C'/'D')
    original_correct_answer: str = ""        # Đáp án đúng gốc
    options: Dict[str, OptionData] = field(default_factory=dict)  # {'A': OptionData, ...}
    stem_elements: list = field(default_factory=list)    # lxml elements thân câu hỏi
    context_elements: list = field(default_factory=list) # lxml elements ngữ cảnh (đoạn văn đọc hiểu)
    context_text: str = ""                   # Text ngữ cảnh (để hiển thị)
    options_were_shuffled: bool = False       # Đáp án có bị trộn không?
    option_mapping: Dict[str, str] = field(default_factory=dict)  # mapping mới→gốc
    block_id: Optional[int] = None           # ID block (nếu thuộc nhóm đọc hiểu)
    is_inline: bool = False                  # Đáp án nằm trên cùng dòng với stem?

    def __repr__(self):
        opts = " | ".join(str(self.options.get(l, "?")) for l in "ABCD")
        return f"Q{self.original_number}: {self.stem_text[:50]}... | {opts}"

    def get_correct_letter(self) -> Optional[str]:
        """Lấy chữ cái đáp án đúng hiện tại."""
        for letter, opt in self.options.items():
            if opt.is_correct:
                return letter
        return None


@dataclass
class QuestionBlock:
    """
    Khối câu hỏi đi kèm ngữ cảnh chung (đoạn văn đọc hiểu).
    Khi trộn, cả block di chuyển cùng nhau.
    """
    block_id: int
    context_elements: list = field(default_factory=list)   # Paragraph elements đoạn văn
    context_text: str = ""
    question_numbers: List[int] = field(default_factory=list)  # Câu thuộc block


@dataclass
class GroupConfig:
    """Cấu hình 1 nhóm câu hỏi (do user nhập)."""
    group_type: int                    # 0, 1, 2, 3
    fixed: bool = False                # Có dấu # (cố định vị trí)?
    question_ranges: List[Tuple[int, int]] = field(default_factory=list)  # [(start, end), ...]

    @property
    def question_numbers(self) -> List[int]:
        """Trả về tất cả số câu thuộc nhóm."""
        nums = []
        for start, end in self.question_ranges:
            nums.extend(range(start, end + 1))
        return nums

    def __repr__(self):
        prefix = "#" if self.fixed else ""
        ranges = ", ".join(f"{s}-{e}" for s, e in self.question_ranges)
        return f"{prefix}g{self.group_type}: {ranges}"


@dataclass
class ExamDocument:
    """Dữ liệu toàn bộ đề thi."""
    filepath: str = ""
    header_elements: list = field(default_factory=list)    # Paragraph elements phần đầu
    header_text: str = ""                                   # Text phần đầu
    questions: List[Question] = field(default_factory=list)
    blocks: List[QuestionBlock] = field(default_factory=list)
    total_questions: int = 0

    def get_question(self, number: int) -> Optional[Question]:
        """Lấy câu hỏi theo số thứ tự gốc."""
        for q in self.questions:
            if q.original_number == number:
                return q
        return None

    def summary(self) -> str:
        """Tóm tắt đề thi."""
        lines = [
            f"📄 File: {self.filepath}",
            f"📊 Tổng câu hỏi: {self.total_questions}",
        ]
        # Đếm câu có đáp án đúng
        correct_count = sum(1 for q in self.questions if q.correct_answer)
        lines.append(f"✅ Đáp án đúng phát hiện: {correct_count}/{self.total_questions}")

        if self.blocks:
            lines.append(f"📖 Số block đọc hiểu: {len(self.blocks)}")

        return "\n".join(lines)


# ============================================================
# Debug: chạy python models.py
# ============================================================
if __name__ == "__main__":
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')

    print("=" * 60)
    print("  DEBUG: Kiểm tra cấu trúc dữ liệu models.py")
    print("=" * 60)

    # Tạo dữ liệu mẫu
    opt_a = OptionData(original_letter='A', current_letter='A', text='Hanoi', is_correct=True)
    opt_b = OptionData(original_letter='B', current_letter='B', text='Ho Chi Minh City')
    opt_c = OptionData(original_letter='C', current_letter='C', text='Da Nang')
    opt_d = OptionData(original_letter='D', current_letter='D', text='Hue')

    q1 = Question(
        original_number=1,
        stem_text="What is the capital of Vietnam?",
        correct_answer='A',
        original_correct_answer='A',
        options={'A': opt_a, 'B': opt_b, 'C': opt_c, 'D': opt_d}
    )

    q2 = Question(
        original_number=2,
        stem_text="She ___ to school every day.",
        correct_answer='B',
        original_correct_answer='B',
        options={
            'A': OptionData('A', 'A', 'go'),
            'B': OptionData('B', 'B', 'goes', is_correct=True),
            'C': OptionData('C', 'C', 'going'),
            'D': OptionData('D', 'D', 'gone'),
        }
    )

    exam = ExamDocument(
        filepath="test_exam.docx",
        questions=[q1, q2],
        total_questions=2,
    )

    g1 = GroupConfig(group_type=0, fixed=True, question_ranges=[(1, 1)])
    g2 = GroupConfig(group_type=3, fixed=False, question_ranges=[(2, 2)])

    print("\n--- ExamDocument ---")
    print(exam.summary())

    print("\n--- Questions ---")
    for q in exam.questions:
        print(f"\n  Câu {q.original_number}: {q.stem_text}")
        print(f"  Đáp án đúng: {q.correct_answer}")
        for letter in 'ABCD':
            opt = q.options.get(letter)
            if opt:
                print(f"    {opt}")

    print("\n--- Groups ---")
    print(f"  {g1}")
    print(f"  {g2}")

    print("\n--- get_correct_letter() ---")
    print(f"  Q1: {q1.get_correct_letter()}")
    print(f"  Q2: {q2.get_correct_letter()}")

    print("\n✅ models.py OK")
