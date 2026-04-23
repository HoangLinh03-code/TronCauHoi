# -*- coding: utf-8 -*-
"""
debug_shuffle.py — Tool debug độc lập cho phần trộn đề.

Mục đích:
  - Kiểm tra logic trộn mà KHÔNG cần writer.py / parser.py phức tạp
  - Cho phép test bằng file DOCX thật hoặc dữ liệu giả
  - In report chi tiết: mapping câu gốc→mới, đáp án gốc→mới
  - Verify tính đúng đắn TRƯỚC KHI chạy main pipeline

Chạy:
    python debug_shuffle.py --demo
    python debug_shuffle.py --demo --groups "g0:1-4, g3:5-20" --num 3
    python debug_shuffle.py de_goc.docx --groups "g0:1-4, g3:5-50" --num 4
    python debug_shuffle.py de_goc.docx --groups "g3:1-50" --num 4 --verbose
    python debug_shuffle.py de_goc.docx --groups "g3:1-50" --num 4 --export-csv
"""

import sys
import os
import re
import random
import csv
import argparse
from copy import deepcopy
from typing import Dict, List, Set, Optional, Tuple

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')


# ════════════════════════════════════════════════════════════════
# Dữ liệu nội bộ (không phụ thuộc models.py)
# ════════════════════════════════════════════════════════════════

class Option:
    """Đáp án đơn giản."""
    __slots__ = ('original_letter', 'current_letter', 'text', 'is_correct')

    def __init__(self, letter: str, text: str, is_correct: bool = False):
        self.original_letter = letter
        self.current_letter  = letter
        self.text            = text
        self.is_correct      = is_correct

    def __repr__(self):
        mark = "★" if self.is_correct else " "
        return f"{mark}{self.current_letter}({self.original_letter}): {self.text[:30]}"


class Question:
    """Câu hỏi đơn giản."""
    def __init__(self, num: int, stem: str, correct: str,
                 options: Dict[str, Option]):
        self.original_number   = num
        self.new_number        = num
        self.stem_text         = stem
        self.correct_answer    = correct
        self.original_correct_answer = correct
        self.options           = options
        self.option_mapping: Dict[str, str] = {l: l for l in 'ABCD'}
        self.options_were_shuffled = False

    def get_correct_letter(self) -> Optional[str]:
        for l, opt in self.options.items():
            if opt.is_correct:
                return l
        return None

    def __repr__(self):
        return (f"Q{self.original_number}→{self.new_number} "
                f"[{self.original_correct_answer}→{self.correct_answer}]")


class Exam:
    """Đề thi đơn giản."""
    def __init__(self):
        self.questions: List[Question] = []
        self.total_questions = 0
        self.blocks = []

    def get_question(self, original_number: int) -> Optional[Question]:
        for q in self.questions:
            if q.original_number == original_number:
                return q
        return None


class GroupConfig:
    """Cấu hình nhóm."""
    def __init__(self, group_type: int, fixed: bool,
                 ranges: List[Tuple[int, int]]):
        self.group_type = group_type
        self.fixed = fixed
        self.question_ranges = ranges
        self._nums: Optional[List[int]] = None

    @property
    def question_numbers(self) -> List[int]:
        if self._nums is None:
            self._nums = []
            for s, e in self.question_ranges:
                self._nums.extend(range(s, e + 1))
        return self._nums

    def __repr__(self):
        prefix = '#' if self.fixed else ''
        rng = ', '.join(f'{s}-{e}' for s, e in self.question_ranges)
        type_names = {0: 'Không trộn', 1: 'Trộn câu',
                      2: 'Trộn ĐA', 3: 'Trộn cả hai'}
        return f"{prefix}g{self.group_type}[{type_names[self.group_type]}]: {rng}"


# ════════════════════════════════════════════════════════════════
# Các hàm shuffle (copy từ shuffler.py, tự chứa)
# ════════════════════════════════════════════════════════════════

def _shuffle_options(q: Question, rng: random.Random) -> None:
    letters = list('ABCD')
    opts = [q.options[l] for l in letters]
    rng.shuffle(opts)
    new_mapping: Dict[str, str] = {}
    for i, letter in enumerate(letters):
        opt = opts[i]
        opt.current_letter = letter
        q.options[letter] = opt
        new_mapping[letter] = opt.original_letter
    q.option_mapping = new_mapping
    new_correct = q.get_correct_letter()
    if new_correct:
        q.correct_answer = new_correct
    q.options_were_shuffled = True


def _shuffle_questions(questions: List[Question], rng: random.Random) -> None:
    rng.shuffle(questions)


def _shuffle_exam(exam: Exam, configs: List[GroupConfig], seed: int) -> Exam:
    rng = random.Random(seed)
    shuffled = deepcopy(exam)

    group_results: Dict[int, List[Question]] = {}
    for gi, group in enumerate(configs):
        qs = [shuffled.get_question(n) for n in group.question_numbers]
        qs = [q for q in qs if q is not None]
        if not qs:
            group_results[gi] = []
            continue
        g = group.group_type
        if g == 1:
            _shuffle_questions(qs, rng)
        elif g == 2:
            for q in qs:
                _shuffle_options(q, rng)
        elif g == 3:
            _shuffle_questions(qs, rng)
            for q in qs:
                _shuffle_options(q, rng)
        group_results[gi] = qs

    # Trộn thứ tự nhóm
    order = list(range(len(configs)))
    movable = [i for i in order if not configs[i].fixed]
    if len(movable) > 1:
        vals = [order[i] for i in movable]
        rng.shuffle(vals)
        for j, idx in enumerate(movable):
            order[idx] = vals[j]

    final = []
    for gi in order:
        final.extend(group_results.get(gi, []))

    for i, q in enumerate(final):
        q.new_number = i + 1

    shuffled.questions = final
    shuffled.total_questions = len(final)
    return shuffled


def _signature(exam: Exam) -> tuple:
    parts = []
    for q in exam.questions:
        if q.options_were_shuffled:
            opt = tuple(q.options[l].original_letter for l in 'ABCD')
        else:
            opt = ('A', 'B', 'C', 'D')
        parts.append((q.original_number, opt))
    return tuple(parts)


def _generate(exam: Exam, configs: List[GroupConfig],
              num: int) -> List[Exam]:
    results = []
    seen: Set[tuple] = set()
    max_try = max(num * 200, 1000)
    print(f"\n🔄 Trộn {num} đề (max {max_try} lần thử)...")
    for _ in range(max_try):
        if len(results) >= num:
            break
        seed = random.randint(0, 2**32 - 1)
        sh = _shuffle_exam(exam, deepcopy(configs), seed)
        sig = _signature(sh)
        if sig not in seen:
            seen.add(sig)
            results.append(sh)
            print(f"  ✅ Đề {len(results):>2}/{num}  seed={seed}")
    if len(results) < num:
        print(f"  ⚠️  Chỉ tạo được {len(results)}/{num} đề khác nhau.")
    return results


# ════════════════════════════════════════════════════════════════
# Verify
# ════════════════════════════════════════════════════════════════

def verify(original: Exam, shuffled: Exam,
           configs: List[GroupConfig]) -> Dict:
    errors: List[str] = []
    warnings: List[str] = []

    # Số câu
    if original.total_questions != shuffled.total_questions:
        errors.append(
            f"Số câu: gốc={original.total_questions}, "
            f"trộn={shuffled.total_questions}"
        )

    # Không trùng câu
    seen_orig: Dict[int, int] = {}
    for q in shuffled.questions:
        if q.original_number in seen_orig:
            errors.append(
                f"Câu gốc {q.original_number} xuất hiện ≥2 lần "
                f"(vị trí mới {seen_orig[q.original_number]} và {q.new_number})"
            )
        seen_orig[q.original_number] = q.new_number

    # Đủ câu
    expected = set(range(1, original.total_questions + 1))
    missing = expected - set(seen_orig.keys())
    extra   = set(seen_orig.keys()) - expected
    if missing:
        errors.append(f"Thiếu câu gốc: {sorted(missing)}")
    if extra:
        errors.append(f"Câu ngoài phạm vi: {sorted(extra)}")

    # Đáp án đúng vẫn đúng
    orig_map = {q.original_number: q for q in original.questions}
    for q in shuffled.questions:
        oq = orig_map.get(q.original_number)
        if not oq:
            continue
        new_opt = q.options.get(q.correct_answer)
        if new_opt is None:
            errors.append(
                f"Q{q.original_number}: correct_answer='{q.correct_answer}' "
                f"không có trong options"
            )
        elif not new_opt.is_correct:
            errors.append(
                f"Q{q.original_number}: correct_answer='{q.correct_answer}' "
                f"nhưng option đó không phải đáp án đúng (is_correct=False)"
            )
        if q.options_were_shuffled:
            mapped = q.option_mapping.get(q.correct_answer)
            if mapped != oq.correct_answer:
                errors.append(
                    f"Q{q.original_number}: option_mapping[{q.correct_answer}]="
                    f"'{mapped}' ≠ đáp án gốc '{oq.correct_answer}'"
                )

    # g0 không đổi đáp án
    for group in configs:
        if group.group_type == 0:
            q_set = set(group.question_numbers)
            for q in shuffled.questions:
                if q.original_number in q_set and q.options_were_shuffled:
                    errors.append(
                        f"g0 Q{q.original_number}: đáp án bị trộn (không được phép)"
                    )

    # Số thứ tự mới liên tục 1→N
    new_nums = [q.new_number for q in shuffled.questions]
    expected_seq = list(range(1, shuffled.total_questions + 1))
    if new_nums != expected_seq:
        errors.append(
            f"Số thứ tự mới KHÔNG liên tục: "
            f"{new_nums[:15]}{'...' if len(new_nums) > 15 else ''}"
        )

    return {
        'ok': len(errors) == 0,
        'errors': errors,
        'warnings': warnings,
    }


# ════════════════════════════════════════════════════════════════
# Hiển thị
# ════════════════════════════════════════════════════════════════

DIVIDER_THIN  = "─" * 75
DIVIDER_THICK = "═" * 75

def print_header(text: str) -> None:
    print(f"\n{DIVIDER_THICK}")
    print(f"  {text}")
    print(DIVIDER_THICK)


def print_section(text: str) -> None:
    print(f"\n{DIVIDER_THIN}")
    print(f"  {text}")
    print(DIVIDER_THIN)


def print_mapping_table(original: Exam, shuffled: Exam,
                        label: str = "", max_rows: int = None) -> None:
    """In bảng mapping câu gốc→mới và đáp án gốc→mới."""
    title = f"📋 {label}" if label else "📋 Bảng mapping"
    print(f"\n  {title}  ({shuffled.total_questions} câu)")
    print(f"  {'Câu mới':>7} | {'Câu gốc':>7} | {'ĐA gốc':>7} | "
          f"{'ĐA mới':>7} | {'Mapping ĐA (A→B=Cũ→Mới)':>28} | Câu hỏi")
    print(f"  {'─'*7}-+-{'─'*7}-+-{'─'*7}-+-{'─'*7}-+-{'─'*28}-+-{'─'*30}")

    rows = shuffled.questions
    if max_rows:
        rows = rows[:max_rows]

    for q in rows:
        mapping_str = ""
        if q.options_were_shuffled:
            # Hiện mapping: NEW→ORIG
            pairs = [f"{new}→{orig}" for new, orig in sorted(q.option_mapping.items())]
            mapping_str = "  ".join(pairs)
        else:
            mapping_str = "(giữ nguyên)"

        stem = q.stem_text[:30].replace('\n', ' ')
        if len(q.stem_text) > 30:
            stem += "…"

        print(f"  {q.new_number:>7} | {q.original_number:>7} | "
              f"{q.original_correct_answer:>7} | {q.correct_answer:>7} | "
              f"{mapping_str:<28} | {stem}")

    if max_rows and shuffled.total_questions > max_rows:
        remaining = shuffled.total_questions - max_rows
        print(f"  {'...':>7}   {'...':>7}   {'...':>7}   {'...':>7}   "
              f"{'... (' + str(remaining) + ' câu nữa)':<28}")


def print_options_detail(original: Exam, shuffled: Exam,
                         q_nums: List[int]) -> None:
    """In chi tiết đáp án của các câu được chỉ định."""
    orig_map = {q.original_number: q for q in original.questions}
    shuf_map = {q.original_number: q for q in shuffled.questions}

    print(f"\n  Chi tiết đáp án (câu gốc {q_nums}):")
    for n in q_nums:
        oq = orig_map.get(n)
        sq = shuf_map.get(n)
        if not oq or not sq:
            print(f"    Q{n}: không tìm thấy")
            continue

        print(f"\n    ── Q{n} (mới: Q{sq.new_number}) ──")
        print(f"    Stem: {sq.stem_text[:60]}")
        for letter in 'ABCD':
            o_opt = oq.options.get(letter)
            s_opt = sq.options.get(letter)
            if not s_opt:
                continue
            is_correct = "★" if s_opt.is_correct else " "
            new_letter = s_opt.current_letter
            orig_letter = s_opt.original_letter
            orig_text = o_opt.text[:30] if o_opt else "?"
            new_text  = s_opt.text[:30]
            print(f"    {is_correct} {new_letter} (gốc={orig_letter}): {new_text}")
        print(f"    → Đáp án gốc: {oq.correct_answer}  |  Đáp án mới: {sq.correct_answer}")


def print_verify_result(result: dict, label: str = "") -> None:
    tag = f" [{label}]" if label else ""
    if result['ok']:
        print(f"  ✅ VERIFY{tag}: PASS")
    else:
        print(f"  ❌ VERIFY{tag}: FAIL — {len(result['errors'])} lỗi")
        for e in result['errors']:
            print(f"     ❌ {e}")
    for w in result['warnings']:
        print(f"     ⚠️  {w}")


def print_answer_key(original: Exam, exams: List[Exam],
                     labels: List[str] = None) -> None:
    """In bảng đáp án tổng hợp (giống Excel)."""
    n = original.total_questions
    orig_map = {q.original_number: q for q in original.questions}
    labels = labels or [f"Mã {i+1:03d}" for i in range(len(exams))]

    # Build answer lookup: exam_idx → {new_num: answer}
    answer_lookups = []
    for exam in exams:
        lookup = {q.new_number: q.correct_answer for q in exam.questions}
        answer_lookups.append(lookup)

    print(f"\n  {'Câu':>6} | {'Đề gốc':>8}", end="")
    for lbl in labels:
        print(f" | {lbl:>8}", end="")
    print()
    print(f"  {'─'*6}-+-{'─'*8}" + "-+-" + "-+-".join("─"*8 for _ in labels))

    for i in range(1, n + 1):
        orig_ans = orig_map.get(i)
        o_str = orig_ans.correct_answer if orig_ans else "?"
        print(f"  {i:>6} | {o_str:>8}", end="")
        for lookup in answer_lookups:
            ans = lookup.get(i, "?")
            print(f" | {ans:>8}", end="")
        print()


# ════════════════════════════════════════════════════════════════
# Export CSV
# ════════════════════════════════════════════════════════════════

def export_csv(original: Exam, exams: List[Exam],
               labels: List[str], outpath: str) -> None:
    """Xuất bảng đáp án ra CSV."""
    orig_map = {q.original_number: q for q in original.questions}
    rows = []
    for exam, lbl in zip(exams, labels):
        for q in exam.questions:
            rows.append({
                'Đề': lbl,
                'Câu mới': q.new_number,
                'Câu gốc': q.original_number,
                'ĐA gốc': q.original_correct_answer,
                'ĐA mới': q.correct_answer,
                'Trộn ĐA': 'Có' if q.options_were_shuffled else 'Không',
            })

    with open(outpath, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)
    print(f"\n  💾 Đã xuất CSV: {outpath}")


# ════════════════════════════════════════════════════════════════
# Dữ liệu mẫu & load DOCX
# ════════════════════════════════════════════════════════════════

def make_demo_exam(n: int = 20) -> Exam:
    """Tạo đề demo."""
    exam = Exam()
    correct_cycle = list('ABCD')
    for i in range(1, n + 1):
        correct = correct_cycle[(i - 1) % 4]
        opts: Dict[str, Option] = {}
        for letter in 'ABCD':
            opts[letter] = Option(
                letter=letter,
                text=f"Đáp án {letter} của câu {i}",
                is_correct=(letter == correct),
            )
        q = Question(
            num=i,
            stem=f"Câu {i}: Đây là câu hỏi thứ {i}. Đáp án đúng là {correct}.",
            correct=correct,
            options=opts,
        )
        exam.questions.append(q)
    exam.total_questions = n
    return exam


def load_docx_exam(filepath: str) -> Exam:
    """
    Load đề từ file DOCX thật (dùng parser.py nếu có, fallback đọc thẳng).
    """
    try:
        # Thử dùng parser.py của project
        sys.path.insert(0, os.path.dirname(filepath))
        sys.path.insert(0, '.')
        from parser import parse_docx as project_parse
        raw = project_parse(filepath)
        # Convert models.ExamDocument → debug Exam
        exam = Exam()
        exam.total_questions = raw.total_questions
        for rq in raw.questions:
            opts: Dict[str, Option] = {}
            for letter, opt_data in rq.options.items():
                opts[letter] = Option(
                    letter=opt_data.original_letter,
                    text=opt_data.text,
                    is_correct=opt_data.is_correct,
                )
            q = Question(
                num=rq.original_number,
                stem=rq.stem_text,
                correct=rq.correct_answer or '?',
                options=opts,
            )
            exam.questions.append(q)
        print(f"  ✅ Dùng parser.py: {exam.total_questions} câu")
        return exam
    except ImportError:
        pass

    # Fallback: đọc thẳng bằng python-docx
    print("  ℹ️  parser.py không có — fallback đọc DOCX trực tiếp")
    return _fallback_load_docx(filepath)


def _fallback_load_docx(filepath: str) -> Exam:
    """Fallback: parse DOCX tối giản."""
    try:
        from docx import Document
    except ImportError:
        print("❌ python-docx chưa cài: pip install python-docx")
        sys.exit(1)

    doc = Document(filepath)
    exam = Exam()

    Q_PAT = re.compile(r'(?:Question|Câu)\s*(\d+)', re.IGNORECASE)
    OPT_4 = re.compile(r'([A-D])\.\s+(.+?)(?=\s+[A-D]\.\s+|$)', re.DOTALL)
    OPT_1 = re.compile(r'^\s*(?:\t)?([A-D])\.\s+(.+)$')

    current_q: Optional[Question] = None
    pending_opts_text: Optional[str] = None  # dòng options chờ gắn vào câu

    def _detect_correct(para) -> Set[str]:
        """Detect đáp án đúng qua màu đỏ."""
        correct_letters: Set[str] = set()
        full_text = para.text
        for run in para.runs:
            try:
                rgb = run.font.color.rgb
                r, g, b = int(rgb[0]), int(rgb[1]), int(rgb[2])
                if r > 180 and g < 100 and b < 100 and run.text.strip():
                    # Tìm xem run này thuộc đáp án nào
                    run_txt = run.text
                    # Tìm letter gần nhất trước run trong full text
                    idx = full_text.find(run_txt)
                    before = full_text[:idx] if idx >= 0 else ""
                    m = re.findall(r'([A-D])\.\s', before)
                    if m:
                        correct_letters.add(m[-1])
            except Exception:
                pass
            if run.font.underline and run.text.strip():
                idx = full_text.find(run.text)
                before = full_text[:idx] if idx >= 0 else ""
                m = re.findall(r'([A-D])\.\s', before)
                if m:
                    correct_letters.add(m[-1])
        return correct_letters

    def _parse_options_from_text(text: str, correct_letters: Set[str]) -> Dict[str, Option]:
        """Parse options từ text dạng 'A. ... B. ... C. ... D. ...'"""
        opts: Dict[str, Option] = {}
        matches = OPT_4.findall(text)
        for letter, opt_text in matches:
            opts[letter] = Option(
                letter=letter,
                text=opt_text.strip(),
                is_correct=(letter in correct_letters),
            )
        return opts

    for para in doc.paragraphs:
        txt = para.text.strip()
        if not txt:
            continue

        # Nhận dạng câu hỏi mới
        qm = Q_PAT.match(txt)
        if qm:
            q_num = int(qm.group(1))

            # Kiểm tra options inline (A. ... B. ... trên cùng dòng)
            inline_opts = OPT_4.findall(txt)
            correct_letters = _detect_correct(para)

            if len(inline_opts) >= 4:
                # Q + ABCD đều inline
                opts = _parse_options_from_text(txt, correct_letters)
                correct = (correct_letters & set(opts.keys()))
                correct_letter = next(iter(correct)) if correct else 'A'
                q = Question(q_num, txt, correct_letter, opts)
                exam.questions.append(q)
                current_q = q
            else:
                # Stem riêng, options ở paragraph sau
                q = Question(q_num, txt, '?', {})
                exam.questions.append(q)
                current_q = q

            continue

        # Nhận dạng dòng options (bắt đầu bằng tab + A.)
        opt_line_match = re.match(r'^\s*\t?\s*([A-D])\.\s', txt)
        inline_opts = OPT_4.findall(txt)
        if (opt_line_match or len(inline_opts) >= 2) and current_q:
            correct_letters = _detect_correct(para)
            opts = _parse_options_from_text(txt, correct_letters)
            if opts:
                current_q.options.update(opts)
                # Set correct answer
                correct = correct_letters & set(opts.keys())
                if correct:
                    letter = next(iter(correct))
                    # Mark is_correct
                    for l, o in current_q.options.items():
                        o.is_correct = False
                    current_q.options[letter].is_correct = True
                    current_q.correct_answer = letter
                    current_q.original_correct_answer = letter

    # Finalize
    valid_qs = []
    for q in exam.questions:
        if q.options and len(q.options) >= 2:
            # Đảm bảo có 4 options
            for letter in 'ABCD':
                if letter not in q.options:
                    q.options[letter] = Option(letter=letter,
                                               text=f"[missing]",
                                               is_correct=False)
            # Đảm bảo có đáp án đúng
            if q.correct_answer == '?' or not q.correct_answer:
                q.correct_answer = 'A'
                q.original_correct_answer = 'A'
                q.options['A'].is_correct = True
            valid_qs.append(q)

    exam.questions = valid_qs
    exam.total_questions = len(valid_qs)
    print(f"  ✅ Fallback parse: {exam.total_questions} câu")
    return exam


# ════════════════════════════════════════════════════════════════
# Parse group string
# ════════════════════════════════════════════════════════════════

def _parse_range(range_str: str) -> List[Tuple[int, int]]:
    """Parse '1-10, 15-20' thành [(1,10),(15,20)]."""
    result = []
    for part in range_str.split(','):
        part = part.strip()
        m = re.match(r'(\d+)\s*[-–]\s*(\d+)', part)
        if m:
            result.append((int(m.group(1)), int(m.group(2))))
        elif part.isdigit():
            n = int(part)
            result.append((n, n))
    return result


def parse_groups(groups_str: str, total_q: int) -> List[GroupConfig]:
    """
    Parse chuỗi phân nhóm.
    VD: 'g0:1-4, g3:5-50'  hoặc  '#g0:1-4, g3:5-50'
    """
    configs: List[GroupConfig] = []
    parts = re.split(r',\s*(?=#?g)', groups_str.strip())

    for part in parts:
        part = part.strip()
        m = re.match(r'(#?)g([0-3])\s*:\s*(.+)', part, re.IGNORECASE)
        if not m:
            print(f"  ⚠️  Bỏ qua nhóm không hợp lệ: '{part}'")
            continue
        fixed = (m.group(1) == '#')
        gtype = int(m.group(2))
        ranges = _parse_range(m.group(3))
        if ranges:
            configs.append(GroupConfig(gtype, fixed, ranges))

    if not configs:
        print(f"  ⚠️  Không parse được nhóm từ '{groups_str}' → dùng g3: 1-{total_q}")
        configs = [GroupConfig(3, False, [(1, total_q)])]

    return configs


# ════════════════════════════════════════════════════════════════
# Main
# ════════════════════════════════════════════════════════════════

def main():
    ap = argparse.ArgumentParser(
        description="Debug tool cho phần trộn đề — kiểm tra TRƯỚC KHI chạy main",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ví dụ:
  python debug_shuffle.py --demo
  python debug_shuffle.py --demo --groups "g0:1-4, g3:5-20" --num 3
  python debug_shuffle.py de_goc.docx
  python debug_shuffle.py de_goc.docx --groups "g0:1-4, g3:5-50" --num 4
  python debug_shuffle.py de_goc.docx --groups "g3:1-50" --num 4 --verbose
  python debug_shuffle.py de_goc.docx --groups "g3:1-50" --num 4 --export-csv
  python debug_shuffle.py de_goc.docx --detail-q 5,12,19
        """
    )
    ap.add_argument("filepath", nargs='?', default=None,
                    help="File DOCX đề gốc (bỏ qua → chạy demo)")
    ap.add_argument("--demo", action="store_true",
                    help="Dùng đề mẫu 20 câu, không cần file")
    ap.add_argument("--demo-size", type=int, default=20,
                    help="Số câu trong đề demo (mặc định: 20)")
    ap.add_argument("--groups", "-g", default=None,
                    help='Phân nhóm, VD: "g0:1-4, g3:5-50"')
    ap.add_argument("--num", "-n", type=int, default=4,
                    help="Số đề cần trộn (mặc định: 4)")
    ap.add_argument("--verbose", "-v", action="store_true",
                    help="In đầy đủ tất cả câu hỏi (mặc định: chỉ 15 câu đầu)")
    ap.add_argument("--detail-q", default=None,
                    help="In chi tiết đáp án của câu gốc (VD: 5,12,19)")
    ap.add_argument("--export-csv", action="store_true",
                    help="Xuất bảng đáp án ra CSV")
    ap.add_argument("--answer-table", action="store_true",
                    help="In bảng đáp án tổng hợp")

    args = ap.parse_args()

    print_header("🔀 DEBUG SHUFFLE — Kiểm tra thuật toán trộn đề")

    # ── Load exam ──────────────────────────────────────────────
    if args.demo or args.filepath is None:
        print(f"\n  Chế độ: DEMO ({args.demo_size} câu mẫu)")
        exam = make_demo_exam(args.demo_size)
        default_groups = f"g0:1-4, g3:5-{args.demo_size}"
    else:
        if not os.path.exists(args.filepath):
            print(f"❌ File không tồn tại: {args.filepath}")
            sys.exit(1)
        print(f"\n  Chế độ: FILE DOCX — {args.filepath}")
        exam = load_docx_exam(args.filepath)
        default_groups = f"g3:1-{exam.total_questions}"

    print(f"  Đề gốc: {exam.total_questions} câu hỏi")

    # ── Parse groups ────────────────────────────────────────────
    groups_str = args.groups or default_groups
    configs = parse_groups(groups_str, exam.total_questions)

    print(f"\n  Phân nhóm:")
    for cfg in configs:
        print(f"    {cfg}  ({len(cfg.question_numbers)} câu)")

    # Kiểm tra coverage
    all_assigned = set()
    for cfg in configs:
        for n in cfg.question_numbers:
            if n in all_assigned:
                print(f"  ⚠️  Câu {n} bị gán 2 nhóm!")
            all_assigned.add(n)

    expected = set(range(1, exam.total_questions + 1))
    unassigned = expected - all_assigned
    if unassigned:
        print(f"  ⚠️  Câu chưa được gán nhóm: {sorted(unassigned)}")
        print(f"      → Tự động thêm vào g3")
        configs.append(GroupConfig(3, False,
                                   [(min(unassigned), max(unassigned))]))

    # ── Trộn ────────────────────────────────────────────────────
    exams = _generate(exam, configs, args.num)

    if not exams:
        print("❌ Không tạo được đề nào")
        sys.exit(1)

    labels = [f"Mã {i+1:03d}" for i in range(len(exams))]

    # ── In kết quả ──────────────────────────────────────────────
    max_rows = None if args.verbose else 15

    for idx, sh in enumerate(exams):
        print_section(f"📋 {labels[idx]}  ({sh.total_questions} câu)")
        print_mapping_table(exam, sh, labels[idx], max_rows=max_rows)

        if not args.verbose and sh.total_questions > (max_rows or 0):
            remaining = sh.total_questions - (max_rows or sh.total_questions)
            if remaining > 0:
                print(f"\n  ... ({remaining} câu nữa, dùng --verbose để xem tất cả)")

        vr = verify(exam, sh, configs)
        print()
        print_verify_result(vr, labels[idx])

        if args.detail_q:
            q_nums = [int(x.strip()) for x in args.detail_q.split(',')
                      if x.strip().isdigit()]
            print_options_detail(exam, sh, q_nums)

    # ── Bảng đáp án ─────────────────────────────────────────────
    if args.answer_table:
        print_section("📊 BẢNG ĐÁP ÁN TỔNG HỢP")
        print_answer_key(exam, exams, labels)

    # ── Uniqueness check ────────────────────────────────────────
    print_section("📊 TỔNG KẾT")
    sigs = [_signature(sh) for sh in exams]
    unique = len(set(sigs))
    print(f"  Số đề trộn:     {len(exams)}/{args.num}")
    print(f"  Unique đề:      {unique}/{len(exams)}")
    if unique == len(exams):
        print("  ✅ Tất cả đề KHÁC NHAU")
    else:
        print("  ❌ Có đề TRÙNG nhau!")

    # Tổng kết verify
    all_ok = all(verify(exam, sh, configs)['ok'] for sh in exams)
    if all_ok:
        print(f"  ✅ Tất cả verify PASS")
    else:
        failed = sum(1 for sh in exams if not verify(exam, sh, configs)['ok'])
        print(f"  ❌ {failed}/{len(exams)} đề FAIL verify")

    # ── Export CSV ──────────────────────────────────────────────
    if args.export_csv:
        out = "debug_shuffle_output.csv"
        export_csv(exam, exams, labels, out)

    print(f"\n{DIVIDER_THICK}")


if __name__ == "__main__":
    main()