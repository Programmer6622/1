from pathlib import Path
from typing import List, Optional

from .errors import ConversionError, ParseError
from .exporter import export_mytest
from .extractors import extract_text
from .models import Question
from .parser import parse_questions_from_pdf, parse_questions_from_text


def extract_raw_text(path: str | Path) -> str:
    file_path = Path(path)
    ext = file_path.suffix.lower()
    if ext in {".xlsx", ".xlsm", ".xls", ".csv"}:
        raise ConversionError("Spreadsheet files are not supported.")
    return extract_text(file_path)


def _parse_questions(file_path: Path) -> List[Question]:
    ext = file_path.suffix.lower()
    if ext in {".xlsx", ".xlsm", ".xls", ".csv"}:
        raise ConversionError("Spreadsheet files are not supported.")
    if ext == ".doc":
        raise ConversionError(".doc is not supported. Please save as .docx.")
    if ext == ".pdf":
        try:
            return parse_questions_from_pdf(file_path)
        except ParseError:
            text = extract_text(file_path)
            return parse_questions_from_text(text)
    text = extract_text(file_path)
    return parse_questions_from_text(text)


def _summarize_numbers(numbers: List[int], limit: int = 80) -> str:
    shown = ", ".join(str(n) for n in numbers[:limit])
    if len(numbers) > limit:
        shown += f", ... (+{len(numbers) - limit} more)"
    return shown


def _summarize_ids(values: List[str], limit: int = 80) -> str:
    shown = ", ".join(values[:limit])
    if len(values) > limit:
        shown += f", ... (+{len(values) - limit} more)"
    return shown


def _has_reliable_source_numbers(questions: List[Question]) -> bool:
    source_numbers = [q.number for q in questions if q.number is not None]
    if len(source_numbers) < 10:
        return False
    coverage = len(source_numbers) / max(len(questions), 1)
    if coverage < 0.6:
        return False
    numbers = sorted(set(source_numbers))
    min_n, max_n = numbers[0], numbers[-1]
    span = max_n - min_n + 1
    return span <= max(len(numbers) * 3, 50)


def _report_question_id(seq: int, q: Question, keep_numbers: bool, use_pairs: bool) -> str:
    if keep_numbers:
        return str(q.number if q.number is not None else seq)
    if use_pairs and q.number is not None:
        return f"{seq}.{q.number}"
    return str(seq)


def _detect_missing_numbers(questions: List[Question]) -> Optional[tuple[int, int, List[int]]]:
    numbered = [q.number for q in questions if q.number is not None]
    if not numbered or len(questions) < 10:
        return None
    coverage = len(numbered) / max(len(questions), 1)
    if coverage < 0.85:
        return None
    numbers = sorted(set(numbered))
    min_n, max_n = numbers[0], numbers[-1]
    span = max_n - min_n + 1
    if span > int(len(numbers) * 1.4):
        return None
    number_set = set(numbers)
    missing = [n for n in range(min_n, max_n + 1) if n not in number_set]
    if not missing:
        return None
    return min_n, max_n, missing


def _detect_missing_source_numbers_for_pairs(
    questions: List[Question],
) -> Optional[tuple[int, int, List[int]]]:
    source_numbers = [q.number for q in questions if q.number is not None]
    if len(source_numbers) < 10:
        return None
    numbers = sorted(set(source_numbers))
    min_n, max_n = numbers[0], numbers[-1]
    missing = [n for n in range(min_n, max_n + 1) if n not in set(numbers)]
    if not missing:
        return None
    return min_n, max_n, missing


def _build_report(questions: List[Question], keep_numbers: bool = False) -> Optional[str]:
    sections: List[str] = []
    use_pairs = (not keep_numbers) and _has_reliable_source_numbers(questions)

    # Missing-number analysis is meaningful only when preserving source numbering.
    if keep_numbers:
        missing_info = _detect_missing_numbers(questions)
        if missing_info:
            min_n, max_n, missing = missing_info
        else:
            numbered = [q.number for q in questions if q.number is not None]
            if numbered:
                min_n, max_n = min(numbered), max(numbered)
                missing = []
            else:
                min_n = max_n = 0
                missing = []
        sections.append(
            "\n".join(
                [
                    "Missing questions check (source numbering).",
                    f"Parsed: {len(questions)}",
                    f"Number range: {min_n}-{max_n}",
                    f"Missing count: {len(missing)}",
                    f"Missing numbers: {_summarize_numbers(missing) if missing else 'None'}",
                ]
            )
        )
    else:
        pair_missing_info = _detect_missing_source_numbers_for_pairs(questions)
        if use_pairs and pair_missing_info:
            min_n, max_n, missing = pair_missing_info
        elif use_pairs:
            source_numbers = [q.number for q in questions if q.number is not None]
            if source_numbers:
                min_n, max_n = min(source_numbers), max(source_numbers)
                missing = []
            else:
                min_n = max_n = 0
                missing = []
        else:
            min_n = max_n = 0
            missing = []
        if use_pairs:
            sections.append(
                "\n".join(
                    [
                        "Missing questions check (paired IDs).",
                        f"Parsed: {len(questions)}",
                        f"Source number range: {min_n}-{max_n}",
                        f"Missing count: {len(missing)}",
                        f"Missing source numbers: {_summarize_numbers(missing) if missing else 'None'}",
                    ]
                )
            )
        else:
            sections.append(
                "\n".join(
                    [
                        "Missing questions check (sequential IDs).",
                        f"Parsed: {len(questions)}",
                        "Source numbering: sparse or unavailable",
                    ]
                )
            )

    no_correct: List[str] = []
    low_options: List[str] = []
    for seq, q in enumerate(questions, start=1):
        q_id = _report_question_id(seq, q, keep_numbers, use_pairs)
        if len(q.options) < 2:
            low_options.append(q_id)
        if not any(opt.is_correct for opt in q.options):
            no_correct.append(q_id)

    if no_correct:
        sections.append(
            "\n".join(
                [
                    "Questions without a detected correct answer.",
                    f"Count: {len(no_correct)}",
                    f"Numbers: {_summarize_ids(no_correct)}",
                ]
            )
        )

    if low_options:
        sections.append(
            "\n".join(
                [
                    "Questions with too few options detected.",
                    f"Count: {len(low_options)}",
                    f"Numbers: {_summarize_ids(low_options)}",
                ]
            )
        )

    if not sections:
        return None
    return "\n\n".join(sections) + "\n"


def _sort_questions_for_output(questions: List[Question]) -> List[Question]:
    numbered = [q for q in questions if q.number is not None]
    unnumbered = [q for q in questions if q.number is None]
    if not numbered:
        return questions
    numbered = sorted(numbered, key=lambda q: int(q.number))
    return numbered + unnumbered


def convert_file(path: str | Path, keep_numbers: bool = False) -> str:
    file_path = Path(path)
    if not file_path.exists():
        raise ConversionError(f"File not found: {file_path}")
    questions = _sort_questions_for_output(_parse_questions(file_path))
    return export_mytest(questions, keep_numbers=keep_numbers)


def convert_file_with_report(
    path: str | Path, keep_numbers: bool = False
) -> tuple[str, Optional[str]]:
    file_path = Path(path)
    if not file_path.exists():
        raise ConversionError(f"File not found: {file_path}")
    questions = _sort_questions_for_output(_parse_questions(file_path))
    output = export_mytest(questions, keep_numbers=keep_numbers)
    report = _build_report(questions, keep_numbers=keep_numbers)
    return output, report
