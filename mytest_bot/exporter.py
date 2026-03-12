from typing import List

from .models import Question


def _has_reliable_source_numbers(questions: List[Question]) -> bool:
    numbered = [q.number for q in questions if q.number is not None]
    if len(numbered) < 10:
        return False
    coverage = len(numbered) / max(len(questions), 1)
    if coverage < 0.6:
        return False
    numbers = sorted(set(numbered))
    min_n, max_n = numbers[0], numbers[-1]
    span = max_n - min_n + 1
    return span <= max(len(numbers) * 3, 50)


def export_mytest(questions: List[Question], keep_numbers: bool = False) -> str:
    lines = []
    counter = 1
    use_paired_ids = (not keep_numbers) and _has_reliable_source_numbers(questions)
    for q in questions:
        if keep_numbers:
            number = q.number if q.number else counter
            number_label = str(number)
        else:
            # Keep sequential id for stability and append source id when available.
            if use_paired_ids and q.number is not None:
                number_label = f"{counter}.{q.number}"
            else:
                number_label = str(counter)
        line = f"#{number_label} {q.text}".strip()
        lines.append(line)
        for opt in q.options:
            prefix = "+" if opt.is_correct else "-"
            lines.append(f"{prefix}{opt.text}")
        lines.append("")
        counter += 1
    return "\n".join(lines).rstrip() + "\n"
