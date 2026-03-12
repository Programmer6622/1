from dataclasses import dataclass
from typing import List, Optional


@dataclass
class Option:
    text: str
    is_correct: bool = False


@dataclass
class Question:
    text: str
    options: List[Option]
    number: Optional[int] = None
