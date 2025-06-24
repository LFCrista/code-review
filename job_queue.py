from dataclasses import dataclass
from typing import Sequence


@dataclass
class Job:
    pdfs: Sequence[str]
    link_gpt: str
    doc_id: str
    prompt: str
    max_per_tab: int
