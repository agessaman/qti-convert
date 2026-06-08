"""Generate randomized test forms from parsed assessment structure."""

import random
from typing import List, Literal, Optional

from qti_parser import item
from qti_parser.assessment import AssessmentBlock, GroupBlock, ItemBlock

ShuffleScope = Literal["group", "test"]


def form_label(index: int) -> str:
    """Return form label for index: 0 -> Form A, 25 -> Form Z, 26 -> Form AA."""
    letters = []
    n = index
    while True:
        letters.append(chr(ord("A") + n % 26))
        n = n // 26 - 1
        if n < 0:
            break
    return "Form " + "".join(reversed(letters))


def _select_from_pool(items: list, selection_count: Optional[int], rng: random.Random) -> list:
    if not items:
        return []
    count = selection_count if selection_count is not None else len(items)
    count = min(count, len(items))
    return rng.sample(items, count)


def generate_questions(
    blocks: List[AssessmentBlock],
    scope: ShuffleScope,
    rng: random.Random,
    *,
    apply_pools: bool = True,
) -> List[dict]:
    """Build question list for one form from assessment blocks."""
    questions: List[dict] = []

    for block in blocks:
        if isinstance(block, ItemBlock):
            questions.append(item.get_question(block.xml_item))
        elif isinstance(block, GroupBlock):
            pool = block.items
            if apply_pools:
                selected = _select_from_pool(pool, block.selection_count, rng)
            else:
                selected = list(pool)
            if scope == "group" and apply_pools and len(selected) > 1:
                selected = selected.copy()
                rng.shuffle(selected)
            for xml_item in selected:
                question = item.get_question(xml_item)
                question["section_id"] = block.ident
                question["section_title"] = block.title
                questions.append(question)

    if scope == "test" and len(questions) > 1:
        rng.shuffle(questions)

    return questions
