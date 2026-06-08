"""HTML stripping utilities for question text."""

import html
import re
from typing import Any, Dict, Optional


def strip_html(text: Optional[str]) -> Optional[str]:
    """Remove HTML tags and decode entities from text."""
    if text is None:
        return None
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"</p\s*>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    text = html.unescape(text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _strip_answer_entry(entry: Dict[str, Any]) -> None:
    if "text" in entry and entry["text"] is not None:
        entry["text"] = strip_html(entry["text"])
    for option in entry.get("options", []):
        if "text" in option and option["text"] is not None:
            option["text"] = strip_html(option["text"])


def strip_question(question: Dict[str, Any]) -> Dict[str, Any]:
    """Strip HTML from a parsed question and its answers."""
    if "text" in question and question["text"] is not None:
        question["text"] = strip_html(question["text"])
    for answer in question.get("answer", []):
        _strip_answer_entry(answer)
    return question


def strip_metadata(metadata: Dict[str, Any]) -> Dict[str, Any]:
    """Strip HTML from assessment metadata fields."""
    for field in ("title", "description"):
        if field in metadata and metadata[field] is not None:
            metadata[field] = strip_html(metadata[field])
    return metadata
