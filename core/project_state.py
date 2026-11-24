# -*- coding: utf-8 -*-
"""
Per-document project state for PDF → Word conversion choices.

Stores the chosen mode and score per page so a later re-convert does not
discard manual picks. Supports undo of the latest page re-conversion.
"""

from __future__ import annotations

import json
import time
from dataclasses import asdict, dataclass, field
from pathlib import Path


@dataclass
class PageChoice:
    mode: str = "faithful"
    score: float | None = None
    target: float = 97.0
    accepted: bool = False
    needs_review: bool = False
    notes: list[str] = field(default_factory=list)
    tried: list[dict] = field(default_factory=list)
    # Snapshot before the last reconvert (mode + score) for undo.
    previous: dict | None = None


@dataclass
class ProjectState:
    version: int = 1
    pdf: str = ""
    docx: str = ""
    document_type: str = "text_document"
    pages: dict[str, PageChoice] = field(default_factory=dict)
    updated_at: float = 0.0

    def page(self, number: int) -> PageChoice:
        key = str(int(number))
        if key not in self.pages:
            self.pages[key] = PageChoice()
        return self.pages[key]

    def strategies(self) -> dict[int, str]:
        """Map for WordOptions.page_strategies."""
        return {int(k): v.mode for k, v in self.pages.items() if v.mode}


def load_project(path: str | Path) -> ProjectState:
    source = Path(path)
    if not source.is_file():
        return ProjectState()
    raw = json.loads(source.read_text(encoding="utf-8"))
    pages: dict[str, PageChoice] = {}
    for key, value in (raw.get("pages") or {}).items():
        pages[str(key)] = PageChoice(
            mode=str(value.get("mode") or "faithful"),
            score=value.get("score"),
            target=float(value.get("target") or 97.0),
            accepted=bool(value.get("accepted")),
            needs_review=bool(value.get("needs_review")),
            notes=list(value.get("notes") or []),
            tried=list(value.get("tried") or []),
            previous=value.get("previous"),
        )
    return ProjectState(
        version=int(raw.get("version") or 1),
        pdf=str(raw.get("pdf") or ""),
        docx=str(raw.get("docx") or ""),
        document_type=str(raw.get("document_type") or "text_document"),
        pages=pages,
        updated_at=float(raw.get("updated_at") or 0.0),
    )


def save_project(state: ProjectState, path: str | Path) -> Path:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    state.updated_at = time.time()
    payload = {
        "version": state.version,
        "pdf": state.pdf,
        "docx": state.docx,
        "document_type": state.document_type,
        "updated_at": state.updated_at,
        "pages": {key: asdict(choice) for key, choice in state.pages.items()},
    }
    target.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return target


def remember_before_reconvert(choice: PageChoice) -> None:
    choice.previous = {
        "mode": choice.mode,
        "score": choice.score,
        "accepted": choice.accepted,
        "needs_review": choice.needs_review,
        "notes": list(choice.notes),
    }


def undo_page_choice(choice: PageChoice) -> bool:
    """Restore the previous mode/score. Returns False when nothing to undo."""
    if not choice.previous:
        return False
    saved = choice.previous
    choice.mode = str(saved.get("mode") or "faithful")
    choice.score = saved.get("score")
    choice.accepted = bool(saved.get("accepted"))
    choice.needs_review = bool(saved.get("needs_review"))
    choice.notes = list(saved.get("notes") or [])
    choice.previous = None
    return True
