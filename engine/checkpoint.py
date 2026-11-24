# -*- coding: utf-8 -*-
"""One saved page at a time, so a stopped conversion continues from that page."""

from __future__ import annotations

import pickle
from pathlib import Path

from engine.paths import checkpoint_dir


def page_path(stamp: str, index: int) -> Path:
    return checkpoint_dir() / stamp / f"{int(index)}.pkl"


def load_page(stamp: str, index: int) -> dict | None:
    if not stamp:
        return None
    path = page_path(stamp, index)
    if not path.is_file():
        return None
    try:
        with path.open("rb") as handle:
            payload = pickle.load(handle)
    except (OSError, pickle.PickleError, EOFError):
        return None
    if not isinstance(payload, dict) or payload.get("error") or "tables" not in payload:
        return None
    return payload


def save_page(stamp: str, index: int, payload: dict) -> None:
    if not stamp or not isinstance(payload, dict):
        return
    path = page_path(stamp, index)
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        temporary = path.with_suffix(".tmp")
        with temporary.open("wb") as handle:
            pickle.dump(payload, handle, protocol=pickle.HIGHEST_PROTOCOL)
        temporary.replace(path)
    except OSError:
        try:
            path.with_suffix(".tmp").unlink(missing_ok=True)
        except OSError:
            return


def saved_count(stamp: str) -> int:
    if not stamp:
        return 0
    folder = checkpoint_dir() / stamp
    if not folder.is_dir():
        return 0
    return sum(1 for _ in folder.glob("*.pkl"))


def clear(stamp: str) -> None:
    if not stamp:
        return
    folder = checkpoint_dir() / stamp
    if not folder.is_dir():
        return
    for path in folder.glob("*"):
        try:
            path.unlink()
        except OSError:
            continue
    try:
        folder.rmdir()
    except OSError:
        return
