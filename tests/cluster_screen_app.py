# -*- coding: utf-8 -*-
"""Tiny screen the cluster test drives. It does not open a PDF."""

import streamlit as st

from excel_review import cluster_decision, progress_text, remaining_cells

MEMBERS = [
    {"page": 1, "plate": str(10000 + index), "field": "color", "value": "صدفي"}
    for index in range(17)
]
OUTLIERS = [
    {"page": 1, "plate": "27191", "field": "color", "value": "فضي"},
    {"page": 1, "plate": "21804", "field": "color", "value": "رمادي"},
]


def _identity(item: dict) -> str:
    return f"{item['page']}|{item['plate']}|{item['field']}"


def _group() -> dict:
    ejected = set(st.session_state.get("ejected") or [])
    return {
        "label": "صدفي",
        "field": "color",
        "members": [item for item in MEMBERS if _identity(item) not in ejected],
        "outliers": [item for item in OUTLIERS if _identity(item) not in ejected],
    }


if "accepted" not in st.session_state:
    st.session_state.accepted = []
    st.session_state.ejected = []
    st.session_state.scanned = False

group = _group()
decision = cluster_decision(group, set(st.session_state.ejected), bool(st.session_state.scanned))
st.markdown(progress_text(len(st.session_state.accepted), 19, 40))
st.write(f"remaining={remaining_cells(19, [], st.session_state.accepted)}")
if st.button("إخراج المختلفة", key="eject"):
    st.session_state.ejected = ["1|27191|color", "1|21804|color"]
    st.rerun()
if st.button("عرض المجموعة", key="scan"):
    st.session_state.scanned = True
    st.rerun()
if st.button(decision["label"], key="accept", disabled=not decision["ready"]):
    st.session_state.accepted = [_identity(item) for item in decision["write"]]
    st.rerun()
