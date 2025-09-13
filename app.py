
"""
pob_dropdown_patch.py
---------------------
Drop-in Streamlit block to render "Place of Birth" text input and a conditional
"Select from dropdown" widget that *disappears* as soon as the text input is cleared.

Usage:
------
from pob_dropdown_patch import render_place_inputs

# somewhere in your layout:
place_info = render_place_inputs(
    label_text="Place of Birth (City, State, Country) *",
    dropdown_label="Select Place (City, State, Country)",
    get_suggestions_fn=get_place_suggestions,   # optional callback: (str)->list[str]
    on_selection_fn=on_place_selected           # optional callback: (str)->None (set coords/UTC, etc.)
)

# place_info contains the current query, selection, and any dependent session state.
"""

from typing import Callable, List, Optional, Dict, Any
import streamlit as st

# Keys used in st.session_state by this module
_POB_QUERY_KEY   = "pob_query"
_POB_OPTIONS_KEY = "pob_options"
_POB_CHOICE_KEY  = "pob_choice"

# Optional dependent keys you may want to clear as well
_DEP_KEYS = ("pob_coords", "utc_offset", "tz_name")

def _clear_place_state():
    """Clear dropdown and dependent state safely."""
    for k in (_POB_QUERY_KEY, _POB_OPTIONS_KEY, _POB_CHOICE_KEY, *_DEP_KEYS):
        st.session_state.pop(k, None)

def _on_place_query_change():
    """Runs when the text input changes; if empty, nuke dependent state."""
    q = st.session_state.get(_POB_QUERY_KEY, "").strip()
    if not q:
        _clear_place_state()

def render_place_inputs(
    label_text: str = "Place of Birth (City, State, Country) *",
    dropdown_label: str = "Select Place (City, State, Country)",
    get_suggestions_fn: Optional[Callable[[str], List[str]]] = None,
    on_selection_fn: Optional[Callable[[str], None]] = None,
) -> Dict[str, Any]:
    """
    Renders the text input and (conditionally) the dropdown.
    The dropdown is removed as soon as the text field is cleared.

    Returns a dict with current 'query' and 'selection'. Any extra dependent
    state is managed by your app / callbacks.
    """
    # 1) Main text input
    st.text_input(
        label_text,
        key=_POB_QUERY_KEY,
        on_change=_on_place_query_change,
    )

    # 2) Conditional dropdown placeholder
    dropdown_ph = st.empty()

    query = st.session_state.get(_POB_QUERY_KEY, "").strip()
    selection = st.session_state.get(_POB_CHOICE_KEY, None)

    if query:
        # Fetch / reuse options
        if get_suggestions_fn is not None:
            try:
                options = get_suggestions_fn(query) or []
            except Exception as e:
                st.warning(f"Suggestion provider failed: {e}")
                options = []
            st.session_state[_POB_OPTIONS_KEY] = options
        else:
            options = st.session_state.get(_POB_OPTIONS_KEY, [])

        if options:
            with dropdown_ph.container():
                new_selection = st.selectbox(
                    dropdown_label,
                    options=options,
                    index=options.index(selection) if selection in options else 0,
                    key=_POB_CHOICE_KEY,
                )
                if new_selection and on_selection_fn:
                    try:
                        on_selection_fn(new_selection)
                    except Exception as e:
                        st.warning(f"Selection handler failed: {e}")
        else:
            dropdown_ph.empty()
    else:
        # Text cleared: wipe everything and remove dropdown from the DOM
        _clear_place_state()
        dropdown_ph.empty()

    return {
        "query": st.session_state.get(_POB_QUERY_KEY),
        "selection": st.session_state.get(_POB_CHOICE_KEY),
        "options": st.session_state.get(_POB_OPTIONS_KEY, []),
    }
