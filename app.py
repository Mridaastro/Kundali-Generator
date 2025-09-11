
# app_pob_utc.py
# Focused, clean implementation of:
# - Place of Birth (POB) text input
# - One-time dropdown for disambiguation
# - Auto-population of UTC offset when a place is committed
#
# Integration notes:
# - Requires `timezonefinder` and `pytz` in requirements.txt
# - Optionally uses Geoapify (recommended). Put GEOAPIFY_API_KEY in Streamlit Secrets.
#
#     [secrets.toml]
#     GEOAPIFY_API_KEY = "your_key_here"
#
# If you don't have Geoapify, you can swap `search_places` to hit Nominatim instead.

import datetime
import json
import urllib.parse
from typing import List, Tuple

import streamlit as st

try:
    from timezonefinder import TimezoneFinder
except Exception as _e:
    st.error("Missing dependency: timezonefinder. Please add 'timezonefinder' to requirements.txt")
    raise

try:
    import pytz
except Exception as _e:
    st.error("Missing dependency: pytz. Please add 'pytz' to requirements.txt")
    raise

try:
    import requests
except Exception as _e:
    st.error("Missing dependency: requests. Please add 'requests' to requirements.txt")
    raise


st.set_page_config(page_title="POB + UTC Demo", layout="wide")

st.markdown(
    "<h2 style='text-align:center;margin-top:0'>Place of Birth & UTC Auto-fill (Clean)</h2>",
    unsafe_allow_html=True,
)

# ----------------------------
# Helpers
# ----------------------------
def _tz_to_hours(tz_str: str) -> float:
    """
    Accept '5.5', '+05:30', '-04:00', '0', etc., and return hours as float.
    """
    s = (tz_str or "").strip()
    if not s:
        raise ValueError("empty tz")
    if ":" in s:
        sign = -1.0 if s.startswith("-") else 1.0
        hh, mm = s.lstrip("+-").split(":", 1)
        return sign * (int(hh) + int(mm)/60.0)
    return float(s)


def get_timezone_offset_simple(lat: float, lon: float) -> float:
    """
    Return the *current* UTC offset (in hours) at the given coordinates.
    Independent of DOB/TOB; this is just to auto-fill the UI immediately.
    Your birth calculations can recompute an exact historical offset later.
    """
    try:
        tf = TimezoneFinder()
        tzname = tf.timezone_at(lat=lat, lng=lon) or "Etc/UTC"
        tz = pytz.timezone(tzname)
        now_local_naive = datetime.datetime.now().replace(second=0, microsecond=0)
        now_local = tz.localize(now_local_naive, is_dst=None)
        return now_local.utcoffset().total_seconds() / 3600.0
    except Exception:
        return 0.0


def hours_to_hhmm_str(hours: float) -> str:
    sign = '+' if hours >= 0 else '-'
    total_minutes = int(round(abs(hours) * 60))
    hh, mm = divmod(total_minutes, 60)
    return f"{sign}{hh:02d}:{mm:02d}"


def geofmt_city_state_country(props: dict) -> str:
    """
    Compose a clean "City, State, Country" string from Geoapify feature properties.
    """
    city = props.get("city") or props.get("name") or props.get("town") or props.get("state_district")
    state = props.get("state") or props.get("county")
    country = props.get("country")
    parts = [p for p in [city, state, country] if p]
    return ", ".join(parts)


def search_places(query: str, api_key: str, limit: int = 6) -> List[Tuple[str, float, float]]:
    """
    Returns a list of (display_name, lat, lon).
    Uses Geoapify Autocomplete.
    """
    if not query or not api_key:
        return []

    url = (
        "https://api.geoapify.com/v1/geocode/autocomplete?"
        + urllib.parse.urlencode(
            {
                "text": query,
                "format": "json",
                "limit": limit,
                "apiKey": api_key,
            }
        )
    )
    try:
        resp = requests.get(url, timeout=8)
        if resp.status_code != 200:
            return []
        data = resp.json()
        results = []
        for feat in data.get("results", []):
            props = feat
            lat = props.get("lat")
            lon = props.get("lon")
            disp = geofmt_city_state_country(props)
            if disp and isinstance(lat, (int, float)) and isinstance(lon, (int, float)):
                # Keep only entries that clearly look like City, State, Country
                if len([p for p in disp.split(",") if p.strip()]) >= 3:
                    results.append((disp, float(lat), float(lon)))
        # Dedup by name while preserving order
        seen = set()
        unique = []
        for disp, lat, lon in results:
            if disp not in seen:
                seen.add(disp)
                unique.append((disp, lat, lon))
        return unique[:limit]
    except Exception:
        return []


# ----------------------------
# UI
# ----------------------------
row1c1, row1c2, row1c3 = st.columns([1.4, 1.2, 0.8])

with row1c1:
    st.markdown("**Place of Birth (City, State, Country)**")
    place_input_val = st.text_input(
        "Place of Birth (City, State, Country)",
        key="place_input",
        label_visibility="collapsed",
        placeholder="Start typing city…",
        help="Type your city; you’ll get a one-time dropdown to choose exact City, State, Country",
    )

with row1c3:
    st.markdown("**UTC offset**")
    st.text_input(
        "UTC offset",
        key="tz_input",
        label_visibility="collapsed",
        disabled=True,  # keep read-only
        placeholder="+00:00",
        help="Auto-fills when you select a place",
    )

# For disambiguation dropdown, we render it *only* when needed.
api_key = st.secrets.get("GEOAPIFY_API_KEY", "")
committed = st.session_state.get("last_place_checked", "")
typed = (place_input_val or "").strip()

def set_committed_place(display: str, lat: float, lon: float) -> None:
    st.session_state["place_input"] = display
    st.session_state["pob_display"] = display
    st.session_state["pob_lat"] = lat
    st.session_state["pob_lon"] = lon
    st.session_state["last_place_checked"] = display
    # Auto-fill UTC
    try:
        offset = get_timezone_offset_simple(lat, lon)
        st.session_state["tz_input"] = hours_to_hhmm_str(offset)
    except Exception:
        pass


# Only look up when user typed something new (not already committed this exact value)
if typed and typed != committed:
    cands = search_places(typed, api_key, limit=6)
    if not cands:
        # Clear lat/lon if user typed a random string
        st.session_state.pop("pob_lat", None)
        st.session_state.pop("pob_lon", None)
        st.session_state.pop("pob_display", None)
    else:
        # If there is a single unambiguous candidate, auto-commit it.
        if len(cands) == 1:
            disp, lat, lon = cands[0]
            set_committed_place(disp, lat, lon)
            # Important: ensure no dropdown remains
            st.session_state.pop("pob_choice", None)
            st.experimental_rerun()
        else:
            # Show a one-time dropdown to choose the exact city/state/country
            with row1c2:
                st.markdown("**Select Place (City, State, Country)**")
                options = ["Select from dropdown"] + [c[0] for c in cands]
                choice = st.selectbox(
                    "Select Place",
                    options,
                    index=0,
                    key="pob_choice",
                    label_visibility="collapsed",
                    help="Pick the exact city, state, country",
                )
                if choice and choice != "Select from dropdown":
                    sel = next((c for c in cands if c[0] == choice), None)
                    if sel:
                        disp, lat, lon = sel
                        set_committed_place(disp, lat, lon)
                        # Clear selectbox so it disappears on rerun
                        st.session_state.pop("pob_choice", None)
                        st.experimental_rerun()

# Debug / visibility panel
with st.expander("Debug (you can hide this in production)"):
    st.write({
        "place_input": st.session_state.get("place_input"),
        "pob_display": st.session_state.get("pob_display"),
        "pob_lat": st.session_state.get("pob_lat"),
        "pob_lon": st.session_state.get("pob_lon"),
        "tz_input": st.session_state.get("tz_input"),
        "last_place_checked": st.session_state.get("last_place_checked"),
    })
    st.caption("On a successful commit, the dropdown (pob_choice) is removed and tz_input is auto-filled.")

st.success("Ready. Type your city, pick the exact place once, and the UTC field will auto-fill.")
