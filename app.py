
# app_pob_utc_v2.py
# Robust, widget-safe Place-of-Birth commit + UTC auto-fill demo.
# Copy the helpers + UI block into your main app.py.
#
# Requirements:
#   streamlit
#   requests
#   timezonefinder
#   pytz

import datetime
import urllib.parse
from typing import List, Tuple

import streamlit as st

try:
    from timezonefinder import TimezoneFinder
    import pytz
    import requests
except Exception as e:
    st.error("Missing dependency. Please ensure 'timezonefinder', 'pytz', and 'requests' are in requirements.txt")
    raise

st.set_page_config(page_title="POB/UTC Demo (Widget-Safe)", layout="wide")

st.markdown("<h2 style='text-align:center'>POB & UTC Autogeneration — Clean & Widget‑Safe</h2>", unsafe_allow_html=True)

# ----------------------------
# Helpers
# ----------------------------
def hours_to_hhmm_str(hours: float) -> str:
    sign = '+' if hours >= 0 else '-'
    total_minutes = int(round(abs(hours) * 60))
    hh, mm = divmod(total_minutes, 60)
    return f"{sign}{hh:02d}:{mm:02d}"

def _tz_to_hours(tz_str: str) -> float:
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
    Current UTC offset (in hours) based on lat/lon using timezonefinder+pytz.
    Independent of DOB/TOB so it can fill immediately.
    """
    tf = TimezoneFinder()
    tzname = tf.timezone_at(lat=lat, lng=lon) or "Etc/UTC"
    tz = pytz.timezone(tzname)
    # use now to avoid DST ambiguity in past dates.
    now_local_naive = datetime.datetime.now().replace(second=0, microsecond=0)
    now_local = tz.localize(now_local_naive, is_dst=None)
    return tz.utcoffset(now_local).total_seconds() / 3600.0

def geofmt_city_state_country(props: dict) -> str:
    city = props.get("city") or props.get("name") or props.get("town") or props.get("state_district")
    state = props.get("state") or props.get("county")
    country = props.get("country")
    parts = [p for p in [city, state, country] if p]
    return ", ".join(parts)

def search_places(query: str, api_key: str, limit: int = 6) -> List[Tuple[str, float, float]]:
    """
    Returns a list of (display_name, lat, lon) using Geoapify Autocomplete.
    """
    if not query or not api_key:
        return []
    url = (
        "https://api.geoapify.com/v1/geocode/autocomplete?"
        + urllib.parse.urlencode(
            {"text": query, "format": "json", "limit": limit, "apiKey": api_key}
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
                if len([p for p in disp.split(",") if p.strip()]) >= 3:
                    results.append((disp, float(lat), float(lon)))
        # Dedup by display
        seen, unique = set(), []
        for disp, lat, lon in results:
            if disp not in seen:
                seen.add(disp)
                unique.append((disp, lat, lon))
        return unique[:limit]
    except Exception:
        return []

# ----------------------------
# Session pre-processing: apply any pending commit BEFORE rendering widgets
# ----------------------------
pending = st.session_state.get("pob_pending_commit")  # {'display': str, 'lat': float, 'lon': float}
if pending:
    st.session_state["place_committed"] = pending["display"]
    st.session_state["pob_display"] = pending["display"]
    st.session_state["pob_lat"] = pending["lat"]
    st.session_state["pob_lon"] = pending["lon"]
    # Compute & store UTC string
    try:
        offset = get_timezone_offset_simple(pending["lat"], pending["lon"])
        st.session_state["tz_input"] = hours_to_hhmm_str(offset)
    except Exception:
        st.session_state["tz_input"] = "+00:00"
    # Clear the pending commit so it won't loop
    st.session_state.pop("pob_pending_commit", None)
    # Also clear the dropdown state to hide it on rerun
    st.session_state.pop("pob_choice", None)

# ----------------------------
# UI
# ----------------------------
row1c1, row1c2, row1c3 = st.columns([1.4, 1.2, 0.8])

with row1c1:
    st.markdown("**Place of Birth (City, State, Country)**")
    place_val_default = st.session_state.get("place_committed", "")
    place_input_val = st.text_input(
        "Place of Birth (City, State, Country)",
        key="place_input",
        value=place_val_default,
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
        disabled=True,
        placeholder="+00:00",
        help="Auto-fills when you select a place",
    )

api_key = st.secrets.get("GEOAPIFY_API_KEY", "")
typed = (place_input_val or "").strip()
committed = st.session_state.get("place_committed", "")

def commit_place(display: str, lat: float, lon: float):
    # Defer mutation until the next rerun to keep widget updates safe
    st.session_state["pob_pending_commit"] = {"display": display, "lat": lat, "lon": lon}
    st.rerun()

# Only search when user typed something different from the committed value
if typed and typed != committed:
    cands = search_places(typed, api_key, limit=6)
    if not cands:
        # Clear lat/lon if they typed a random string
        st.session_state.pop("pob_lat", None)
        st.session_state.pop("pob_lon", None)
        st.session_state.pop("pob_display", None)
    else:
        if len(cands) == 1:
            disp, lat, lon = cands[0]
            commit_place(disp, lat, lon)
        else:
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
                        commit_place(disp, lat, lon)

with st.expander("Debug (hide in production)"):
    st.write({
        "typed": typed,
        "place_committed": st.session_state.get("place_committed"),
        "pob_display": st.session_state.get("pob_display"),
        "pob_lat": st.session_state.get("pob_lat"),
        "pob_lon": st.session_state.get("pob_lon"),
        "tz_input": st.session_state.get("tz_input"),
        "pob_choice (dropdown)": st.session_state.get("pob_choice"),
        "pob_pending_commit": st.session_state.get("pob_pending_commit"),
    })

st.success("Ready. Type city → choose exact place once → textbox updates and UTC auto-fills.")
