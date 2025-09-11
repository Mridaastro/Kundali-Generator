
import streamlit as st
import json, re, urllib.parse, urllib.request
from dataclasses import dataclass
from typing import List, Optional, Tuple
import math

# --- Timezone helpers (no DOB/TOB dependency) ---
try:
    from timezonefinder import TimezoneFinder
    from datetime import datetime, timezone
    try:
        # Python 3.9+
        from zoneinfo import ZoneInfo
    except Exception:
        ZoneInfo = None
except Exception:
    TimezoneFinder = None
    ZoneInfo = None

st.set_page_config(page_title="MridaAstro — Kundali (POB Patch)", page_icon="🪐", layout="centered")

st.title("MridaAstro — Kundali Generator (POB-patched)")

# --------- Secrets / API keys ----------
GEOAPIFY_API_KEY = ""
try:
    # Streamlit secrets (preferred)
    GEOAPIFY_API_KEY = st.secrets.get("GEOAPIFY_API_KEY", "")
except Exception:
    GEOAPIFY_API_KEY = ""

# Also allow an env-var fallback if running outside Streamlit Cloud
import os
if not GEOAPIFY_API_KEY:
    GEOAPIFY_API_KEY = os.environ.get("GEOAPIFY_API_KEY", "")

if not GEOAPIFY_API_KEY:
    st.info("Add your **GEOAPIFY_API_KEY** to `st.secrets` (or environment) for place suggestions & validation.", icon="ℹ️")

# ---- Session keys ----
LAT_KEY = "pob_lat"
LON_KEY = "pob_lon"
TZ_KEY = "tz_input"
PLACE_KEY = "place_input"
CHOICE_IDX_KEY = "place_choice_idx"
TZ_ERR_KEY = "tz_error_msg"

for k, v in [(LAT_KEY, None), (LON_KEY, None), (TZ_KEY, ""), (PLACE_KEY, ""), (CHOICE_IDX_KEY, 0)]:
    st.session_state.setdefault(k, v)

# --- UI: Basic inputs (name/dob/tob are placeholders to keep structure familiar) ---
col1, col2 = st.columns(2)
with col1:
    name = st.text_input("Name", key="name_input", placeholder="Person Name")
with col2:
    st.text_input("UTC offset", key=TZ_KEY, help="Auto‑detected from Place of Birth", disabled=True)

# Place of Birth text box
place_typed = st.text_input("Place of Birth (City, State, Country)",
                            key=PLACE_KEY,
                            placeholder="e.g., Jabalpur, Madhya Pradesh, India")

# ---------- Core logic: Strict geocoding + Suggestions ----------

def _norm(s: str) -> str:
    s = (s or "").strip().lower()
    return re.sub(r"[^a-z]", "", s)

@dataclass
class Suggestion:
    label: str  # "City, State, Country"
    lat: float
    lon: float

def geocode_suggestions(typed: str, api_key: str, limit: int = 7) -> List[Suggestion]:
    """Return suggestions for ambiguous city names (city/town/village level only)."""
    out: List[Suggestion] = []
    typed = (typed or "").strip()
    if not typed or not api_key:
        return out

    # Use only the first token before comma as the city-key
    city_key = [p.strip() for p in typed.split(",") if p.strip()]
    city_key = city_key[0] if city_key else typed

    base = "https://api.geoapify.com/v1/geocode/search?"
    q = urllib.parse.urlencode({"text": city_key, "format": "json", "limit": max(10, limit), "apiKey": api_key})

    try:
        with urllib.request.urlopen(base + q, timeout=15) as r:
            j = json.loads(r.read().decode())
    except Exception:
        return out

    seen = set()
    for res in (j.get("results") or []):
        city = res.get("city") or res.get("town") or res.get("village") or ""
        if _norm(city) != _norm(city_key):
            continue  # keep exact city-name matches only
        state = res.get("state") or ""
        country = res.get("country") or ""
        lat = res.get("lat"); lon = res.get("lon")
        if lat is None or lon is None:
            continue
        label = ", ".join([p for p in (city, state, country) if p])
        key = (_norm(city), _norm(state), _norm(country))
        if key in seen:
            continue
        seen.add(key)
        out.append(Suggestion(label=label, lat=float(lat), lon=float(lon)))
        if len(out) >= limit:
            break
    return out

def geocode_strict(place: str, api_key: str) -> Tuple[float, float, str]:
    """
    STRICT geocoding:
    - Requires 'City, State, Country' (3 parts).
    - Accept only if city/state/country match (case/diacritic-insensitive).
    Returns (lat, lon, formatted_label).
    """
    if not api_key:
        raise RuntimeError("Place not found.")
    raw = (place or "").strip()
    parts = [p.strip() for p in raw.split(",") if p.strip()]
    if len(parts) < 3:
        raise RuntimeError("Please enter 'City, State, Country'.")

    typed_city, typed_state, typed_country = parts[0], parts[1], parts[-1]

    base = "https://api.geoapify.com/v1/geocode/search?"
    q = urllib.parse.urlencode({"text": raw, "format": "json", "limit": 1, "apiKey": api_key})
    with urllib.request.urlopen(base + q, timeout=15) as r:
        j = json.loads(r.read().decode())

    if not j.get("results"):
        raise RuntimeError("Place not found.")

    res = j["results"][0]
    city_res = res.get("city") or res.get("town") or res.get("village") or res.get("municipality") or res.get("county") or ""
    state_res = res.get("state") or ""
    country_res = res.get("country") or ""

    # City must match directly OR appear verbatim in 'formatted'
    if _norm(city_res) != _norm(typed_city):
        fmt = res.get("formatted", "")
        pat = r"\b" + re.escape(typed_city.strip()) + r"\b"
        if not re.search(pat, fmt, flags=re.IGNORECASE):
            raise RuntimeError("Place not found. Please enter City, State, Country correctly.")

    if _norm(state_res) != _norm(typed_state):
        raise RuntimeError("Place not found. Please enter City, State, Country correctly.")

    if _norm(country_res) not in (_norm(typed_country), "bharat", "hindustan", "india"):
        raise RuntimeError("Place not found. Please enter City, State, Country correctly.")

    lat = float(res["lat"]); lon = float(res["lon"])
    return lat, lon, res.get("formatted", raw)

def get_timezone_offset_simple(lat: float, lon: float) -> str:
    """Return UTC offset like '+5.5' based on current rules (independent of DOB/TOB)."""
    if TimezoneFinder is None:
        return ""
    tf = TimezoneFinder()
    tzname = tf.timezone_at(lat=lat, lng=lon)
    if not tzname:
        return ""
    # Compute offset *now* (not at DOB), as per requirement
    if ZoneInfo is None:
        return ""
    now_utc = datetime.now(timezone.utc)
    local = now_utc.astimezone(ZoneInfo(tzname))
    # offset in hours (can be fractional like 5.5)
    offset_hours = local.utcoffset().total_seconds() / 3600.0
    # format trimmed (avoid trailing .0)
    # keep .5 or .75 if present
    if abs(offset_hours - round(offset_hours)) < 1e-9:
        return f"{int(round(offset_hours)):+d}"
    # keep up to 2 decimals without trailing zeros
    s = f"{offset_hours:+.2f}".rstrip("0").rstrip(".")
    return s

# ---------- Behavior ----------
place_raw = (st.session_state.get(PLACE_KEY) or "").strip()
need_help = bool(place_raw) and (place_raw.count(",") < 2)

# If user typed only City (or City, State) → show disambiguation dropdown
if GEOAPIFY_API_KEY and need_help:
    sug = geocode_suggestions(place_raw, GEOAPIFY_API_KEY, limit=7)
    if len(sug) > 1:
        labels = ["— Select exact place —"] + [s.label for s in sug]
        st.selectbox("Similar city names found",
                     options=list(range(len(labels))),
                     format_func=lambda i: labels[i],
                     index=st.session_state.get(CHOICE_IDX_KEY, 0) or 0,
                     key=CHOICE_IDX_KEY)
        idx = st.session_state.get(CHOICE_IDX_KEY, 0) or 0
        if idx > 0:
            chosen = sug[idx - 1]
            st.session_state[PLACE_KEY] = chosen.label
            st.session_state[LAT_KEY] = chosen.lat
            st.session_state[LON_KEY] = chosen.lon
            try:
                st.session_state[TZ_KEY] = get_timezone_offset_simple(chosen.lat, chosen.lon)
                st.session_state.pop(TZ_ERR_KEY, None)
            except Exception:
                st.session_state[TZ_KEY] = ""
            st.rerun()
    elif len(sug) == 1:
        chosen = sug[0]
        st.session_state[PLACE_KEY] = chosen.label
        st.session_state[LAT_KEY] = chosen.lat
        st.session_state[LON_KEY] = chosen.lon
        try:
            st.session_state[TZ_KEY] = get_timezone_offset_simple(chosen.lat, chosen.lon)
            st.session_state.pop(TZ_ERR_KEY, None)
        except Exception:
            st.session_state[TZ_KEY] = ""
        st.rerun()

# If full "City, State, Country" is present → strict validate + set lat/lon + auto UTC
if GEOAPIFY_API_KEY and (not need_help) and place_raw:
    try:
        lat, lon, disp = geocode_strict(place_raw, GEOAPIFY_API_KEY)
        st.session_state[LAT_KEY] = lat
        st.session_state[LON_KEY] = lon
        st.session_state[TZ_KEY] = get_timezone_offset_simple(lat, lon)
        st.session_state.pop(TZ_ERR_KEY, None)
    except Exception as e:
        st.session_state[TZ_KEY] = ""
        st.session_state[LAT_KEY] = None
        st.session_state[LON_KEY] = None
        st.session_state[TZ_ERR_KEY] = str(e)

# Show resolved coordinates
col_lat, col_lon = st.columns(2)
with col_lat:
    st.text_input("Latitude", value=("" if st.session_state[LAT_KEY] is None else f"{st.session_state[LAT_KEY]:.6f}"), disabled=True)
with col_lon:
    st.text_input("Longitude", value=("" if st.session_state[LON_KEY] is None else f"{st.session_state[LON_KEY]:.6f}"), disabled=True)

# Error (if any)
if st.session_state.get(TZ_ERR_KEY):
    st.error(st.session_state[TZ_ERR_KEY])

st.divider()
# --- Placeholder 'Generate' button to mirror your app's flow ---
# (This doesn't generate files; it's here so you can drop-in test the POB/UTC behavior.)
if st.button("Generate Kundali"):
    if not st.session_state.get(PLACE_KEY):
        st.warning("Please select a valid Place of Birth (City, State, Country).")
    elif st.session_state.get(LAT_KEY) is None or st.session_state.get(LON_KEY) is None:
        st.warning("Could not resolve coordinates for the selected place.")
    elif not st.session_state.get(TZ_KEY):
        st.warning("UTC offset could not be determined for the selected place.")
    else:
        st.success("Place validated ✔︎  Latitude/Longitude and UTC populated.")
        st.write({
            "Place": st.session_state[PLACE_KEY],
            "Latitude": st.session_state[LAT_KEY],
            "Longitude": st.session_state[LON_KEY],
            "UTC offset (now)": st.session_state[TZ_KEY],
        })
