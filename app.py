
import os, json, re, urllib.parse, urllib.request
from dataclasses import dataclass
from typing import List, Tuple
from datetime import datetime, date, time, timedelta, timezone
from io import BytesIO

import streamlit as st

# ---------- Page setup ----------
st.set_page_config(page_title="MRIDAASTRO — Kundali", page_icon="🪔", layout="centered")
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cinzel+Decorative:wght@400;700;900&display=swap');
h1, h2, .mrida-title { font-family: "Cinzel Decorative", serif !important; }
[data-testid="stHeader"] { background: transparent; }
</style>
""", unsafe_allow_html=True)
st.markdown('<h1 class="mrida-title">MRIDAASTRO — Kundali Generator</h1>', unsafe_allow_html=True)

# ---------- Utilities ----------
def _norm(s: str) -> str:
    s = (s or "").strip().lower()
    return re.sub(r"[^a-z]", "", s)

def dms_str(deg: float, is_lat: bool) -> str:
    # 23.17028 -> 23°10'13"N ; for lon -> E/W
    d = int(deg if deg >= 0 else -deg)
    m_full = abs(deg - (d if deg >= 0 else -d)) * 60
    m = int(m_full)
    s = int(round((m_full - m) * 60))
    hemi = ("N" if deg >= 0 else "S") if is_lat else ("E" if deg >= 0 else "W")
    return f"{d}°{m:02d}′{s:02d}″{hemi}"

def geo_api_key() -> str:
    try:
        return st.secrets.get("GEOAPIFY_API_KEY", "") or ""
    except Exception:
        return os.environ.get("GEOAPIFY_API_KEY", "") or ""

# ---------- Timezone helpers (no hardcoding) ----------
try:
    from timezonefinder import TimezoneFinder
    from zoneinfo import ZoneInfo  # py3.9+
except Exception:
    TimezoneFinder = None
    ZoneInfo = None

def offset_for_datetime(lat: float, lon: float, d: date, t: time) -> float:
    """Offset hours (+/-) at the *birth moment* for given lat/lon."""
    if TimezoneFinder is None or ZoneInfo is None:
        return 0.0
    tf = TimezoneFinder()
    tzname = tf.timezone_at(lat=lat, lng=lon)
    if not tzname:
        return 0.0
    local = datetime(d.year, d.month, d.day, t.hour, t.minute, t.second, tzinfo=ZoneInfo(tzname))
    # Convert its offset to hours
    return (local.utcoffset() or timedelta()).total_seconds() / 3600.0

def format_offset(hours: float) -> str:
    sign = '+' if hours >= 0 else '-'
    total_minutes = int(round(abs(hours) * 60))
    hh, mm = divmod(total_minutes, 60)
    return f"{sign}{hh:02d}:{mm:02d}"

def parse_offset_to_hours(s: str) -> float:
    s = (s or "").strip()
    if not s: return 0.0
    m = re.match(r"^([+-])?(\d{1,2}):(\d{2})$", s)
    if not m:
        try: return float(s)
        except Exception: return 0.0
    sign = -1 if m.group(1) == '-' else 1
    h = int(m.group(2)); mnt = int(m.group(3))
    return sign * (h + mnt/60.0)

# ---------- Geocoding ----------
@dataclass
class Suggestion:
    label: str  # "City, State, Country"
    lat: float
    lon: float

@st.cache_data(show_spinner=False, ttl=3600)
def geocode_suggestions(city_only: str, api_key: str, limit: int = 12) -> List[Suggestion]:
    """Return unique 'City, State, Country' suggestions for a given city name."""
    city_only = (city_only or "").strip()
    if not city_only or not api_key:
        return []
    base = "https://api.geoapify.com/v1/geocode/search?"
    q = urllib.parse.urlencode({"text": city_only, "format": "json", "limit": max(20, limit), "apiKey": api_key})
    try:
        with urllib.request.urlopen(base + q, timeout=15) as r:
            j = json.loads(r.read().decode())
    except Exception:
        return []
    seen = set(); out: List[Suggestion] = []
    for res in (j.get("results") or []):
        city = res.get("city") or res.get("town") or res.get("village") or ""
        if _norm(city) != _norm(city_only):
            continue
        state = res.get("state") or ""
        country = res.get("country") or ""
        lat = res.get("lat"); lon = res.get("lon")
        if lat is None or lon is None:
            continue
        label = ", ".join([p for p in (city, state, country) if p])
        key = (_norm(city), _norm(state), _norm(country))
        if key in seen:
            continue
        seen.add(key); out.append(Suggestion(label, float(lat), float(lon)))
        if len(out) >= limit: break
    return out

@st.cache_data(show_spinner=False, ttl=3600)
def geocode_strict(place: str, api_key: str) -> Tuple[float, float, str]:
    """
    Strictly resolve 'City, State, Country' to (lat, lon, formatted).
    Requires 3 parts; validates city/state/country case-insensitively.
    """
    if not api_key:
        raise RuntimeError("Geo API key missing.")
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
    if _norm(city_res) != _norm(typed_city):
        fmt = res.get("formatted", "")
        if not re.search(r"\b"+re.escape(typed_city)+r"\b", fmt, flags=re.IGNORECASE):
            raise RuntimeError("Place not found. Please enter City, State, Country correctly.")
    if _norm(state_res) != _norm(typed_state):
        raise RuntimeError("Place not found. Please enter City, State, Country correctly.")
    if _norm(country_res) != _norm(typed_country):
        raise RuntimeError("Place not found. Please enter City, State, Country correctly.")
    return float(res["lat"]), float(res["lon"]), res.get("formatted", raw)

# ---------- Session keys ----------
LAT_KEY="pob_lat"; LON_KEY="pob_lon"; PLACE_KEY="place_input"; UTC_KEY="tz_display"
ERR_KEY="pob_err"; CHOICE_IDX_KEY="pob_choice_idx"; FILL_KEY="pob_fill_pending"
for k,v in [(LAT_KEY,None),(LON_KEY,None),(PLACE_KEY,""),(UTC_KEY,""),(ERR_KEY,None),(CHOICE_IDX_KEY,0),(FILL_KEY,None)]:
    st.session_state.setdefault(k,v)

# If a fill was staged by a dropdown, apply it before rendering
if st.session_state.get(FILL_KEY):
    st.session_state[PLACE_KEY] = st.session_state[FILL_KEY]
    st.session_state[FILL_KEY] = None

# ---------- UI: Place of Birth ----------
api_key = geo_api_key()
place = st.text_input("Place of Birth (City, State, Country)", key=PLACE_KEY, placeholder="e.g., Jabalpur, Madhya Pradesh, India").strip()

# If field was cleared, nuke coords + utc immediately
if not place:
    st.session_state[LAT_KEY] = None
    st.session_state[LON_KEY] = None
    st.session_state[UTC_KEY] = ""
    st.session_state[ERR_KEY] = None

need_help = bool(place) and (place.count(",") < 2)

if api_key and need_help:
    options = geocode_suggestions(place.split(",")[0], api_key, limit=12)
    if options:
        labels = ["— Select exact place —"] + [o.label for o in options]
        idx = st.selectbox("Similar city names found", options=list(range(len(labels))), index=st.session_state.get(CHOICE_IDX_KEY,0) or 0, format_func=lambda i: labels[i], key=CHOICE_IDX_KEY)
        if idx and 0 < idx <= len(options):
            choice = options[idx-1]
            st.session_state[FILL_KEY] = choice.label
            st.session_state[LAT_KEY] = choice.lat
            st.session_state[LON_KEY] = choice.lon
            # If DOB/TOB already entered, compute exact UTC; otherwise leave blank till form filled
            dob = st.session_state.get("dob_input"); tob = st.session_state.get("tob_input")
            if isinstance(dob, date) and isinstance(tob, time):
                st.session_state[UTC_KEY] = format_offset(offset_for_datetime(choice.lat, choice.lon, dob, tob))
            st.experimental_rerun()

if api_key and (not need_help) and place:
    try:
        lat, lon, _ = geocode_strict(place, api_key)
        st.session_state[LAT_KEY] = lat
        st.session_state[LON_KEY] = lon
        dob = st.session_state.get("dob_input"); tob = st.session_state.get("tob_input")
        if isinstance(dob, date) and isinstance(tob, time):
            st.session_state[UTC_KEY] = format_offset(offset_for_datetime(lat, lon, dob, tob))
        st.session_state[ERR_KEY] = None
    except Exception as e:
        st.session_state[LAT_KEY] = None
        st.session_state[LON_KEY] = None
        st.session_state[UTC_KEY] = ""
        st.session_state[ERR_KEY] = str(e)

# Show coordinates & UTC
colA, colB, colC = st.columns(3)
with colA:
    st.text_input("Latitude", value=("—" if st.session_state[LAT_KEY] is None else f"{st.session_state[LAT_KEY]:.6f}"), disabled=True)
with colB:
    st.text_input("Longitude", value=("—" if st.session_state[LON_KEY] is None else f"{st.session_state[LON_KEY]:.6f}"), disabled=True)
with colC:
    st.text_input("UTC offset (auto-detected)", key=UTC_KEY, disabled=True)

if place and st.session_state[LAT_KEY] is not None and st.session_state[LON_KEY] is not None:
    from math import isfinite
    st.caption(f"📍 Coordinates: {dms_str(st.session_state[LAT_KEY], True)} {dms_str(st.session_state[LON_KEY], False)}")

if st.session_state.get(ERR_KEY):
    st.error(st.session_state[ERR_KEY])

st.divider()

# ---------- Form for Name/DOB/TOB + Submit ----------
with st.form("form"):
    col1, col2 = st.columns(2)
    with col1:
        name = st.text_input("Name", key="name_input", placeholder="Person Name")
    with col2:
        dob = st.date_input("Date of Birth", key="dob_input", value=date(1990,1,1), min_value=date(1800,1,1), max_value=date(2100,12,31))
    col3, col4 = st.columns(2)
    with col3:
        tob = st.time_input("Time of Birth", key="tob_input", value=time(12,0), step=timedelta(minutes=1))
    with col4:
        st.write("")

    submit = st.form_submit_button("Generate Kundali")

# If DOB/TOB were entered after place, compute UTC now
if st.session_state.get(LAT_KEY) is not None and st.session_state.get(LON_KEY) is not None:
    dob = st.session_state.get("dob_input"); tob = st.session_state.get("tob_input")
    if isinstance(dob, date) and isinstance(tob, time):
        st.session_state[UTC_KEY] = format_offset(offset_for_datetime(st.session_state[LAT_KEY], st.session_state[LON_KEY], dob, tob))

# ---------- Compute chart (Swiss Ephemeris) ----------
SWISSEPHEMERIS_AVAILABLE = True
try:
    import swisseph as swe
except Exception:
    SWISSEPHEMERIS_AVAILABLE = False

def to_jd_ut(d: date, t: time, tz_hours: float) -> float:
    local_dt = datetime(d.year, d.month, d.day, t.hour, t.minute, t.second)
    ut_dt = local_dt - timedelta(hours=tz_hours)
    return swe.julday(ut_dt.year, ut_dt.month, ut_dt.day, ut_dt.hour + ut_dt.minute/60.0 + ut_dt.second/3600.0)

def normalize(deg: float) -> float:
    while deg < 0: deg += 360.0
    while deg >= 360.0: deg -= 360.0
    return deg

def mid_deg(a: float, b: float) -> float:
    # midpoint along the shorter arc from a to b
    a = normalize(a); b = normalize(b)
    d = (b - a + 540.0) % 360.0 - 180.0  # -180..+180
    return normalize(a + d/2.0)

def chalit_table(jd_ut: float, lat: float, lon: float):
    # Get tropical Placidus cusps then convert to sidereal using Lahiri
    swe.set_sid_mode(swe.SIDM_LAHIRI, 0, 0)
    flags = swe.FLG_SWIEPH  # tropical for cusps
    cusps_trop, ascmc = swe.houses_ex(jd_ut, lat, lon, b'P', flags)  # 'P' Placidus
    # Ayanamsa in degrees (tropical -> sidereal)
    ayan = swe.get_ayanamsa_ut(jd_ut)
    cusps_sid = [normalize(c - ayan) for c in cusps_trop[:12]]

    # BhavMadhya (mid-bhav) is the sidereal cusp itself
    mid_bhav = cusps_sid[:12]

    # BhavBegin by Sripati: midpoint of previous and current cusp
    begin = []
    for i in range(12):
        prev_cusp = cusps_sid[(i - 1) % 12]
        curr_cusp = cusps_sid[i]
        begin.append(mid_deg(prev_cusp, curr_cusp))

    # Rashi sign for begin and mid
    signs = ["Aries","Taurus","Gemini","Cancer","Leo","Virgo","Libra","Scorpio","Sagittarius","Capricorn","Aquarius","Pisces"]
    def sign_name(d): return signs[int(normalize(d)//30)]

    rows = []
    for i in range(12):
        rows.append({
            "Bhav": i+1,
            "Rashi": sign_name(begin[i]),
            "BhavBegin": begin[i],
            "RashiMid": sign_name(mid_bhav[i]),
            "MidBhav": mid_bhav[i],
        })
    return rows

def deg_to_dms_str(d: float) -> str:
    d = normalize(d)
    deg = int(d)
    m_full = (d - deg) * 60
    minute = int(m_full)
    sec = int(round((m_full - minute) * 60))
    if sec == 60: sec = 0; minute += 1
    if minute == 60: minute = 0; deg = (deg + 1) % 360
    return f"{deg:02d}.{minute:02d}.{sec:02d}"

if submit:
    errs = []
    name = (st.session_state.get("name_input") or "").strip()
    if not name: errs.append("Name is required.")
    place_final = (st.session_state.get(PLACE_KEY) or "").strip()
    if not place_final: errs.append("Place of Birth is required.")
    lat = st.session_state.get(LAT_KEY); lon = st.session_state.get(LON_KEY)
    if lat is None or lon is None: errs.append("Coordinates not resolved.")
    tz_str = (st.session_state.get(UTC_KEY) or "").strip()
    if not tz_str: errs.append("UTC offset not set.")
    if errs:
        for e in errs: st.error(e)
    else:
        if not SWISSEPHEMERIS_AVAILABLE:
            st.error("Swiss Ephemeris is not installed. Add 'swisseph' to requirements.txt.")
        else:
            tz_hours = parse_offset_to_hours(tz_str)
            jd_ut = to_jd_ut(st.session_state["dob_input"], st.session_state["tob_input"], tz_hours)
            rows = chalit_table(jd_ut, float(lat), float(lon))

            # Show a preview table
            import pandas as pd
            df = pd.DataFrame([
                {
                    "Bhav": r["Bhav"],
                    "Rashi": r["Rashi"],
                    "BhavBegin": deg_to_dms_str(r["BhavBegin"]),
                    "RashiMid": r["RashiMid"],
                    "MidBhav": deg_to_dms_str(r["MidBhav"]),
                } for r in rows
            ])
            with st.expander("Chalit (Sripati/Placidus) — Preview", expanded=True):
                st.dataframe(df, use_container_width=True)

            st.success("Kundali computed.")
