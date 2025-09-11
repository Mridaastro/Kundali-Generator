
import streamlit as st
import json, re, urllib.parse, urllib.request
from dataclasses import dataclass
from typing import List, Tuple
from datetime import datetime, date, time, timedelta, timezone
import os
from io import BytesIO

# ============ Page config ============
st.set_page_config(page_title="MRIDAASTRO — Kundali", page_icon="🪔", layout="centered")
st.title("MRIDAASTRO — Kundali Generator")

# ---------- optional background (if assets exist) ----------
def _apply_bg():
    try:
        import base64
        from pathlib import Path as _P
        for p in [_P("assets/ganesha_bg.png"), _P("assets/login_bg.png"), _P("assets/bg.png")]:
            if p.exists():
                b64 = base64.b64encode(p.read_bytes()).decode()
                st.markdown(f"""
                <style>
                [data-testid="stAppViewContainer"] {{
                    background: url('data:image/png;base64,{b64}') no-repeat center top fixed;
                    background-size: cover;
                }}
                [data-testid="stHeader"] {{
                    background: transparent;
                }}
                </style>
                """, unsafe_allow_html=True)
                break
    except Exception:
        pass

_apply_bg()

# ============ Secrets / API keys ============
GEOAPIFY_API_KEY = ""
try:
    GEOAPIFY_API_KEY = st.secrets.get("GEOAPIFY_API_KEY", "")
except Exception:
    pass
if not GEOAPIFY_API_KEY:
    GEOAPIFY_API_KEY = os.environ.get("GEOAPIFY_API_KEY", "")

# ============ Timezone helpers ============
try:
    from timezonefinder import TimezoneFinder
    from zoneinfo import ZoneInfo
except Exception:
    TimezoneFinder = None
    ZoneInfo = None

def get_timezone_offset_simple(lat: float, lon: float) -> str:
    """Return UTC offset like '+5.5' based on *current* rules (DOB-independent)."""
    if TimezoneFinder is None or ZoneInfo is None:
        return ""
    tf = TimezoneFinder()
    tzname = tf.timezone_at(lat=lat, lng=lon)
    if not tzname:
        return ""
    now_utc = datetime.now(timezone.utc)
    local = now_utc.astimezone(ZoneInfo(tzname))
    offset_hours = local.utcoffset().total_seconds() / 3600.0
    if abs(offset_hours - round(offset_hours)) < 1e-9:
        return f"{int(round(offset_hours)):+d}"
    s = f"{offset_hours:+.2f}".rstrip("0").rstrip(".")
    return s

# ============ POB helpers ============
def _norm(s: str) -> str:
    s = (s or "").strip().lower()
    return re.sub(r"[^a-z]", "", s)

@dataclass
class Suggestion:
    label: str  # "City, State, Country"
    lat: float
    lon: float

@st.cache_data(show_spinner=False, ttl=3600)
def geocode_suggestions(typed: str, api_key: str, limit: int = 7) -> List[Suggestion]:
    """Return suggestions for ambiguous city names (city/town/village only)."""
    out: List[Suggestion] = []
    typed = (typed or "").strip()
    if not typed or not api_key:
        return out
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
        seen.add(key)
        out.append(Suggestion(label=label, lat=float(lat), lon=float(lon)))
        if len(out) >= limit:
            break
    return out

@st.cache_data(show_spinner=False, ttl=3600)
def geocode_strict(place: str, api_key: str) -> Tuple[float, float, str]:
    """
    STRICT geocoding:
    - Requires 'City, State, Country' (3 parts).
    - Accept only if city/state/country match (case-insensitive).
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

# ===== Session keys =====
LAT_KEY = "pob_lat"
LON_KEY = "pob_lon"
TZ_KEY = "tz_input"              # internal state (string like +5.5)
PLACE_KEY = "place_input"
CHOICE_IDX_KEY = "place_choice_idx"
CHOICE_OPTIONS_KEY = "place_choice_options"
FILL_REQUEST_KEY = "place_fill_request"
ERR_KEY = "tz_error_msg"

for k, v in [(LAT_KEY,None), (LON_KEY,None), (TZ_KEY,""), (PLACE_KEY,""), (CHOICE_IDX_KEY,0), (CHOICE_OPTIONS_KEY, []), (FILL_REQUEST_KEY, None), (ERR_KEY, None)]:
    st.session_state.setdefault(k, v)

# ===== Apply any pending programmatic fill BEFORE creating the widget =====
if st.session_state.get(FILL_REQUEST_KEY):
    st.session_state[PLACE_KEY] = st.session_state[FILL_REQUEST_KEY]
    st.session_state[FILL_REQUEST_KEY] = None

# ===== Top row: Name + UTC (UTC is read-only and decoupled from state key) =====
col1, col2 = st.columns(2)
with col1:
    name = st.text_input("Name", key="name_input", placeholder="Person Name")
with col2:
    st.text_input("UTC offset", value=st.session_state[TZ_KEY], key="tz_display", help="Auto-detected from Place of Birth", disabled=True)

# ===== Place of Birth (with suggestions) =====
place_typed = st.text_input("Place of Birth (City, State, Country)", key=PLACE_KEY, placeholder="e.g., Jabalpur, Madhya Pradesh, India").strip()
need_help = bool(place_typed) and (place_typed.count(",") < 2)

def _apply_choice_callback():
    idx = st.session_state.get(CHOICE_IDX_KEY, 0) or 0
    options: List[Suggestion] = st.session_state.get(CHOICE_OPTIONS_KEY, [])
    if idx > 0 and 0 < idx <= len(options):
        chosen = options[idx - 1]
        st.session_state[FILL_REQUEST_KEY] = chosen.label
        st.session_state[LAT_KEY] = chosen.lat
        st.session_state[LON_KEY] = chosen.lon
        try:
            st.session_state[TZ_KEY] = get_timezone_offset_simple(chosen.lat, chosen.lon)
            st.session_state[ERR_KEY] = None
        except Exception:
            st.session_state[TZ_KEY] = ""

if GEOAPIFY_API_KEY and need_help:
    suggestions = geocode_suggestions(place_typed, GEOAPIFY_API_KEY, limit=7)
    if suggestions:
        st.session_state[CHOICE_OPTIONS_KEY] = suggestions
        labels = ["— Select exact place —"] + [s.label for s in suggestions]
        st.selectbox("Similar city names found",
                     options=list(range(len(labels))),
                     format_func=lambda i: labels[i],
                     index=st.session_state.get(CHOICE_IDX_KEY, 0) or 0,
                     key=CHOICE_IDX_KEY,
                     on_change=_apply_choice_callback)

# If full 'City, State, Country' typed → strict validate + set lat/lon + auto UTC
if GEOAPIFY_API_KEY and (not need_help) and place_typed:
    try:
        lat, lon, disp = geocode_strict(place_typed, GEOAPIFY_API_KEY)
        st.session_state[LAT_KEY] = lat
        st.session_state[LON_KEY] = lon
        st.session_state[TZ_KEY] = get_timezone_offset_simple(lat, lon)
        st.session_state[ERR_KEY] = None
    except Exception as e:
        st.session_state[TZ_KEY] = ""
        st.session_state[LAT_KEY] = None
        st.session_state[LON_KEY] = None
        st.session_state[ERR_KEY] = str(e)

# ===== Show resolved coordinates =====
col_lat, col_lon = st.columns(2)
with col_lat:
    st.text_input("Latitude", value=("" if st.session_state[LAT_KEY] is None else f"{st.session_state[LAT_KEY]:.6f}"), disabled=True)
with col_lon:
    st.text_input("Longitude", value=("" if st.session_state[LON_KEY] is None else f"{st.session_state[LON_KEY]:.6f}"), disabled=True)

if st.session_state.get(ERR_KEY):
    st.error(st.session_state[ERR_KEY])

# ===== Row: DOB / TOB =====
row2c1, row2c2 = st.columns(2)
with row2c1:
    dob = st.date_input("Date of Birth", key="dob_input", value=date(1990,1,1), min_value=date(1800,1,1), max_value=date(2100,12,31))
with row2c2:
    tob = st.time_input("Time of Birth", key="tob_input", value=time(12,0), step=timedelta(minutes=1))

st.divider()

# ============ Swiss Ephemeris ============
SWISSEPHEMERIS_AVAILABLE = True
try:
    import swisseph as swe
except Exception:
    SWISSEPHEMERIS_AVAILABLE = False

def parse_tz_float(s: str) -> float:
    s = (s or "").strip()
    if not s: return 0.0
    try:
        return float(s.replace("UTC","").replace("+"," +").replace("−","-").replace("–","-").split()[-1])
    except Exception:
        try:
            return float(s)
        except Exception:
            return 0.0

def to_jd_ut(d: date, t: time, tz_offset_hours: float) -> float:
    local_dt = datetime(d.year, d.month, d.day, t.hour, t.minute, t.second)
    ut_dt = local_dt - timedelta(hours=tz_offset_hours)
    return swe.julday(ut_dt.year, ut_dt.month, ut_dt.day, ut_dt.hour + ut_dt.minute/60.0 + ut_dt.second/3600.0)

def compute_chart(jd_ut: float, lat: float, lon: float):
    """Return dict with ascendant, cusps, and planet longitudes (sidereal/Lahiri)."""
    if not SWISSEPHEMERIS_AVAILABLE:
        return {"error": "Swiss Ephemeris not installed"}
    flags = swe.FLG_SWIEPH | swe.FLG_SIDEREAL
    swe.set_sid_mode(swe.SIDM_LAHIRI, 0, 0)
    try:
        cusps, ascmc = swe.houses_ex(jd_ut, lat, lon, b'P', flags)
        asc = ascmc[0]
    except Exception:
        cusps, ascmc, asc = [], [], None
    planet_codes = {
        'Su': swe.SUN, 'Mo': swe.MOON, 'Me': swe.MERCURY, 'Ve': swe.VENUS,
        'Ma': swe.MARS, 'Ju': swe.JUPITER, 'Sa': swe.SATURN, 'Ra': swe.MEAN_NODE
    }
    positions = {}
    for k, pcode in planet_codes.items():
        try:
            elon, latp, dist, speed = swe.calc_ut(jd_ut, pcode, flags)
            positions[k] = elon % 360.0
        except Exception:
            positions[k] = None
    if positions.get('Ra') is not None:
        positions['Ke'] = (positions['Ra'] + 180.0) % 360.0
    else:
        positions['Ke'] = None
    return {"ascendant": asc, "cusps": list(cusps) if cusps else [], "positions": positions}

# ============ DOCX export ============
def make_docx(name: str, place: str, lat: float, lon: float, tz_str: str, jd_ut: float, chart: dict) -> bytes:
    from docx import Document
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()
    h = doc.add_heading(f"Kundali — {name}", 0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph()
    r = p.add_run(f"Place: {place}\nLatitude: {lat:.6f}   Longitude: {lon:.6f}   UTC: {tz_str}\nJD(UT): {jd_ut:.5f}")
    r.font.size = Pt(11)

    doc.add_paragraph(" ")

    tbl = doc.add_table(rows=1, cols=3)
    hdr = tbl.rows[0].cells
    hdr[0].text = "Body"; hdr[1].text = "Longitude (°)"; hdr[2].text = "Sign"
    def sign_name(deg):
        idx = int((deg % 360.0) // 30) + 1
        return ["Aries","Taurus","Gemini","Cancer","Leo","Virgo","Libra","Scorpio","Sagittarius","Capricorn","Aquarius","Pisces"][idx-1]

    for body, deg in chart["positions"].items():
        row = tbl.add_row().cells
        row[0].text = body
        row[1].text = ("—" if deg is None else f"{deg:.2f}")
        row[2].text = ("—" if deg is None else sign_name(deg))

    if chart.get("ascendant") is not None:
        doc.add_paragraph(f"Ascendant: {chart['ascendant']:.2f}° ({sign_name(chart['ascendant'])})")
    if chart.get("cusps"):
        doc.add_paragraph("House cusps (°): " + ", ".join(f"{c:.2f}" for c in chart["cusps"][:12]))

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# ============ Generate button ============
colg1, colg2, colg3 = st.columns([1,1,1])
with colg2:
    generate_clicked = st.button("Generate Kundali")

if generate_clicked:
    # Validation
    errors = []
    if not name.strip():
        errors.append("Name is required.")
    place = st.session_state.get(PLACE_KEY, "").strip()
    if not place:
        errors.append("Place of Birth is required (City, State, Country).")
    lat = st.session_state.get(LAT_KEY, None)
    lon = st.session_state.get(LON_KEY, None)
    if lat is None or lon is None:
        errors.append("Latitude/Longitude could not be resolved for the selected place.")
    tz_str = (st.session_state.get(TZ_KEY, "") or "").strip()
    if not tz_str:
        errors.append("UTC offset not set. (It should auto-populate after selecting the place.)")

    if errors:
        for e in errors:
            st.error(e)
    else:
        tz_hours = parse_tz_float(tz_str)
        if not SWISSEPHEMERIS_AVAILABLE:
            st.error("Swiss Ephemeris is not installed in this environment. Please add 'swisseph' to requirements.txt.")
        else:
            with st.spinner("Computing chart..."):
                try:
                    jd_ut = to_jd_ut(dob, tob, tz_hours)
                    chart = compute_chart(jd_ut, float(lat), float(lon))
                    if chart.get("error"):
                        st.error(chart["error"])
                    else:
                        st.success("Kundali computed successfully.")
                        st.json({
                            "JD_UT": jd_ut,
                            "Ascendant_deg": chart.get("ascendant"),
                            "Positions": chart.get("positions")
                        })
                        # Build DOCX
                        content = make_docx(name.strip(), place, float(lat), float(lon), tz_str, jd_ut, chart)
                        # Prefer PersonName_Horoscope per your earlier ask
                        base = name.strip().replace(" ", "_") or "Kundali"
                        file_name = f"{base}_Horoscope.docx"
                        st.download_button("Download Horoscope (DOCX)", data=content, file_name=file_name, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                except Exception as e:
                    st.error(f"Computation failed: {e}")
