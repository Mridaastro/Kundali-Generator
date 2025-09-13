
import streamlit as st
from datetime import datetime, date, time
from io import BytesIO
import zipfile

# Optional Swiss Ephemeris (if available)
try:
    import swisseph as swe
    HAVE_SWE = True
except Exception:
    HAVE_SWE = False

st.set_page_config(page_title="MridaAstro", layout="wide")

st.markdown("""
<style>
[data-testid="stExpander"] > summary { font-weight: 700 !important; }
[data-testid="stExpander"] { max-width: 800px; margin-left: auto; margin-right: auto; }
@media (min-width: 1200px){ [data-testid="stExpander"] { max-width: 50vw; } }
</style>
""", unsafe_allow_html=True)

def _clear_place_state():
    for k in ("pob_query","pob_options","pob_choice","pob_coords","utc_offset","tz_name"):
        st.session_state.pop(k, None)

def _on_place_query_change():
    q = st.session_state.get("pob_query","").strip()
    if not q:
        _clear_place_state()

def export_docx_bytes(paragraphs) -> bytes:
    try:
        from docx import Document
        doc = Document()
        for p in paragraphs:
            doc.add_paragraph(p)
        buf = BytesIO()
        doc.save(buf)
        buf.seek(0)
        data = buf.getvalue()
        if not data.startswith(b"PK"):
            raise RuntimeError("DOCX serialization failed")
        with zipfile.ZipFile(BytesIO(data)) as zf:
            assert "[Content_Types].xml" in zf.namelist()
        return data
    except Exception as e:
        st.error(f"Could not create DOCX: {e}")
        return b""

def setup_sidereal_true_lahiri():
    if not HAVE_SWE: return
    try:
        swe.set_sid_mode(swe.SIDM_LAHIRI, 0, 0)
    except Exception:
        pass

def compute_nodes_true_lahiri(jd_ut: float):
    if not HAVE_SWE: return (None, None)
    setup_sidereal_true_lahiri()
    FLG = swe.FLG_SWIEPH | swe.FLG_SIDEREAL
    xx, _ = swe.calc_ut(jd_ut, 11, FLG)   # 11 = TRUE_NODE
    rahu = xx[0] % 360.0
    ketu = (rahu + 180.0) % 360.0
    return rahu, ketu

# Header
st.markdown("<h1 style='text-align:center;margin:0;'>MRIDAASTRO</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align:center;font-size:1.3rem;margin-top:0;'>In the light of divine, let your soul journey shine</p>", unsafe_allow_html=True)
st.markdown("<hr style='width:200px;margin:8px auto 24px auto;border:0;border-top:3px solid #000;'>", unsafe_allow_html=True)

# Inputs
col1, col2 = st.columns(2)
with col1:
    name = st.text_input("Name *", key="name")
with col2:
    st.text_input("Place of Birth (City, State, Country) *", key="pob_query", on_change=_on_place_query_change)

dropdown_ph = st.empty()

def _demo_suggest(q: str):
    base = [
        "Mandla, Madhya Pradesh, India",
        "Mumbai, Maharashtra, India",
        "Mandla, Montana, United States",
    ]
    ql = q.lower()
    return [x for x in base if ql in x.lower()]

q = st.session_state.get("pob_query", "").strip()
if q:
    options = _demo_suggest(q)
    st.session_state["pob_options"] = options
    if options:
        with dropdown_ph.container():
            st.selectbox("Select Place (City, State, Country)", options=options, key="pob_choice")
    else:
        dropdown_ph.empty()
else:
    _clear_place_state()
    dropdown_ph.empty()

col3, col4 = st.columns(2)
with col3:
    dob = st.date_input("Date of Birth *", value=date.today())
with col4:
    tob = st.time_input("Time of Birth *", value=time(0, 0))

utc_value = st.session_state.get("utc_offset", "")
st.text_input("UTC offset (enter Place of Birth first)", value=utc_value, key="utc_field")

# Centered button
btn_l, btn_m, btn_r = st.columns([1,1,1])
with btn_m:
    generate_clicked = st.button("Generate Kundali", use_container_width=True, key="gen_btn")

# Advanced settings
a_l, a_m, a_r = st.columns([1,2,1])
with a_m:
    with st.expander("🎨 Advanced Settings", expanded=False):
        color_options = {
            "Orange (Default)": "#FF6600",
            "Pink Lace": "#F0D7F5",
            "Mint": "#99EDC3",
            "Coral": "#FE7D6A",
            "Rose": "#FC94AF",
            "Pink": "#FF0049",
            "Green": "#48FF00",
            "Purple": "#A020F0",
            "Yellow": "#FFB700",
        }
        selected_color = st.selectbox(
            "Color Scheme",
            options=list(color_options.keys()),
            index=0,
            key="color_scheme",
            label_visibility="collapsed"
        )
        st.caption("Advanced settings appear here.")

if generate_clicked:
    if not name or not q:
        st.error("Please fill Name and Place of Birth.")
    else:
        paras = [
            f"Name: {name}",
            f"Place: {st.session_state.get('pob_choice') or q}",
            f"DOB: {dob.isoformat()}",
            f"TOB: {tob.strftime('%H:%M')}",
        ]
        if HAVE_SWE:
            try:
                jd_ut = swe.julday(dob.year, dob.month, dob.day, tob.hour + tob.minute/60.0)
                rahu, ketu = compute_nodes_true_lahiri(jd_ut)
                if rahu is not None:
                    paras.append(f"Rahu (sid): {rahu:.6f}°  Ketu (sid): {ketu:.6f}°  [True+Lahiri]")
            except Exception as e:
                st.warning(f"SWE demo calc skipped: {e}")
        data = export_docx_bytes(paras)
        if data:
            st.download_button(
                "⬇️ Download Kundali (DOCX)",
                data=data,
                file_name=f"{name or 'Kundali'}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
            st.success("Kundali generated.")
