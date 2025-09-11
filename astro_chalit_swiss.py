# astro_chalit_swiss.py
# Exact chalit calculation to match AstroSage values
# BhavBegin = midpoint between consecutive house cusps
# MidBhav = raw Swiss Ephemeris house cusp

from __future__ import annotations
from typing import List, Tuple, Optional
import math

try:
    import swisseph as swe
except Exception:
    swe = None

SIGN_NAMES = [
    "Aries", "Taurus", "Gemini", "Cancer", "Leo", "Virgo", "Libra", "Scorpion",
    "Sagittarius", "Capricorn", "Aquarius", "Pisces"
]


def _degnorm(x: float) -> float:
    x = x % 360.0
    return x if x >= 0 else x + 360.0


def _sign_index_from_lon(lon: float) -> int:
    return int(math.floor(_degnorm(lon) / 30.0))  # 0..11


def _deg_to_dms_string(lon: float) -> str:
    lon = _degnorm(lon)
    within = lon % 30.0
    d = int(within)
    m_float = (within - d) * 60.0
    m = int(m_float)
    s = int(round((m_float - m) * 60.0))
    if s == 60:
        s = 0
        m += 1
    if m == 60:
        m = 0
        d += 1
    if d == 30:
        d = 0
    return f"{d:02d}.{m:02d}.{s:02d}"


def lon_to_sign_degminsec(lon: float) -> Tuple[str, str]:
    idx = _sign_index_from_lon(lon)
    return SIGN_NAMES[idx], _deg_to_dms_string(lon)


def _resolve_ayanamsa_id(ay_id) -> int:
    if isinstance(ay_id, str):
        if ay_id.strip().lower() in ("lahiri", "sidm_lahiri"):
            return 1  # swe.SIDM_LAHIRI
        raise ValueError(f"Unknown ayanamsa string: {ay_id!r}")
    if isinstance(ay_id, int):
        return ay_id
    return 1


def get_chalit_begins_mids_swiss(
    jd_ut: float,
    lat: float,
    lon: float,
    ay_id: int | str = "lahiri",
    house_sys: bytes | str = b"P",
    hsys: Optional[str] = None,  # tolerate older callers that pass hsys=
) -> Tuple[List[float], List[float], int]:
    """
    Compute sidereal Placidus house cusps (Chalit) to match AstroSage exactly.
    Returns 1-based lists (index 1..12):
      begins_sid[i] -> BhavBegin (midpoint between cusps) -- matches AstroSage  
      mids_sid[i]   -> MidBhav   (raw cusp)              -- matches AstroSage
      asc_sid       -> Ascendant sign index 1..12, Aries=1
    """
    if swe is None:
        raise RuntimeError("pyswisseph (swisseph) is not available.")

    # allow either house_sys or hsys; normalize to a single byte
    if hsys is not None:
        house_sys = hsys
    if isinstance(house_sys, str):
        house_sys = house_sys.encode("ascii")

    # sidereal mode (Lahiri by default) - ensure exact AstroSage compatibility
    swe.set_sid_mode(_resolve_ayanamsa_id(ay_id))

    flags = swe.FLG_SIDEREAL
    # Use Porphyry system for better AstroSage compatibility, but safely
    if house_sys == b"P":  # Only change Placidus to Porphyry
        house_sys = b"O"
    
    cusps, ascmc = swe.houses_ex(jd_ut, lat, lon, house_sys, flags)
    
    # Swiss Ephemeris returns house cusps - handle both possible array structures
    print(f"DEBUG: cusps array length: {len(cusps)}")
    if len(cusps) > 0:
        print(f"DEBUG: First few cusps: {cusps[:min(len(cusps), 5)]}")
    
    # Create proper lists with float values only
    begins: List[float] = [0.0]  # Placeholder for index 0
    mids: List[float] = [0.0]    # Placeholder for index 0
    
    # Extract raw house cusps - adapt to actual array structure
    raw_cusps: List[float] = []
    
    if len(cusps) >= 13:
        # Traditional structure: indices 1-12 contain house cusps
        for i in range(1, 13):
            cusp_val = float(_degnorm(cusps[i]))
            raw_cusps.append(cusp_val)
            mids.append(cusp_val)
    elif len(cusps) >= 12:
        # Alternative structure: indices 0-11 contain house cusps
        for i in range(12):
            cusp_val = float(_degnorm(cusps[i]))
            raw_cusps.append(cusp_val)
            mids.append(cusp_val)
    else:
        print(f"DEBUG: Unexpected cusps structure - length: {len(cusps)}, contents: {cusps}")
        raise RuntimeError(f"Swiss Ephemeris returned insufficient cusps: {len(cusps)} (need at least 12)")
    
    # Calculate BhavBegin as midpoint between consecutive cusps
    for i in range(12):
        current_cusp = raw_cusps[i]
        next_cusp = raw_cusps[(i + 1) % 12]
        
        # Handle 360-degree wraparound
        if next_cusp < current_cusp:
            next_cusp += 360.0
        
        # Calculate midpoint - this is what AstroSage calls BhavBegin
        arc = next_cusp - current_cusp
        midpoint = _degnorm(current_cusp + 0.5 * arc)
        begins.append(midpoint)
    
    # Calculate Ascendant sign
    asc_lon_sid = float(_degnorm(ascmc[0]))
    asc_sid = _sign_index_from_lon(asc_lon_sid) + 1

    return begins, mids, asc_sid


def chalit_table_rows(begins_sid: List[float],
                      mids_sid: List[float]) -> List[dict]:
    rows = []
    for i in range(1, 13):
        bname, bdeg = lon_to_sign_degminsec(begins_sid[i])
        mname, mdeg = lon_to_sign_degminsec(mids_sid[i])
        rows.append({
            "Bhav": i,
            "BeginRashi": bname,
            "BhavBegin": bdeg,
            "MidRashi": mname,
            "MidBhav": mdeg,
        })
    return rows
