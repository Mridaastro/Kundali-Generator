# chalit_kundali_renderer.py
# Build "house_planets" for the existing kundali VML, driven by Chalit (Bhava) houses.
# This first pass focuses on correct house assignment + visual annotations:
#  - If a planet's rāśi-house != Chalit-house, show a corner-shift arrow (↗ for forward, ↙ for backward) with integer ° shift.
#  - For close pairs (≤ 6°) *within the SAME Chalit house*, insert a tiny "↔N°" marker after the first planet.
#  - Returned structure is compatible with app's kundali_with_planets(...) that expects {house: [{'txt': 'Su', 'flags': {...}}, ...]}
# Author: ChatGPT

from __future__ import annotations
import math
from typing import Dict, List, Tuple

# Planet code order expected in the app (example): 'Su','Mo','Ma','Me','Ju','Ve','Sa','Ra','Ke'
# We will pass the same codes through in labels.
ARROW_FWD = "↗"
ARROW_BWD = "↙"
ARROW_PAIR = "↔"


def _norm(x: float) -> float:
    x %= 360.0
    if x < 0:
        x += 360.0
    return x


def _delta_deg(a: float, b: float) -> float:
    """Smallest absolute separation on the circle, in degrees [0,180]."""
    d = abs((_norm(a) - _norm(b)) % 360.0)
    if d > 180.0:
        d = 360.0 - d
    return d


def rasi_house_of_lon(lon_sid: float, lagna_sign: int) -> int:
    """
    Map a sidereal longitude to the rāśi house number (1..12) using lagna_sign as the 1st house sign.
    lagna_sign is 1..12 for Aries..Pisces.
    """
    sign = int(_norm(lon_sid) // 30.0) + 1  # 1..12
    house = ((sign - lagna_sign) % 12) + 1  # 1..12
    return house


def house_index_for_lon_chalit(lon_sid: float, begins_sid: List[float]) -> int:
    """
    Given longitude and house begins (list of 12, for H1..H12), return the house index 1..12
    where lon lies in [begin_i, begin_{i+1}) moving forward on the circle.
    """
    lon = _norm(lon_sid)
    # Ensure begins is 12-length list of 0..360, arbitrary starting point (H1 begin)
    # We'll walk through each interval.
    for i in range(12):
        a = begins_sid[i]
        b = begins_sid[(i + 1) % 12]
        # interval [a, b) forward on circle
        if (b - a) % 360 == 0:
            # degenerate, treat as false
            continue
        # compute forward distance from a to lon and a to b
        da = (lon - a) % 360.0
        dab = (b - a) % 360.0
        if 0.0 <= da < dab:
            return i + 1  # 1..12
    # Fallback: if numerics miss due to boundary, snap to last begin <= lon in forward sense
    # Find begin that is the last <= lon in circular order
    idx = max(range(12),
              key=lambda k: (_norm(lon - begins_sid[k]) % 360.0 < 180.0, -(
                  (_norm(lon - begins_sid[k])) % 360.0)))
    return idx + 1


def shift_degree_to_boundary(lon_sid: float, begins_sid: List[float],
                             target_house: int, direction: str) -> int:
    """
    For a shifted planet, compute the integer degree offset shown near the arrow:
      - For forward shift: degrees from begin[target_house] to planet.
      - For backward shift: degrees from planet to begin[target_house].
    Return rounded integer degrees (no minutes) as requested.
    """
    a = begins_sid[(target_house - 1) % 12]
    lon = _norm(lon_sid)
    if direction == "forward":
        d = (lon - a) % 360.0
    else:
        d = (a - lon) % 360.0
    if d > 180.0:
        d = 360.0 - d
    return int(round(d))


def house_planets_from_chalit(
        sidelons: Dict[str, float], lagna_sign: int,
        begins_sid: List[float]) -> Dict[int, List[dict]]:
    """
    Build the per-house planet label lists using Chalit assignment and the user's visual spec.
    Inputs:
      - sidelons: dict planet_code -> sidereal longitude in degrees (0..360)
      - lagna_sign: int 1..12 (Aries..Pisces) sign for lagna
      - begins_sid: list of 12 floats (H1..H12) for Chalit begins (sidereal)
    Returns:
      - dict mapping house (1..12) -> list of {'txt': label_string, 'flags': {}}
    """
    # First pass: find chalit house and rasi house for each planet
    planet_info = []
    for P, lon in sidelons.items():
        rasi_h = rasi_house_of_lon(lon, lagna_sign)
        chalit_h = house_index_for_lon_chalit(lon, begins_sid)
        if chalit_h == rasi_h:
            lab = f"{P}"
            info = {
                "planet": P,
                "lon": lon,
                "house": chalit_h,
                "label": lab,
                "shift": None
            }
        else:
            # Decide forward/backward relative to rasi house → chalit house
            # In North Indian style, houses increase clockwise as per our indexing.
            # We'll define "forward" as moving to (rasi_h + 1) mod 12; if chalit == rasi-1, it's backward.
            diff = (chalit_h - rasi_h) % 12
            if diff == 1:
                direction = "forward"
                arrow = ARROW_FWD
            elif diff == 11:
                direction = "backward"
                arrow = ARROW_BWD
            else:
                # Larger jumps are rare; pick the nearest direction by minimal step
                if diff <= 6:
                    direction = "forward"
                    arrow = ARROW_FWD
                else:
                    direction = "backward"
                    arrow = ARROW_BWD
            degs = shift_degree_to_boundary(lon, begins_sid, chalit_h,
                                            direction)
            lab = f"{P} {arrow}{degs}°"
            info = {
                "planet": P,
                "lon": lon,
                "house": chalit_h,
                "label": lab,
                "shift": direction
            }
        planet_info.append(info)

    # Group by chalit house
    by_house: Dict[int, List[dict]] = {h: [] for h in range(1, 13)}
    for it in planet_info:
        by_house[it["house"]].append({"txt": it["label"], "flags": {}})

    # Second pass: add close-pair annotations (≤6°) within each house.
    for h in range(1, 13):
        items = [x for x in planet_info if x["house"] == h]
        # sort by longitude along the interval beginning (optional, for stable pairing)
        items.sort(key=lambda k: k["lon"])
        used = set()
        for i in range(len(items)):
            if i in used:
                continue
            for j in range(i + 1, len(items)):
                if j in used:
                    continue
                sep = _delta_deg(items[i]["lon"], items[j]["lon"])
                if sep <= 6.0:
                    # insert a small arrow label right after the first planet
                    ins_idx = 0
                    # find position of items[i] inside by_house[h]
                    for pos, entry in enumerate(by_house[h]):
                        if entry["txt"].startswith(items[i]["planet"]):
                            ins_idx = pos + 1
                            break
                    by_house[h].insert(ins_idx, {
                        "txt": f"{ARROW_PAIR}{int(round(sep))}°",
                        "flags": {}
                    })
                    used.add(i)
                    used.add(j)
                    break

    return by_house
