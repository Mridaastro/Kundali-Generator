
# chalit_kundali_vml.py
# Exact on-diagram placement for Chalit chart in DOCX (VML).
# - Planet markers placed along within-house baseline by BhavBegin→BhavEnd fraction
# - Double-headed arrows for ≤6° pairs in same house (degrees-only labels)
# - Forward/backward shift arrows across house border from rashi corner (degrees-only labels)

from math import fmod
from typing import Dict, List, Tuple
from docx.oxml import parse_xml

HN_ABBR = {'Su': 'सू', 'Mo': 'चं', 'Ma': 'मं', 'Me': 'बु', 'Ju': 'गु', 'Ve': 'शु', 'Sa': 'श', 'Ra': 'रा', 'Ke': 'के'}


def _safe_get_label(code, planet_labels):
    base = HN_ABBR.get(code, code)
    if isinstance(planet_labels, dict) and code in planet_labels and planet_labels[code]:
        return str(planet_labels[code])
    return base

def _flag(flags, code, key):
    try:
        return bool(flags.get(code, {}).get(key, False))
    except Exception:
        return False

def _n360(x: float) -> float:
    x = fmod(x, 360.0)
    return x if x >= 0 else x + 360.0

def _fwd_arc(a: float, b: float) -> float:
    return (b - a) % 360.0

def _deg_only(x: float) -> int:
    return int(round(x))

def _sign_index(lon_sid: float) -> int:
    return int(_n360(lon_sid) // 30) + 1  # 1..12

def _house_rasi(lagna_sign: int, lon_sid: float) -> int:
    sign = _sign_index(lon_sid)
    return ((sign - lagna_sign) % 12) + 1

def _house_chalit(begins: List[float], lon_sid: float) -> int:
    L = _n360(lon_sid)
    for n in range(1, 13):
        a = begins[n]
        b = begins[1] if n == 12 else begins[n+1]
        if _fwd_arc(a, L) < _fwd_arc(a, b):
            return n
    return 12

def _arc_fraction(start: float, end: float, pt: float) -> float:
    total = _fwd_arc(start, end) or 1e-9
    part  = _fwd_arc(start, pt)
    t = part / total
    return 0.0 if t < 0 else (1.0 if t > 1 else t)

def _house_polys(S: float):
    TM = (S/2, 0); RM = (S, S/2); BM = (S/2, S); LM = (0, S/2)
    O  = (S/2, S/2)
    P_lt = (S/4, S/4); P_rt = (3*S/4, S/4); P_rb = (3*S/4, 3*S/4); P_lb = (S/4, 3*S/4)
    houses = {
        1:  [TM, P_rt, O, P_lt],
        2:  [(0, 0), TM, P_lt],
        3:  [(0, 0), LM, P_lt],
        4:  [LM, O, P_lt, P_lb],
        5:  [LM, (0, S), P_lb],
        6:  [(0, S), BM, P_lb],
        7:  [BM, P_rb, O, P_lb],
        8:  [BM, (S, S), P_rb],
        9:  [RM, (S, S), P_rb],
        10: [RM, O, P_rt, P_rb],
        11: [(S, 0), RM, P_rt],
        12: [TM, (S, 0), P_rt],
    }
    return houses

def _poly_centroid(poly):
    A = Cx = Cy = 0.0
    n = len(poly)
    for i in range(n):
        x1,y1 = poly[i]; x2,y2 = poly[(i+1)%n]
        cross = x1*y2 - x2*y1
        A += cross; Cx += (x1+x2)*cross; Cy += (y1+y2)*cross
    A *= 0.5
    if abs(A) < 1e-9:
        xs,ys = zip(*poly); return (sum(xs)/n, sum(ys)/n)
    return (Cx/(6*A), Cy/(6*A))

def _baseline_for_house(poly, S: float, house_num: int, start_anchor, end_anchor):
    """Compute baseline oriented from start_anchor toward end_anchor direction.
    Creates deterministic baseline where p1 aligns with start cusp direction 
    and p2 aligns with end cusp direction."""
    cx, cy = _poly_centroid(poly)
    pad = max(6.0, S*0.035)
    w = max(28.0, S*0.18)
    
    start_x, start_y = start_anchor
    end_x, end_y = end_anchor
    
    # For houses with same start/end cusp (H3,H6,H9,H12), create horizontal baseline
    if abs(start_x - end_x) < 1e-6 and abs(start_y - end_y) < 1e-6:
        # Single point cusp - create horizontal baseline centered at centroid
        half_w = w / 2
        x1 = max(pad, min(cx - half_w, S - pad))
        y1 = max(pad, min(cy, S - pad))
        x2 = max(pad, min(cx + half_w, S - pad))
        y2 = y1
        return (x1, y1), (x2, y2)
    
    # Calculate baseline direction vector from start to end cusp
    dx = end_x - start_x
    dy = end_y - start_y
    length = max(1e-9, (dx*dx + dy*dy)**0.5)
    dx_norm = dx / length
    dy_norm = dy / length
    
    # Create baseline parallel to cusp direction (not perpendicular)
    # This ensures p1→p2 follows start_anchor→end_anchor direction
    half_w = w / 2
    
    # Position baseline points along the cusp direction
    # p1 is positioned toward start_anchor direction
    # p2 is positioned toward end_anchor direction
    x1 = cx - dx_norm * half_w
    y1 = cy - dy_norm * half_w
    x2 = cx + dx_norm * half_w
    y2 = cy + dy_norm * half_w
    
    # Apply bounds checking
    x1 = max(pad, min(x1, S - pad))
    y1 = max(pad, min(y1, S - pad))
    x2 = max(pad, min(x2, S - pad))
    y2 = max(pad, min(y2, S - pad))
    
    return (x1, y1), (x2, y2)

def _interpolate(p1, p2, t: float):
    return (p1[0] + (p2[0]-p1[0])*t, p1[1] + (p2[1]-p1[1])*t)

def _get_cusp_corners_for_house(house_num: int, S: float):
    """Get the corner points that represent cusps for a given house.
    Returns (start_corner, end_corner) corresponding to BhavBegin and BhavEnd.
    
    Triangular houses use outer midpoints as cusps.
    Quad houses use inner corners as cusps.
    """
    TM = (S/2, 0); RM = (S, S/2); BM = (S/2, S); LM = (0, S/2)
    O  = (S/2, S/2)
    P_lt = (S/4, S/4); P_rt = (3*S/4, S/4); P_rb = (3*S/4, 3*S/4); P_lb = (S/4, 3*S/4)
    
    # Corrected cusp corners for proper triangular/quad house mapping:
    # Triangular houses (2,3,5,6,8,9,11,12) use outer midpoints
    # Quad houses (1,4,7,10) use inner corners
    cusp_mapping = {
        1:  (P_lt, P_rt),     # Quad house 1: left inner to right inner
        2:  (TM, LM),         # Triangular house 2: top middle to left middle
        3:  (LM, LM),         # Triangular house 3: left middle (single point cusp)
        4:  (P_lb, P_lt),     # Quad house 4: left bottom inner to left top inner
        5:  (LM, BM),         # Triangular house 5: left middle to bottom middle
        6:  (BM, BM),         # Triangular house 6: bottom middle (single point cusp)
        7:  (P_rb, P_lb),     # Quad house 7: right bottom inner to left bottom inner
        8:  (BM, RM),         # Triangular house 8: bottom middle to right middle
        9:  (RM, RM),         # Triangular house 9: right middle (single point cusp)
        10: (P_rt, P_rb),     # Quad house 10: right top inner to right bottom inner
        11: (RM, TM),         # Triangular house 11: right middle to top middle
        12: (TM, TM),         # Triangular house 12: top middle (single point cusp)
    }
    
    start_corner, end_corner = cusp_mapping.get(house_num, (P_lt, P_rt))
    return start_corner, end_corner

def _compute_cusp_anchors(poly, S: float, house_num: int):
    """Compute start and end cusp anchor points for a house polygon.
    Uses proper geometric cusp mapping instead of heuristics."""
    start_corner, end_corner = _get_cusp_corners_for_house(house_num, S)
    
    # Apply padding to ensure anchors are within bounds
    pad = max(4.0, S*0.02)
    start_anchor = (max(pad, min(start_corner[0], S-pad)), max(pad, min(start_corner[1], S-pad)))
    end_anchor = (max(pad, min(end_corner[0], S-pad)), max(pad, min(end_corner[1], S-pad)))
    
    return start_anchor, end_anchor

def _apply_cusp_positioning(baseline_xy, t: float, mid_t: float, cusp_anchors, 
                           start_deg: float, end_deg: float, planet_deg: float,
                           cusp_snap_deg: float, cusp_bias_deg: float):
    """Apply cusp snapping and mid-split positioning logic."""
    start_anchor, end_anchor = cusp_anchors
    
    # Calculate proximity to start and end cusps
    start_proximity = abs(_fwd_arc(start_deg, planet_deg))
    if start_proximity > 180:
        start_proximity = 360 - start_proximity
    
    end_proximity = abs(_fwd_arc(planet_deg, end_deg))
    if end_proximity > 180:
        end_proximity = 360 - end_proximity
    
    # Determine if we should snap to cusp
    if start_proximity <= cusp_snap_deg:
        return start_anchor
    elif end_proximity <= cusp_snap_deg:
        return end_anchor
    
    # Apply mid-split strictness - ensure placement in correct half
    if t < mid_t:
        # First half: start to mid
        # Adjust t to be relative to first half
        adjusted_t = t / mid_t if mid_t > 0 else 0
        mid_baseline = _interpolate(baseline_xy[0], _interpolate(baseline_xy[0], baseline_xy[1], mid_t), adjusted_t)
    else:
        # Second half: mid to end
        # Adjust t to be relative to second half
        adjusted_t = (t - mid_t) / (1 - mid_t) if mid_t < 1 else 0
        mid_point = _interpolate(baseline_xy[0], baseline_xy[1], mid_t)
        mid_baseline = _interpolate(mid_point, baseline_xy[1], adjusted_t)
    
    # Apply cusp bias if within bias range
    if start_proximity <= cusp_bias_deg:
        bias_factor = 1 - (start_proximity / cusp_bias_deg)
        bias_factor = bias_factor ** 2  # Quadratic bias
        return _interpolate(mid_baseline, start_anchor, bias_factor)
    elif end_proximity <= cusp_bias_deg:
        bias_factor = 1 - (end_proximity / cusp_bias_deg)
        bias_factor = bias_factor ** 2  # Quadratic bias
        return _interpolate(mid_baseline, end_anchor, bias_factor)
    
    return mid_baseline

def _border_anchor_for_shift(houses, h_rasi: int, forward: bool, S: float):
    poly = houses[h_rasi]
    cx, cy = _poly_centroid(poly)
    neighbor = 1 if (h_rasi == 12 and forward) else (h_rasi + 1 if forward else (h_rasi - 1 if h_rasi > 1 else 12))
    n_cx, n_cy = _poly_centroid(houses[neighbor])
    t = 0.55
    sx = cx + (n_cx - cx)*t
    sy = cy + (n_cy - cy)*t
    pad = max(4.0, S*0.02)
    sx = max(pad, min(sx, S - pad))
    sy = max(pad, min(sy, S - pad))
    return (sx, sy)

def _planet_label(code: str) -> str:
    return HN_ABBR.get(code, code)

def render_kundali_chalit(
    size_pt: float,
    lagna_sign: int,
    sidelons: Dict[str, float],     # Su,Mo,Ma,Me,Ju,Ve,Sa,Ra,Ke in degrees 0..360 (sidereal)
    begins_sid: List[float],         # 1-based BhavBegin sidereal (index 1..12)
    mids_sid: List[float],           # 1-based Bhava Madhya sidereal (index 1..12)
    pair_threshold_deg: float = 6.0,
    color: str = "#FF6600",          # User-selected color for theming
    cusp_snap_deg: float = 0.5,      # Snap to corner/cusp if within this many degrees
    cusp_bias_deg: float = 2.0       # Bias toward corner/cusp if within this many degrees
, planet_labels=None, planet_flags=None):
    """Return a VML group (XML element) to append to a python-docx cell."""
    S = float(size_pt)
    houses = _house_polys(S)
    
    # Generate color variants for theming
    def hex_to_rgb(hex_color):
        hex_color = hex_color.lstrip('#')
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    
    def rgb_to_hex(rgb):
        return '#{:02x}{:02x}{:02x}'.format(int(rgb[0]), int(rgb[1]), int(rgb[2]))
    
    def lighten_color(hex_color, factor=0.7):
        r, g, b = hex_to_rgb(hex_color)
        r = int(r + (255 - r) * factor)
        g = int(g + (255 - g) * factor)
        b = int(b + (255 - b) * factor)
        return rgb_to_hex((r, g, b))
    
    def darken_color(hex_color, factor=0.3):
        r, g, b = hex_to_rgb(hex_color)
        r = int(r * (1 - factor))
        g = int(g * (1 - factor))
        b = int(b * (1 - factor))
        return rgb_to_hex((r, g, b))
    
    # Create color variants for different elements - make borders much darker
    stroke_color = darken_color(color, 0.6)  # Much darker variant for borders
    fill_color = lighten_color(color, 0.7)   # Light variant for fill

    baselines = {}
    mids_xy   = {}
    cusp_anchors = {}
    mid_fractions = {}
    for h in range(1,13):
        # Compute cusp anchors first to get proper start/end positions
        cusp_anchors[h] = _compute_cusp_anchors(houses[h], S, h)
        start_anchor, end_anchor = cusp_anchors[h]
        
        # Create baseline oriented with cusp direction
        p1, p2 = _baseline_for_house(houses[h], S, h, start_anchor, end_anchor)
        baselines[h] = (p1, p2)
        mids_xy[h]   = _interpolate(p1, p2, 0.5)
        
        # Calculate mid fraction for this house
        start = begins_sid[h]
        end = begins_sid[1] if h == 12 else begins_sid[h+1]
        mid = mids_sid[h]
        mid_fraction = _arc_fraction(start, end, mid)
        
        # Baseline orientation is now deterministic from _baseline_for_house()
        # No need for swap logic since p1→p2 direction is aligned with start→end cusp direction
        
        mid_fractions[h] = mid_fraction

    placements = []
    order = ['Su','Mo','Ma','Me','Ju','Ve','Sa','Ra','Ke']
    for code in order:
        lon = _n360(sidelons[code])
        h_r = ((int(_n360(lon)//30) + 1 - lagna_sign) % 12) + 1
        # per-chalit house
        h_c = _house_chalit(begins_sid, lon)

        start = begins_sid[h_c]
        end   = begins_sid[1] if h_c == 12 else begins_sid[h_c+1]
        t     = _arc_fraction(start, end, lon)

        p1, p2 = baselines[h_c]
        # Apply enhanced cusp positioning with mid-split logic
        chalit_xy = _apply_cusp_positioning(
            (p1, p2), t, mid_fractions[h_c], cusp_anchors[h_c],
            start, end, lon, cusp_snap_deg, cusp_bias_deg
        )
        print(f"DEBUG: Planet {code} - rasi_house={h_r}, chalit_house={h_c}, lon={lon:.1f}, position=({chalit_xy[0]:.1f},{chalit_xy[1]:.1f})")

        disp_xy = chalit_xy
        effective_xy = chalit_xy
        shift_arrow = None

        if h_c != h_r:
            if h_c == (h_r % 12) + 1:  # forward shift
                boundary = begins_sid[(h_r % 12) + 1]
                extra = _fwd_arc(boundary, lon)
                degs  = _deg_only(extra)
                # Use enhanced cusp positioning for the arrow start position as well
                start_anchor_h_r, end_anchor_h_r = cusp_anchors[h_r]
                start_xy = _interpolate(start_anchor_h_r, end_anchor_h_r, 0.8)  # Near end of previous house
                shift_arrow = dict(start=start_xy, end=chalit_xy, label=f"{degs}°")
                disp_xy = start_xy  # Keep arrow start at border
                effective_xy = chalit_xy  # Enhanced position in destination house
                print(f"DEBUG: Forward shift {code} from house {h_r} to {h_c}, arrow: ({start_xy[0]:.1f},{start_xy[1]:.1f}) -> ({chalit_xy[0]:.1f},{chalit_xy[1]:.1f})")
            elif h_c == (h_r - 2) % 12 + 1:  # backward shift  
                boundary = begins_sid[h_r]
                extra = _fwd_arc(lon, boundary)
                degs  = _deg_only(extra)
                # Use enhanced cusp positioning for the arrow start position as well
                start_anchor_h_r, end_anchor_h_r = cusp_anchors[h_r]
                start_xy = _interpolate(start_anchor_h_r, end_anchor_h_r, 0.2)  # Near start of previous house
                shift_arrow = dict(start=start_xy, end=chalit_xy, label=f"{degs}°")
                disp_xy = start_xy  # Keep arrow start at border
                effective_xy = chalit_xy  # Enhanced position in destination house
                print(f"DEBUG: Backward shift {code} from house {h_r} to {h_c}, arrow: ({start_xy[0]:.1f},{start_xy[1]:.1f}) -> ({chalit_xy[0]:.1f},{chalit_xy[1]:.1f})")

        placements.append(dict(
            code=code, label=_planet_label(code),
            h_r=h_r, h_c=h_c, lon=lon, t=t,
            disp_xy=disp_xy, eff_xy=effective_xy,
            shift=shift_arrow
        ))

    # Close-pair arrows (≤ 6°) within each chalit house
    pair_arrows = []
    for h in range(1,13):
        Ps = [p for p in placements if p['h_c'] == h]
        if len(Ps) < 2:
            continue
        Ps.sort(key=lambda p: p['t'])
        start = begins_sid[h]; end = begins_sid[1] if h == 12 else begins_sid[h+1]
        span  = _fwd_arc(start, end) or 1e-9
        for i in range(len(Ps)-1):
            A, B = Ps[i], Ps[i+1]
            sep_deg = abs(B['t'] - A['t']) * span
            if sep_deg <= pair_threshold_deg + 1e-9:
                pair_arrows.append(dict(
                    start=A['eff_xy'], end=B['eff_xy'],
                    label=f"{_deg_only(sep_deg)}°"
                ))

    # Compose VML (frame + diagonals + baselines + planets + arrows)
    frame = f'''
      <v:rect style="position:absolute;left:0;top:0;width:{S}pt;height:{S}pt;z-index:1" strokecolor="{stroke_color}" strokeweight="4pt" fillcolor="{fill_color}"/>
      <v:line style="position:absolute;z-index:2" from="0,0" to="{S},{S}" strokecolor="{stroke_color}" strokeweight="2.5pt"/>
      <v:line style="position:absolute;z-index:2" from="{S},0" to="0,{S}" strokecolor="{stroke_color}" strokeweight="2.5pt"/>
      <v:line style="position:absolute;z-index:2" from="{S/2},0" to="{S},{S/2}" strokecolor="{stroke_color}" strokeweight="2.5pt"/>
      <v:line style="position:absolute;z-index:2" from="{S},{S/2}" to="{S/2},{S}" strokecolor="{stroke_color}" strokeweight="2.5pt"/>
      <v:line style="position:absolute;z-index:2" from="{S/2},{S}" to="0,{S/2}" strokecolor="{stroke_color}" strokeweight="2.5pt"/>
      <v:line style="position:absolute;z-index:2" from="0,{S/2}" to="{S/2},0" strokecolor="{stroke_color}" strokeweight="2.5pt"/>
    '''

    shapes = [frame]

    # House numbers + baselines + mid ticks
    def _house_labels(lagna_sign:int)->dict:
        order = [str(((lagna_sign - 1 + i) % 12) + 1) for i in range(12)]
        return {i+1: order[i] for i in range(12)}
    labels = _house_labels(lagna_sign)

    house_polys = _house_polys(S)
    for h in range(1,13):
        cx, cy = _poly_centroid(house_polys[h])
        num_w, num_h = 10, 12
        left = cx - num_w/2; top = cy - num_h/2
        shapes.append(f'''
        <v:rect style="position:absolute;left:{left}pt;top:{top}pt;width:{num_w}pt;height:{num_h}pt;z-index:80" fillcolor="#ffffff" strokecolor="none">
          <v:textbox inset="0,0,0,0"><w:txbxContent xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>{labels[h]}</w:t></w:r></w:p>
          </w:txbxContent></v:textbox>
        </v:rect>''')
        # baseline removed per requirement
        # mid tick removed as per requirement

    # Planet markers + shift arrows
    mark_w, mark_h = 16, 12
    for p in placements:
        x, y = p['disp_xy']
        left = x - mark_w/2; top = y - mark_h/2
        shapes.append(f'''
        <v:rect style="position:absolute;left:{left}pt;top:{top}pt;width:{mark_w}pt;height:{mark_h}pt;z-index:6" strokecolor="none" fillcolor="#ffffff" strokeweight="0.75pt">
          <v:textbox inset="0,0,0,0"><w:txbxContent xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>{_safe_get_label(p['code'], planet_labels)}</w:t></w:r></w:p>
          </w:txbxContent></v:textbox>
        </v:rect>''')
        # overlays
        if planet_flags:
            is_self = _flag(planet_flags, p['code'], 'self')
            is_varg = _flag(planet_flags, p['code'], 'vargottama')
            if is_self:
                cx_o = left + mark_w - 4; cy_o = top + mark_h - 4
                shapes.append(f'''<v:oval style="position:absolute;left:{cx_o-2}pt;top:{cy_o-2}pt;width:4pt;height:4pt;z-index:8" fillcolor="#000000" strokecolor="none"/>''')
            if is_varg:
                badge_w, badge_h = 6, 6
                bx = left + mark_w - badge_w + 0.5; by = top - badge_h/2
                shapes.append(f'''<v:rect style="position:absolute;left:{bx}pt;top:{by}pt;width:{badge_w}pt;height:{badge_h}pt;z-index:8" fillcolor="#ffffff" strokecolor="#666666" strokeweight="0.5pt"/>''')
        if p['shift']:
            a = p['shift']
            sx, sy = a['start']; ex, ey = a['end']
            # shorten arrow to ~1/3rd of original length
            ex = sx + (ex - sx) * SHIFT_ARROW_SCALE

            ey = sy + (ey - sy) * SHIFT_ARROW_SCALE

            shapes.append(f'''
            <v:line style="position:absolute;z-index:7" from="{sx},{sy}" to="{ex},{ey}" strokecolor="#333333" strokeweight="1pt">
              <v:stroke endarrow="classic"/>
            </v:line>
            <v:rect style="position:absolute;left:{(sx+ex)/2 - 8}pt;top:{(sy+ey)/2 + SHIFT_LABEL_OFFSET_PT}pt;width:16pt;height:10pt;z-index:8" strokecolor="none">
              <v:textbox inset="0,0,0,0"><w:txbxContent xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>{a['label']}</w:t></w:r></w:p>
              </w:txbxContent></v:textbox>
            </v:rect>''')

    # Close-pair double-headed arrows
    for ar in pair_arrows:
        (sx,sy) = ar['start']; (ex,ey) = ar['end']
        shapes.append(f'''
        <v:line style="position:absolute;z-index:9" from="{sx},{sy}" to="{ex},{ey}" strokecolor="#7a2e2e" strokeweight="1pt">
          <v:stroke endarrow="classic" startarrow="classic"/>
        </v:line>
        <v:rect style="position:absolute;left:{(sx+ex)/2 - 8}pt;top:{(sy+ey)/2 + SHIFT_LABEL_OFFSET_PT}pt;width:16pt;height:10pt;z-index:10" strokecolor="none">
          <v:textbox inset="0,0,0,0"><w:txbxContent xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>{ar['label']}</w:t></w:r></w:p>
          </w:txbxContent></v:textbox>
        </v:rect>''')

    xml = f'''
    <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r>
      <w:pict xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word"><w10:wrap type="topAndBottom"/>
        <v:group style="position:relative;margin-left:auto;margin-right:auto;margin-top:0;width:{S}pt;height:{int(S*0.80)}pt" coordorigin="0,0" coordsize="{S},{S}">
          {''.join(shapes)}
        </v:group>
      </w:pict>
    </w:r></w:p>
    '''
    return parse_xml(xml)
