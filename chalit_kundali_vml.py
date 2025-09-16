# chalit_kundali_vml.py
# Exact on-diagram placement for Chalit chart in DOCX (VML).
# - Planet markers placed along within-house baseline by BhavBegin→BhavEnd fraction
# - Double-headed arrows for ≤6° pairs in same house (degrees-only labels)
# - Forward/backward shift arrows across house border from rashi corner (degrees-only labels)
# - Self-ruling planets shown as a thin circle INSIDE the textbox

from math import fmod
from typing import Dict, List
from docx.oxml import parse_xml

# Hindi planet abbreviations (fallback to code if not present)
HN_ABBR = {
    'Su': 'सू', 'Mo': 'चं', 'Ma': 'मं', 'Me': 'बु',
    'Ju': 'गु', 'Ve': 'शु', 'Sa': 'श', 'Ra': 'रा', 'Ke': 'के'
}

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

def _house_chalit(begins: List[float], lon_sid: float) -> int:
    """Return chalit house (1..12) from BhavBegin list (1-based)."""
    L = _n360(lon_sid)
    for n in range(1, 13):
        a = begins[n]
        b = begins[1] if n == 12 else begins[n+1]
        if _fwd_arc(a, L) < _fwd_arc(a, b):
            return n
    return 12

def _arc_fraction(start: float, end: float, pt: float) -> float:
    """0..1 fraction of pt along [start -> end) on a forward arc."""
    total = _fwd_arc(start, end) or 1e-9
    part  = _fwd_arc(start, pt)
    t = part / total
    return 0.0 if t < 0 else (1.0 if t > 1 else t)

def _house_polys(S: float):
    # Outer midpoints
    TM = (S/2, 0); RM = (S, S/2); BM = (S/2, S); LM = (0, S/2)
    O  = (S/2, S/2)
    # Inner corners
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

def _bbox(cx, cy, w, h):
    return (cx - w/2.0, cy - h/2.0, cx + w/2.0, cy + h/2.0)

def _point_in_poly(pt, poly):
    """Ray casting point in polygon (inclusive)."""
    x, y = pt
    inside = False
    n = len(poly)
    for i in range(n):
        x1, y1 = poly[i]
        x2, y2 = poly[(i + 1) % n]
        # Check if point is between y1 and y2
        if ((y1 > y) != (y2 > y)):
            xinters = (x2 - x1) * (y - y1) / (y2 - y1 + 1e-12) + x1
            if x <= xinters:
                inside = not inside
        # On edge tolerance
        # horizontal/vertical tolerance can be added if needed
    return inside

def _inset_toward_centroid(point, house_poly, inset_pt: float):
    """Move 'point' by 'inset_pt' toward polygon centroid (staying inside)."""
    cx, cy = _poly_centroid(house_poly)
    sx, sy = point
    dx, dy = (cx - sx), (cy - sy)
    d = (dx*dx + dy*dy) ** 0.5 or 1.0
    ux, uy = dx / d, dy / d
    return (sx + ux * inset_pt, sy + uy * inset_pt)

def _nudge_away_from_point(x, y, bx, by, radius=18.0):
    dx, dy = x - bx, y - by
    d2 = dx*dx + dy*dy
    if d2 == 0:
        # push right by radius
        return x + radius, y
    if d2 < radius*radius:
        d = (d2) ** 0.5
        ux, uy = dx / d, dy / d
        return bx + ux * radius, by + uy * radius
    return x, y

def _get_cusp_corners_for_house(house_num: int, S: float):
    """Get (start_cusp, end_cusp) for a house.
    Triangular houses use outer midpoints; quads use inner corners.
    """
    TM = (S/2, 0); RM = (S, S/2); BM = (S/2, S); LM = (0, S/2)
    P_lt = (S/4, S/4); P_rt = (3*S/4, S/4); P_rb = (3*S/4, 3*S/4); P_lb = (S/4, 3*S/4)

    cusp_mapping = {
        1:  (P_lt, P_rt),
        2:  (TM, LM),
        3:  (LM, LM),
        4:  (P_lb, P_lt),
        5:  (LM, BM),
        6:  (BM, BM),
        7:  (P_rb, P_lb),
        8:  (BM, RM),
        9:  (RM, RM),
        10: (P_rt, P_rb),
        11: (RM, TM),
        12: (TM, TM),
    }
    return cusp_mapping.get(house_num, (P_lt, P_rt))

def _compute_cusp_anchors(poly, S: float, house_num: int):
    start_corner, end_corner = _get_cusp_corners_for_house(house_num, S)
    pad = max(4.0, S*0.02)
    start_anchor = (max(pad, min(start_corner[0], S-pad)), max(pad, min(start_corner[1], S-pad)))
    end_anchor   = (max(pad, min(end_corner[0],   S-pad)), max(pad, min(end_corner[1],   S-pad)))
    return start_anchor, end_anchor

def _baseline_for_house(poly, S: float, house_num: int, start_anchor, end_anchor):
    """Baseline oriented from start cusp toward end cusp."""
    cx, cy = _poly_centroid(poly)
    pad = max(6.0, S*0.035)
    w = max(28.0, S*0.18)

    start_x, start_y = start_anchor
    end_x, end_y     = end_anchor

    # Houses with identical start/end cusp: horizontal baseline through centroid
    if abs(start_x - end_x) < 1e-6 and abs(start_y - end_y) < 1e-6:
        half_w = w / 2
        x1 = max(pad, min(cx - half_w, S - pad)); y1 = max(pad, min(cy, S - pad))
        x2 = max(pad, min(cx + half_w, S - pad)); y2 = y1
        return (x1, y1), (x2, y2)

    dx = end_x - start_x; dy = end_y - start_y
    length = max(1e-9, (dx*dx + dy*dy)**0.5)
    dxn, dyn = dx / length, dy / length

    half_w = w / 2
    x1 = cx - dxn * half_w; y1 = cy - dyn * half_w
    x2 = cx + dxn * half_w; y2 = cy + dyn * half_w

    x1 = max(pad, min(x1, S - pad)); y1 = max(pad, min(y1, S - pad))
    x2 = max(pad, min(x2, S - pad)); y2 = max(pad, min(y2, S - pad))
    return (x1, y1), (x2, y2)

def _interpolate(p1, p2, t: float):
    return (p1[0] + (p2[0]-p1[0])*t, p1[1] + (p2[1]-p1[1])*t)

def _border_anchor_for_shift(houses, h_rasi: int, forward: bool, S: float):
    """A point near the border between the rashi house and its next/previous house."""
    poly_r = houses[h_rasi]
    cx, cy = _poly_centroid(poly_r)
    nbr = 1 if (h_rasi == 12 and forward) else (h_rasi + 1 if forward else (h_rasi - 1 if h_rasi > 1 else 12))
    nx, ny = _poly_centroid(houses[nbr])
    # move from rashi centroid toward neighbor centroid; t slightly past halfway puts us near the border
    t = 0.52
    x = cx + (nx - cx) * t
    y = cy + (ny - cy) * t
    pad = max(4.0, S*0.02)
    x = max(pad, min(x, S - pad))
    y = max(pad, min(y, S - pad))
    return (x, y)

def _apply_cusp_positioning(baseline_xy, t: float, mid_t: float, cusp_anchors,
                            start_deg: float, end_deg: float, planet_deg: float,
                            cusp_snap_deg: float, cusp_bias_deg: float):
    """Apply cusp snapping + mid-split positioning logic for within-house placement."""
    start_anchor, end_anchor = cusp_anchors

    # proximity to cusps in degrees
    start_prox = abs(_fwd_arc(start_deg, planet_deg)); 
    if start_prox > 180: start_prox = 360 - start_prox
    end_prox   = abs(_fwd_arc(planet_deg, end_deg)); 
    if end_prox > 180: end_prox = 360 - end_prox

    if start_prox <= cusp_snap_deg: return start_anchor
    if end_prox   <= cusp_snap_deg: return end_anchor

    if t < mid_t:
        adj = t / mid_t if mid_t > 0 else 0
        pos = _interpolate(baseline_xy[0], _interpolate(baseline_xy[0], baseline_xy[1], mid_t), adj)
    else:
        adj = (t - mid_t) / (1 - mid_t) if mid_t < 1 else 0
        mid = _interpolate(baseline_xy[0], baseline_xy[1], mid_t)
        pos = _interpolate(mid, baseline_xy[1], adj)

    if start_prox <= cusp_bias_deg:
        b = 1 - (start_prox / cusp_bias_deg); b *= b
        return _interpolate(pos, start_anchor, b)
    if end_prox <= cusp_bias_deg:
        b = 1 - (end_prox / cusp_bias_deg); b *= b
        return _interpolate(pos, end_anchor, b)
    return pos

def render_kundali_chalit(
    size_pt: float,
    lagna_sign: int,
    sidelons: Dict[str, float],     # Su,Mo,Ma,Me,Ju,Ve,Sa,Ra,Ke (sidereal, degrees)
    begins_sid: List[float],         # 1-based BhavBegin sidereal (index 1..12)
    mids_sid: List[float],           # 1-based Bhava Madhya sidereal (index 1..12)
    pair_threshold_deg: float = 6.0,
    color: str = "#8B4513",          # theme base
    cusp_snap_deg: float = 0.5,      # Snap to cusp if within N degrees
    cusp_bias_deg: float = 2.0,      # Bias toward cusp within this many degrees
    planet_labels=None, planet_flags=None,
    shift_arrow_scale: float = 0.3333, shift_label_offset_pt: float = 8.0
):
    """Return a VML group (XML element) to append to a python-docx cell."""
    S = float(size_pt)
    houses = _house_polys(S)

    # simple theme helpers
    def hex_to_rgb(h): h=h.lstrip('#'); return (int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))
    def rgb_to_hex(r,g,b): return '#%02x%02x%02x'%(max(0,min(255,r)),max(0,min(255,g)),max(0,min(255,b)))
    def lighten(h, f=0.7):
        r,g,b=hex_to_rgb(h); r=int(r+(255-r)*f); g=int(g+(255-g)*f); b=int(b+(255-b)*f); return rgb_to_hex(r,g,b)
    def darken(h, f=0.6):
        r,g,b=hex_to_rgb(h); r=int(r*(1-f)); g=int(g*(1-f)); b=int(b*(1-f)); return rgb_to_hex(r,g,b)

    stroke_color = darken(color, 0.55)
    fill_color   = lighten(color, 0.75)

    # Precompute baselines and cusp anchors
    baselines, cusp_anchors, mid_fractions = {}, {}, {}
    for h in range(1,13):
        cusp_anchors[h] = _compute_cusp_anchors(houses[h], S, h)
        p1, p2 = _baseline_for_house(houses[h], S, h, *cusp_anchors[h])
        baselines[h] = (p1, p2)

        start = begins_sid[h]; end = begins_sid[1] if h==12 else begins_sid[h+1]
        mid   = mids_sid[h]
        mid_fractions[h] = _arc_fraction(start, end, mid)

    # Build planet placements
    placements = []
    for code in ['Su','Mo','Ma','Me','Ju','Ve','Sa','Ra','Ke']:
        lon = _n360(sidelons[code])
        # rashi house (1..12) relative to lagna
        h_r = ((int(_n360(lon)//30) + 1 - lagna_sign) % 12) + 1
        # chalit house (1..12)
        h_c = _house_chalit(begins_sid, lon)

        start = begins_sid[h_c]
        end   = begins_sid[1] if h_c == 12 else begins_sid[h_c+1]
        t     = _arc_fraction(start, end, lon)

        p1, p2 = baselines[h_c]
        chalit_xy = _apply_cusp_positioning((p1, p2), t, mid_fractions[h_c], cusp_anchors[h_c],
                                            start, end, lon, cusp_snap_deg, cusp_bias_deg)

        disp_xy = chalit_xy
        effective_xy = chalit_xy
        shift_arrow = None

        if h_c != h_r:
            # very small inset so the glyph sits right near the correct border
            inset_shift_pt = max(2.0, S * 0.007)

            if h_c == (h_r % 12) + 1:  # forward shift → border with next house
                boundary = begins_sid[(h_r % 12) + 1]
                extra = _fwd_arc(boundary, lon); degs = _deg_only(extra)

                border_xy = _border_anchor_for_shift(houses, h_r, True, S)
                disp_xy   = _inset_toward_centroid(border_xy, houses[h_r], inset_shift_pt)

                disp_xy = _ensure_point_inside(disp_xy, houses[h_r], 6.0, 8)
                chalit_xy = _ensure_point_inside(chalit_xy, houses[h_c], 6.0, 8)
                shift_arrow = dict(start=disp_xy, end=chalit_xy, label=f"{degs}°")
                effective_xy = chalit_xy

            elif h_c == (h_r - 2) % 12 + 1:  # backward shift → border with previous house
                boundary = begins_sid[h_r]
                extra = _fwd_arc(lon, boundary); degs = _deg_only(extra)

                border_xy = _border_anchor_for_shift(houses, h_r, False, S)
                disp_xy   = _inset_toward_centroid(border_xy, houses[h_r], inset_shift_pt)

                disp_xy = _ensure_point_inside(disp_xy, houses[h_r], 6.0, 8)
                chalit_xy = _ensure_point_inside(chalit_xy, houses[h_c], 6.0, 8)
                shift_arrow = dict(start=disp_xy, end=chalit_xy, label=f"{degs}°")
                effective_xy = chalit_xy

        placements.append(dict(
            code=code, label=HN_ABBR.get(code, code),
            h_r=h_r, h_c=h_c, lon=lon, t=t,
            disp_xy=disp_xy, eff_xy=effective_xy,
            shift=shift_arrow
        ))

    # Close-pair arrows (≤ threshold) within each chalit house
    pair_arrows = []
    for h in range(1,13):
        Ps = [p for p in placements if p['h_c'] == h]
        if len(Ps) < 2: continue
        Ps.sort(key=lambda p: p['t'])
        start = begins_sid[h]; end = begins_sid[1] if h == 12 else begins_sid[h+1]
        span  = _fwd_arc(start, end) or 1e-9
        for i in range(len(Ps)-1):
            A, B = Ps[i], Ps[i+1]
            sep_deg = abs(B['t'] - A['t']) * span
            if sep_deg <= pair_threshold_deg + 1e-9:
                pair_arrows.append(dict(start=A['eff_xy'], end=B['eff_xy'], label=f"{_deg_only(sep_deg)}°"))

    # VML frame (no midlines)
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

    # House numbers
    def _house_labels(lagna_sign:int)->dict:
        order = [str(((lagna_sign - 1 + i) % 12) + 1) for i in range(12)]
        return {i+1: order[i] for i in range(12)}
    labels = _house_labels(lagna_sign)
    for h in range(1,13):
        cx, cy = _poly_centroid(houses[h])
        num_w, num_h = 10, 12
        left = cx - num_w/2; top = cy - num_h/2
        shapes.append(f'''
        <v:rect style="position:absolute;left:{left}pt;top:{top}pt;width:{num_w}pt;height:{num_h}pt;z-index:80" fillcolor="#ffffff" strokecolor="none">
          <v:textbox inset="0,0,0,0"><w:txbxContent xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>{labels[h]}</w:t></w:r></w:p>
          </w:txbxContent></v:textbox>
        </v:rect>''')

    # Planet markers + shift arrows
    # Increased textbox size to prevent glyph clipping
    mark_w, mark_h = 18, 14
    # --- Lane layout to avoid planet overlap per house ---
    # Build per-house index lists
    per_house = {h: [] for h in range(1,13)}
    for i, p in enumerate(placements):
        per_house[p['h_c']].append(i)
    # Normal vectors and sorting along baseline
    for h, idxs in per_house.items():
        if not idxs:
            continue
        # Sort by param t (along the baseline)
        idxs.sort(key=lambda i: placements[i]['t'])
        # Baseline unit normal for this house
        p1, p2 = baselines[h]
        dx, dy = (p2[0] - p1[0]), (p2[1] - p1[1])
        L = (dx*dx + dy*dy) ** 0.5 or 1.0
        # Perpendicular (normal) unit vector
        nx, ny = -dy / L, dx / L
        # Space between lanes
        lane_step = max(mark_h * 1.15, 12.0)
        # Assign lanes in the pattern 0, +1, -1, +2, -2, ...
        def lane_for(k):
            # k is 0-based order in the sorted list
            if k == 0:
                return 0
            n = (k + 1) // 2
            return n if k % 2 == 1 else -n
        placed_rects = []
        for k, i_p in enumerate(idxs):
            p = placements[i_p]
            x0, y0 = p['disp_xy']
            lane = lane_for(k)
            x = x0 + nx * lane * lane_step
            y = y0 + ny * lane * lane_step
            # Keep inside the house polygon; try mirrored lane, else inset retries
            if not _point_in_poly((x, y), houses[h]):
                lane_alt = -lane
                x_alt = x0 + nx * lane_alt * lane_step
                y_alt = y0 + ny * lane_alt * lane_step
                if _point_in_poly((x_alt, y_alt), houses[h]):
                    x, y = x_alt, y_alt
                else:
                    fx, fy = x, y
                    for _ in range(6):
                        fx, fy = _inset_toward_centroid((fx, fy), houses[h], 6.0)
                        if _point_in_poly((fx, fy), houses[h]):
                            break
                    x, y = fx, fy
            # Slide along baseline if overlapping previously placed labels in this house
            tx, ty = (dx / L, dy / L) if L else (1.0, 0.0)
            step_t = max(mark_w * 0.9, 10.0)
            tries = 0
            rect = _bbox(x, y, mark_w, mark_h)
            def any_rects_overlap(r):
                return any(not (r[2] <= pr[0] or pr[2] <= r[0] or r[3] <= pr[1] or pr[3] <= r[1]) for pr in placed_rects)
            while any_rects_overlap(rect) and tries < 12:
                # alternate +t and -t moves
                direction = 1 if tries % 2 == 0 else -1
                x += direction * tx * step_t
                y += direction * ty * step_t
                # keep inside polygon; if not, backtrack and reduce step
                if not _point_in_poly((x,y), houses[h]):
                    x -= direction * tx * step_t
                    y -= direction * ty * step_t
                    step_t *= 0.75
                rect = _bbox(x, y, mark_w, mark_h)
                tries += 1
            placed_rects.append(rect)
            # Exclude collision with the house-number box near centroid
            cx, cy = _poly_centroid(houses[h])
            x, y   = _nudge_away_from_point(x, y, cx, cy, radius=18.0)
            x, y = _ensure_rect_inside(x, y, mark_w, mark_h, houses[h], step_inset=4.0, max_iter=12)
            p['disp_xy_adj'] = (x, y)

    for p in placements:
        x, y = p.get('disp_xy_adj', p['disp_xy'])
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

            # Self-ruling: draw a thin ring INSIDE the textbox (no fill)
            if is_self:
                pad = 2.0                                   # clearance to textbox edge
                d = max(8.0, min(mark_w, mark_h) - 2*pad)   # ring diameter
                cx = left + mark_w/2; cy = top + mark_h/2   # center of label
                shapes.append(
                    f'''<v:oval style="position:absolute;left:{cx - d/2}pt;top:{cy - d/2}pt;width:{d}pt;height:{d}pt;z-index:8"
                         fillcolor="none" strokecolor="#333333" strokeweight="0.75pt"/>'''
                )

            # Vargottama badge (unchanged)
            if is_varg:
                badge_w, badge_h = 6, 6
                bx = left + mark_w - badge_w + 0.5; by = top - badge_h/2
                shapes.append(f'''<v:rect style="position:absolute;left:{bx}pt;top:{by}pt;width:{badge_w}pt;height:{badge_h}pt;z-index:8" fillcolor="#ffffff" strokecolor="#666666" strokeweight="0.5pt"/>''')

        # Shift arrow (start offset away from label; shorten to ~1/3)
        if p['shift']:
            a = p['shift']
            sx, sy = p.get('disp_xy_adj', p['disp_xy']); ex, ey = a['end']
            dx, dy = (ex - sx), (ey - sy)
            d = (dx*dx + dy*dy) ** 0.5 or 1.0
            ux, uy = dx/d, dy/d
            gap = max(6.0, mark_h * 0.65)                  # offset so arrow doesn't overlap label
            sx2, sy2 = sx + ux*gap, sy + uy*gap
            ex2 = sx2 + (ex - sx2) * shift_arrow_scale
            ey2 = sy2 + (ey - sy2) * shift_arrow_scale
            shapes.append(f'''
            <v:line style="position:absolute;z-index:7" from="{sx2},{sy2}" to="{ex2},{ey2}" strokecolor="#333333" strokeweight="1pt">
              <v:stroke endarrow="classic"/>
            </v:line>
            <v:rect style="position:absolute;left:{(sx2+ex2)/2 - 8}pt;top:{(sy2+ey2)/2 + shift_label_offset_pt}pt;width:16pt;height:10pt;z-index:8" strokecolor="none">
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
        <v:rect style="position:absolute;left:{(sx+ex)/2 - 8}pt;top:{(sy+ey)/2 + shift_label_offset_pt}pt;width:16pt;height:10pt;z-index:10" strokecolor="none">
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
