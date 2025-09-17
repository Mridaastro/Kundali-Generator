
# chalit_kundali_vml.py
# Exact on-diagram placement for Chalit chart in DOCX (VML).
# - Planet markers placed along within-house baseline by BhavBegin→BhavEnd fraction
# - Double-headed arrows for ≤6° pairs in same house (degrees-only labels)
# - Forward/backward shift arrows across house border from rashi corner (degrees-only labels)

from math import fmod
from typing import Dict, List, Tuple
from docx.oxml import parse_xml
from kundali_markers_lib import _calculate_planet_status, UP_ARROW, DOWN_ARROW, COMBUST, VARG_SQ

HN_ABBR = {'Su': 'सू', 'Mo': 'चं', 'Ma': 'मं', 'Me': 'बु', 'Ju': 'गु', 'Ve': 'शु', 'Sa': 'श', 'Ra': 'रा', 'Ke': 'के'}

# ---- Visual constants for shift arrows ----
ARROW_LEN_PX = 9
ARROW_START_GAP = 6
ARROW_LABEL_OFFSET = 12


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

def _baseline_for_house(poly, S: float):
    cx, cy = _poly_centroid(poly)
    pad = max(6.0, S*0.035)
    w = max(28.0, S*0.18)
    x1 = max(pad, min(cx - w/2, S - pad - w))
    x2 = x1 + w
    y  = max(pad, min(cy + S*0.05, S - pad))
    return (x1, y), (x2, y)

def _interpolate(p1, p2, t: float):
    return (p1[0] + (p2[0]-p1[0])*t, p1[1] + (p2[1]-p1[1])*t)

def _border_anchor_for_shift(houses, h_rasi: int, forward: bool, S: float):
    poly = houses[h_rasi]
    cx, cy = _poly_centroid(poly)
    neighbor = 1 if (h_rasi == 12 and forward) else (h_rasi + 1 if forward else (h_rasi - 1 if h_rasi > 1 else 12))
    n_cx, n_cy = _poly_centroid(houses[neighbor])
    t = 0.44
    sx = cx + (n_cx - cx)*t
    sy = cy + (n_cy - cy)*t
    pad = max(6.0, S*0.035)
    sx = max(pad, min(sx, S - pad))
    sy = max(pad, min(sy, S - pad))
    return (sx, sy)

def _planet_label_with_markers(code: str, lon: float, sidelons: dict) -> str:
    base = HN_ABBR.get(code, code)
    try:
        flags = _calculate_planet_status(code, lon, sidelons, chart_type='D1')
    except Exception:
        flags = {'exalt': False, 'debil': False, 'comb': False, 'varg': False}
    disp = base
    if flags.get('exalt'): disp += UP_ARROW
    if flags.get('debil'): disp += DOWN_ARROW
    if flags.get('comb'):  disp += COMBUST
    if flags.get('varg'):  disp += VARG_SQ
    return disp


def render_kundali_chalit(
    size_pt: float,
    lagna_sign: int,
    sidelons: Dict[str, float],     # Su,Mo,Ma,Me,Ju,Ve,Sa,Ra,Ke in degrees 0..360 (sidereal)
    begins_sid: List[float],         # 1-based BhavBegin sidereal (index 1..12)
    mids_sid: List[float],           # 1-based Bhava Madhya sidereal (index 1..12)
    pair_threshold_deg: float = 6.0
):
    """Return a VML group (XML element) to append to a python-docx cell."""
    S = float(size_pt)
    houses = _house_polys(S)

    baselines = {}
    mids_xy   = {}
    for h in range(1,13):
        p1, p2 = _baseline_for_house(houses[h], S)
        baselines[h] = (p1, p2)
        mids_xy[h]   = _interpolate(p1, p2, 0.5)

    placements = []
    order = ['Su','Mo','Ma','Me','Ju','Ve','Sa','Ra','Ke']
    for code in order:
        lon = _n360(sidelons[code])
        _flags = _calculate_planet_status(code, lon, sidelons, chart_type='D1')
        h_r = ((int(_n360(lon)//30) + 1 - lagna_sign) % 12) + 1
        # per-chalit house
        h_c = _house_chalit(begins_sid, lon)

        start = begins_sid[h_c]
        end   = begins_sid[1] if h_c == 12 else begins_sid[h_c+1]
        t     = _arc_fraction(start, end, lon)

        p1, p2 = baselines[h_c]
        chalit_xy = _interpolate(p1, p2, t)

        disp_xy = chalit_xy
        effective_xy = chalit_xy
        shift_arrow = None

        if h_c != h_r:
            if h_c == (h_r % 12) + 1:  # forward shift
                boundary = begins_sid[(h_r % 12) + 1]
                extra = _fwd_arc(boundary, lon)
                degs  = _deg_only(extra)
                start_xy = _border_anchor_for_shift(houses, h_r, forward=True, S=S)
                shift_arrow = dict(start=start_xy, end=chalit_xy, label=f"{degs}°")
                disp_xy = start_xy
                effective_xy = chalit_xy
            elif h_c == (h_r - 2) % 12 + 1:  # backward shift
                boundary = begins_sid[h_r]
                extra = _fwd_arc(lon, boundary)
                degs  = _deg_only(extra)
                start_xy = _border_anchor_for_shift(houses, h_r, forward=False, S=S)
                shift_arrow = dict(start=start_xy, end=chalit_xy, label=f"{degs}°")
                disp_xy = start_xy
                effective_xy = chalit_xy

        placements.append(dict(
            code=code, label=_planet_label_with_markers(code, lon, sidelons),
            h_r=h_r, h_c=h_c, lon=lon, t=t,
            disp_xy=disp_xy, eff_xy=effective_xy,
            shift=shift_arrow, self_flag=bool(_flags.get('own') or _flags.get('self') or _flags.get('own_sign'))
        ))


    # Re-layout labels: if >2 in a display house, arrange in rows of 2
    display_groups = {h: [] for h in range(1,13)}
    for p in placements:
        disp_house = p['h_r'] if p['shift'] else p['h_c']
        display_groups[disp_house].append(p)
    for h in range(1,13):
        group = display_groups[h]
        if len(group) <= 2:
            continue
        # Sort by 't' for stable ordering
        group.sort(key=lambda p: p['t'])
        minx, miny, maxx, maxy = _bbox_of_poly(houses[h])
        pad = max(4.0, S*0.03)
        cols = 2
        rows = (len(group) + cols - 1) // cols
        gw = (maxx - minx) - 2*pad
        gh = (maxy - miny) - 2*pad
        # Cell center positions
        xs = [minx + pad + (i + 0.5) * gw / cols for i in range(cols)]
        ys = [miny + pad + (r + 0.5) * gh / rows for r in range(rows)]
        for idx, p in enumerate(group):
            r = idx // cols
            c = idx % cols
            p['disp_xy'] = (xs[c], ys[r])
    
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
      <v:rect style="position:absolute;left:0;top:0;width:{S}pt;height:{S}pt;z-index:1" strokecolor="#CC6600" strokeweight="3pt" fillcolor="#ffdcc8"/>
      <v:line style="position:absolute;z-index:2" from="0,0" to="{S},{S}" strokecolor="#CC6600" strokeweight="1.25pt"/>
      <v:line style="position:absolute;z-index:2" from="{S},0" to="0,{S}" strokecolor="#CC6600" strokeweight="1.25pt"/>
      <v:line style="position:absolute;z-index:2" from="{S/2},0" to="{S},{S/2}" strokecolor="#CC6600" strokeweight="1.25pt"/>
      <v:line style="position:absolute;z-index:2" from="{S},{S/2}" to="{S/2},{S}" strokecolor="#CC6600" strokeweight="1.25pt"/>
      <v:line style="position:absolute;z-index:2" from="{S/2},{S}" to="0,{S/2}" strokecolor="#CC6600" strokeweight="1.25pt"/>
      <v:line style="position:absolute;z-index:2" from="0,{S/2}" to="{S/2},0" strokecolor="#CC6600" strokeweight="1.25pt"/>
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
        # center baseline + tick removed
    # Planet markers + shift arrows
    mark_w, mark_h = 22, 16
    for p in placements:
        x, y = p['disp_xy']
        left = x - mark_w/2; top = y - mark_h/2
        # Clamp inside the appropriate house bbox (rāśi for shifted, else Chalit)
        house_idx = p['h_r'] if p['shift'] else p['h_c']
        minx, miny, maxx, maxy = _bbox_of_poly(house_polys[house_idx])
        _pad = max(2.0, S*0.01)
        left = max(minx + _pad, min(left, maxx - _pad - mark_w))
        top  = max(miny + _pad, min(top,  maxy - _pad - mark_h))
        shapes.append(f'''
        <v:roundrect arcsize="0.3" style="position:absolute;left:{left}pt;top:{top}pt;width:{mark_w}pt;height:{mark_h}pt;z-index:6" stroked="f" strokecolor="none" fillcolor="#ffffff">
          <v:textbox inset="0,0,0,0"><w:txbxContent xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>{p['label']}</w:t></w:r></w:p>
          </w:txbxContent></v:textbox>
        </v:roundrect>''')
        # Encircle self-ruling planets
        if p.get('self_flag'):
            circ_pad = 2.0
            shapes.append(f'''
            <v:oval style="position:absolute;left:{left - circ_pad}pt;top:{top - circ_pad}pt;width:{mark_w + 2*circ_pad}pt;height:{mark_h + 2*circ_pad}pt;z-index:8" strokecolor="#000000" strokeweight="0.75pt" fillcolor="none"/>''')
        if p['shift']:
            circ_pad = 2.0
            shapes.append(f'''
            <v:oval style="position:absolute;left:{left - circ_pad}pt;top:{top - circ_pad}pt;width:{mark_w + 2*circ_pad}pt;height:{mark_h + 2*circ_pad}pt;z-index:8" strokecolor="#000000" strokeweight="0.75pt" fillcolor="none"/>''')

            a = p['shift']
            sx, sy = a['start']; ex, ey = a['end']
            dx = ex - sx; dy = ey - sy
            norm = (dx*dx + dy*dy) ** 0.5 or 1.0
            nx = dx / norm; ny = dy / norm
            # start gap + constant very-short arrow
            sx2 = sx + nx*ARROW_START_GAP
            sy2 = sy + ny*ARROW_START_GAP
            ex2 = sx2 + nx*ARROW_LEN_PX
            ey2 = sy2 + ny*ARROW_LEN_PX
            # degree label above arrow (perpendicular offset)
            mx = (sx2 + ex2)/2.0; my = (sy2 + ey2)/2.0
            lx = mx - ny*ARROW_LABEL_OFFSET
            ly = my + nx*ARROW_LABEL_OFFSET
            shapes.append(f'''
            <v:line style="position:absolute;z-index:9" from="{sx2},{sy2}" to="{ex2},{ey2}" strokecolor="#333333" strokeweight="1pt">
              <v:stroke endarrow="classic"/>
            </v:line>
            <v:rect style="position:absolute;left:{lx - 9}pt;top:{ly - 10}pt;width:18pt;height:12pt;z-index:8" strokecolor="none">
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
        <v:rect style="position:absolute;left:{(sx+ex)/2 - 8}pt;top:{(sy+ey)/2 - 10}pt;width:16pt;height:10pt;z-index:10" strokecolor="none">
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

def _bbox_of_poly(poly):
    xs = [x for x,_ in poly]; ys = [y for _,y in poly]
    return min(xs), min(ys), max(xs), max(ys)
