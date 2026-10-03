# -*- coding: utf-8 -*-
"""
extract_stress_ranges.py
Extract delta-S11 (stress range) along pipeline KP between step pairs of an
Abaqus ODB and write one "delta S11" report per fatigue case, in the format
read by the macro in FatigueDamageCal_ManualWorkflows_Hydrotest_Design.xlsm.

Run with the Abaqus interpreter (Python 2.7 or 3.x releases). Everything can
be given on the command line; anything left out comes from USER INPUT below.

    abaqus python extract_stress_ranges.py --odb model.odb --list
    abaqus python extract_stress_ranges.py --odb model.odb --step-pair 22 23 14 16
        --phase "DESIGN OPERATION" HYDROTEST --s11 --abs-max-s11
    abaqus python extract_stress_ranges.py --odb model.odb --step-pair 22 23
        --range-from-abs-max      # abs max of heat-up step minus last frame of cool-down
    abaqus python extract_stress_ranges.py -h          # all options

Output (folder OUT_DIR):
    <case>.rpt        tab-separated: Pipeline Distance [m], then DELTA_S11 at
                      INNER/OUTER fibre, angles -90, 0, 90, 180 (stress in Pa)
    manifest.json     list of reports with phase / load case, read by
                      run_fatigue_workbook.py
    pair_comparison.txt   when several pairs share a phase and load case: each
                      pair's maximum stress range and its KP; CHECK_*.rpt/.png
                      (local U2 and COPEN along KP) if they differ by more
                      than --compare-tol percent
    S11_steps.rpt     (--s11) S11 of every requested step
    U2_steps.rpt      (--u2) local U2 along KP of every requested step
    ABSMAX_S11.rpt    (--abs-max-s11) absolute maximum S11 over all frames
"""
from __future__ import print_function, division
import json
import math
import os
import re
import sys

# ============================ USER INPUT =====================================
ODB = "BPTiber_With_BM_FL6_TestBase.odb"   # <-- default ODB; a case can name its own

# One entry per fatigue case = one stress-range block in the workbook.
#   phase      EARLY UNDRAINED | EARLY DRAINED | MIDDLE DRAINED | LATE DRAINED
#              | HYDROTEST | DESIGN OPERATION
#   load_case  FCD | HCD | PCD   (ignored for HYDROTEST / DESIGN OPERATION)
#   odb        optional: ODB file for this case (default: ODB above)
#   pairs      list of (step A, step B); delta = S11(A) - S11(B) at frame FRAME.
#              With several pairs the largest |delta| at each point is kept.
#              A step is given by a unique part of its name ("Step 014") or by
#              its 1-based position in the ODB (14).
CASES = [
    {"label": "Hydrotest", "phase": "HYDROTEST", "load_case": "FCD",
     "pairs": [("Step 014", "Step 015"),      # Hydrotest - Depressure
               ("Step 016", "Step 017")]},    # Re-hydrotest - Re-depressure
    {"label": "DesignOp", "phase": "DESIGN OPERATION", "load_case": "FCD",
     "pairs": [("Step 022", "Step 023")]},    # Design - FCD after Design
    # {"label": "EarlyUndr_FCD", "phase": "EARLY UNDRAINED", "load_case": "FCD",
    #  "odb": "other_model.odb",
    #  "pairs": [("Step 019", "Step 020")]},  # 1st Operation - FCD
]

INSTANCE = None             # None = the only instance in the ODB (flat input: PART-1-1)
ELSET = "PIPE_ELEMENTS"     # pipeline element set (instance or assembly level)
FRAME = -1                  # frame within each step (-1 = last)

# KP [m] = cumulative element length from the first element, at element mid-length.
ORDER_BY = "label"          # "label", or "x" / "y" / "z" to sort by coordinate
KP_START = 0.0              # KP [m] at the start of the first element
LENGTH_TO_M = 1.0           # model length unit -> m
STRESS_TO_PA = 1.0          # model stress unit -> Pa (the workbook divides by 1e6)

# Section points. By default the fibre and angle are read from the section
# point descriptions stored in the ODB, e.g.
#   "Angle = 90.0000, (1-fraction = 0.000000, 2-fraction = 0.677710)"
#   "Angle = 45.0000, Radius = 0.6777"
# The radial position is sqrt(f1^2 + f2^2), or the Radius value. The smallest
# radial position of a section is the inner fibre, the largest (1.0) the outer
# fibre; points in between are ignored.
# Descriptions containing INNER / OUTER are also understood. If neither works,
# run with --list and fill this map by hand:
#   {section point number: ("INNER" or "OUTER", angle)}
SECTION_POINT_MAP = {}

# True: stress range = S11 of largest magnitude over ALL frames of step A
# (heat-up) minus S11 at frame FRAME of step B (cool-down).
# False: frame FRAME of both steps. Command line: --range-from-abs-max
RANGE_FROM_ABS_MAX = False

# Buckle check: the ODB stores nodal displacements in the global system. Local
# U2 is obtained by projecting the global displacement on the horizontal
# direction normal to the as-laid pipe axis at each node (positive to the left
# when looking towards increasing KP). VERTICAL_AXIS is the global vertical.
VERTICAL_AXIS = "z"         # "x", "y" or "z"; command line: --vertical-axis
U2_AS_STORED = False        # True: use the stored U2 component without conversion

OUT_DIR = "fatigue_reports"
# =============================================================================

TARGET_ANGLES = (-90.0, 0.0, 90.0, 180.0)
FIBERS = ("INNER", "OUTER")
ANGLE_RE = re.compile(r"angle\s*=?\s*([-+]?\d+(?:\.\d*)?(?:[eE][-+]?\d+)?)", re.I)


NUM = r"([-+]?\d+(?:\.\d*)?(?:[eE][-+]?\d+)?)"
FRACTION_RE = re.compile(r"1-fraction\s*=\s*" + NUM + r"\s*,\s*2-fraction\s*=\s*" + NUM, re.I)
RADIUS_RE = re.compile(r"radius\s*=\s*" + NUM, re.I)
EL_CAT = {}       # element label -> section category name
CAT_RADII = {}    # section category name -> (smallest, largest) radial fraction


def radial_fraction(sp):
    """Radial position (radius / outer radius) of a pipe section point from its
    description. Abaqus writes it in one of two forms:
      'Angle = 90.0000, (1-fraction = 0.000000, 2-fraction = 0.677710)'
          the two fractions are the local 1 and 2 coordinates over the outer
          radius, so the radial fraction is sqrt(f1^2 + f2^2)
      'Angle = 45.0000, Radius = 0.6777'
    """
    text = sp.description or ""
    m = FRACTION_RE.search(text)
    if m:
        return math.hypot(float(m.group(1)), float(m.group(2)))
    m = RADIUS_RE.search(text)
    if m:
        return float(m.group(1))
    return None


def prepare_categories(elems):
    """Record each element's section category and, per category, the smallest
    and largest radial fraction: smallest = inner fibre, largest = outer."""
    EL_CAT.clear()
    CAT_RADII.clear()
    cats = {}
    for e in elems:
        cat = e.sectionCategory
        EL_CAT[e.label] = cat.name
        cats[cat.name] = cat
    for name, cat in cats.items():
        fr = [f for f in (radial_fraction(sp) for sp in cat.sectionPoints) if f is not None]
        if fr:
            CAT_RADII[name] = (min(fr), max(fr))
    return cats


def classify_section_point(sp, cat_name=None):
    """Return (fibre, angle) for a target section point, else None."""
    if sp is None:
        return None
    if SECTION_POINT_MAP:
        hit = SECTION_POINT_MAP.get(sp.number)
        if hit is None:
            return None
        fiber, angle = hit[0].upper(), float(hit[1])
    else:
        text = sp.description or ""
        up = text.upper()
        frac = radial_fraction(sp)
        if "INNER" in up or "INSIDE" in up:
            fiber = "INNER"
        elif "OUTER" in up or "OUTSIDE" in up:
            fiber = "OUTER"
        elif frac is not None and cat_name in CAT_RADII:
            rmin, rmax = CAT_RADII[cat_name]
            if abs(frac - rmax) < 1.0e-3:
                fiber = "OUTER"
            elif abs(frac - rmin) < 1.0e-3:
                fiber = "INNER"
            else:
                return None
        else:
            return None
        m = ANGLE_RE.search(text)
        if not m:
            return None
        angle = float(m.group(1))
    if abs(angle + 180.0) < 0.01:
        angle = 180.0
    for t in TARGET_ANGLES:
        if abs(angle - t) < 0.01:
            return fiber, t
    return None


def get_instance(odb):
    insts = odb.rootAssembly.instances
    if INSTANCE:
        return insts[INSTANCE]
    names = [n for n in insts.keys() if n.upper() != "ASSEMBLY"]
    if len(names) != 1:
        raise KeyError("Set INSTANCE; the ODB has instances: %s" % ", ".join(names))
    return insts[names[0]]


def get_region(odb):
    """Return (instance, region, element list) for the pipeline set."""
    asm = odb.rootAssembly
    inst = get_instance(odb)
    for name in (ELSET, ELSET.upper()):
        if name in inst.elementSets.keys():
            region = inst.elementSets[name]
            return inst, region, list(region.elements)
        if name in asm.elementSets.keys():
            region = asm.elementSets[name]
            for i, iname in enumerate(region.instanceNames):
                if iname == inst.name:
                    return inst, region, list(region.elements[i])
    raise KeyError("Element set %s not found. Instance sets: %s"
                   % (ELSET, ", ".join(sorted(inst.elementSets.keys())[:40])))


def build_kp(inst, elems):
    """Ordered element labels, {element: KP [m] at mid-length}, {node: KP [m]}."""
    wanted = set()
    for e in elems:
        wanted.update(e.connectivity)
    coords = {}
    for n in inst.nodes:
        if n.label in wanted:
            coords[n.label] = [float(c) for c in n.coordinates]
    rows = []
    for e in elems:
        a = coords[e.connectivity[0]]
        b = coords[e.connectivity[-1]]
        length = math.sqrt(sum((b[i] - a[i]) ** 2 for i in range(3)))
        mid = [(a[i] + b[i]) / 2.0 for i in range(3)]
        rows.append((e.label, length, mid, e.connectivity[0], e.connectivity[-1]))
    if ORDER_BY == "label":
        rows.sort(key=lambda r: r[0])
    else:
        axis = {"x": 0, "y": 1, "z": 2}[ORDER_BY.lower()]
        rows.sort(key=lambda r: r[2][axis])
    kp, node_kp, order, s = {}, {}, [], 0.0
    for label, length, _mid, n_first, n_last in rows:
        kp[label] = KP_START + (s + length / 2.0) * LENGTH_TO_M
        node_kp.setdefault(n_first, KP_START + s * LENGTH_TO_M)
        order.append(label)
        s += length
        node_kp.setdefault(n_last, KP_START + s * LENGTH_TO_M)
    return order, kp, node_kp, coords


def resolve_step(odb, spec):
    names = list(odb.steps.keys())
    if isinstance(spec, int):
        if not 1 <= spec <= len(names):
            raise KeyError("Step number %d out of range 1..%d" % (spec, len(names)))
        return names[spec - 1]
    if spec in names:
        return spec
    hits = [n for n in names if spec.lower() in n.lower()]
    if len(hits) == 1:
        return hits[0]
    raise KeyError("Step '%s' matches %d steps. Steps in ODB:\n  %s"
                   % (spec, len(hits), "\n  ".join(names)))


def read_s11(frame, region, inst_name):
    """{(element, integration point, fibre, angle): S11} at the target points."""
    fo = frame.fieldOutputs["S"].getSubset(region=region)
    i11 = list(fo.componentLabels).index("S11")
    out, cache = {}, {}

    def classify(sp, label):
        cat = EL_CAT.get(label)
        key = (cat, None if sp is None else sp.number)
        if key not in cache:
            cache[key] = classify_section_point(sp, cat)
        return cache[key]

    try:
        blocks = fo.bulkDataBlocks
    except AttributeError:
        blocks = None
    if blocks is not None:
        for b in blocks:
            if b.instance is not None and b.instance.name != inst_name:
                continue
            sp = b.sectionPoint
            labels, ips, data = b.elementLabels, b.integrationPoints, b.data
            for k in range(len(labels)):
                label = int(labels[k])
                cls = classify(sp, label)
                if cls is None:
                    continue
                row = data[k]
                val = float(row[i11]) if hasattr(row, "__len__") else float(row)
                out[(label, int(ips[k]), cls[0], cls[1])] = val
        return out
    for v in fo.values:
        if v.instance is not None and v.instance.name != inst_name:
            continue
        cls = classify(v.sectionPoint, v.elementLabel)
        if cls is not None:
            out[(v.elementLabel, v.integrationPoint, cls[0], cls[1])] = float(v.data[i11])
    return out


def list_odb(odb, elems, region):
    print("Steps in ODB:")
    for i, name in enumerate(odb.steps.keys(), 1):
        print("  %2d  %s  (%d frames)" % (i, name, len(odb.steps[name].frames)))
    cats = prepare_categories(elems)
    print("Section points of the pipeline elements (* = used):")
    for cname in sorted(cats):
        print("  category: %s" % cname)
        for sp in cats[cname].sectionPoints:
            cls = classify_section_point(sp, cname)
            print("   %s %3d  %-45s %s" % ("*" if cls else " ", sp.number,
                                           sp.description, cls if cls else ""))
    step = odb.steps[list(odb.steps.keys())[-1]]
    fo = step.frames[-1].fieldOutputs
    if "S" not in fo.keys():
        print("WARNING: no S field output in the last step.")
        return
    data = read_s11(step.frames[-1], region, get_instance(odb).name)
    found = sorted(set((k[2], k[3]) for k in data))
    print("S11 output found at target points in the last step: %s" % found)


def last_frame(odb, name, region, inst_name, cache):
    """S11 at frame FRAME of a step, read once."""
    if name not in cache:
        print("  reading S11: %s" % name)
        cache[name] = read_s11(odb.steps[name].frames[FRAME], region, inst_name)
    return cache[name]


ABS_MAX_CACHE = {}


def abs_max_raw(odb, name, region, inst_name):
    """{(el, ip, fibre, angle): signed S11 of largest magnitude over all frames}."""
    key = (id(odb), name)
    if key not in ABS_MAX_CACHE:
        frames = odb.steps[name].frames
        print("  absolute maximum S11 over %d frames: %s" % (len(frames), name))
        out = {}
        for frame in frames:
            if "S" not in frame.fieldOutputs.keys():
                continue
            for k, v in read_s11(frame, region, inst_name).items():
                if k not in out or abs(v) > abs(out[k]):
                    out[k] = v
        ABS_MAX_CACHE[key] = out
    return ABS_MAX_CACHE[key]


def case_deltas(odb, case, region, inst_name, cache):
    """{(element, fibre, angle): delta S11 [Pa]}, largest |delta| over pairs."""
    out, names = {}, []
    for spec_a, spec_b in case["pairs"]:
        step_a, step_b = resolve_step(odb, spec_a), resolve_step(odb, spec_b)
        names.append("%s - %s" % (step_a, step_b))
        if RANGE_FROM_ABS_MAX:
            names[-1] = "abs max over all frames of %s - frame %d of %s" \
                        % (step_a, FRAME, step_b)
            data_a = abs_max_raw(odb, step_a, region, inst_name)
        else:
            data_a = last_frame(odb, step_a, region, inst_name, cache)
        data_b = last_frame(odb, step_b, region, inst_name, cache)
        for key, sa in data_a.items():
            if key not in data_b:
                continue
            d = (sa - data_b[key]) * STRESS_TO_PA
            k2 = (key[0], key[2], key[3])
            if k2 not in out or abs(d) > abs(out[k2]):
                out[k2] = d
    return out, names


def safe_name(text):
    bad = [w for w in ("INNER", "OUTER", "ANGLE", "DISTANCE", "[", "]", "\t")
           if w in text.upper()]
    if bad:
        raise ValueError("Case label '%s' must not contain %s" % (text, bad))
    return text


def write_report(path, case, order, kp, deltas, step_names, odb_path):
    columns = [(f, a) for f in FIBERS for a in TARGET_ANGLES]
    present = set((k[1], k[2]) for k in deltas)
    missing = [c for c in columns if c not in present]
    if missing:
        raise RuntimeError(
            "No S11 output at %s for case %s. Request S at the inner and outer "
            "section points at angles -90, 0, 90, 180 (see --list)."
            % (missing, case["label"]))
    label = safe_name(case["label"])
    fh = open(path, "w")
    fh.write("Delta S11 report for the fatigue workbook\n")
    fh.write("ODB: %s\n" % odb_path)
    fh.write("Case: %s / %s / %s\n" % (label, case["phase"], case["load_case"]))
    for n in step_names:
        fh.write("Step pair: %s\n" % n)
    fh.write("Units: KP in m, stress in Pa. Delta = first step minus second step.\n")
    if RANGE_FROM_ABS_MAX:
        fh.write("First step: S11 of largest magnitude over all frames, with its sign.\n")
    head = ["Pipeline Distance [m]"]
    for f, a in columns:
        head.append("DELTA_S11 %s FIBER Angle = %.1f [%s]" % (f, a, label))
    fh.write("\t".join(head) + "\n")
    rows, peak, peak_kp = 0, 0.0, None
    for el in order:
        vals = [deltas.get((el, f, a)) for f, a in columns]
        if all(v is None for v in vals):
            continue
        fh.write("%.4f\t" % kp[el]
                 + "\t".join("" if v is None else "%.6E" % v for v in vals) + "\n")
        rows += 1
        m = max(abs(v) for v in vals if v is not None)
        if m > peak:
            peak, peak_kp = m, kp[el]
    fh.close()
    return rows, peak, peak_kp


def write_table(path, title_lines, blocks, order, kp, tag):
    """Tab-separated table: KP, then for each (name, data) block the eight
    fibre/angle columns. data = {(element, fibre, angle): value in Pa}."""
    columns = [(f, a) for f in FIBERS for a in TARGET_ANGLES]
    fh = open(path, "w")
    for line in title_lines:
        fh.write(line + "\n")
    head = ["Pipeline Distance [m]"]
    for name, _data in blocks:
        for f, a in columns:
            head.append("%s %s FIBER Angle = %.1f [%s]" % (tag, f, a, name))
    fh.write("\t".join(head) + "\n")
    for el in order:
        vals = []
        for _name, data in blocks:
            vals += [data.get((el, f, a)) for f, a in columns]
        fh.write("%.4f\t" % kp[el]
                 + "\t".join("" if v is None else "%.6E" % v for v in vals) + "\n")
    fh.close()


def per_element(data):
    """{(el, ip, fibre, angle): v} -> {(el, fibre, angle): v of largest |v|} in Pa."""
    out = {}
    for key, v in data.items():
        k2 = (key[0], key[2], key[3])
        v = v * STRESS_TO_PA
        if k2 not in out or abs(v) > abs(out[k2]):
            out[k2] = v
    return out


def read_nodal(frame, key, component, inst_name, wanted):
    """{node label: value} of a nodal field output at the wanted nodes.
    component is e.g. 'U2', or None for a scalar field such as COPEN."""
    fo = frame.fieldOutputs[key]
    idx = list(fo.componentLabels).index(component) if component else 0
    out = {}

    def pick(row):
        return float(row[idx]) if hasattr(row, "__len__") else float(row)

    try:
        blocks = fo.bulkDataBlocks
    except AttributeError:
        blocks = None
    if blocks is not None:
        for b in blocks:
            if b.instance is not None and b.instance.name != inst_name:
                continue
            labels, data = b.nodeLabels, b.data
            for k in range(len(labels)):
                n = int(labels[k])
                if n in wanted:
                    out[n] = pick(data[k])
        return out
    for v in fo.values:
        if v.instance is not None and v.instance.name != inst_name:
            continue
        if v.nodeLabel in wanted:
            out[v.nodeLabel] = pick(v.data)
    return out


def read_nodal_vector(frame, key, inst_name, wanted):
    """{node label: (v1, v2, v3)} of a nodal vector field output (global system)."""
    fo = frame.fieldOutputs[key]
    out = {}

    def pick(row):
        vals = [float(x) for x in row][:3]
        return tuple(vals + [0.0] * (3 - len(vals)))

    try:
        blocks = fo.bulkDataBlocks
    except AttributeError:
        blocks = None
    if blocks is not None:
        for b in blocks:
            if b.instance is not None and b.instance.name != inst_name:
                continue
            labels, data = b.nodeLabels, b.data
            for k in range(len(labels)):
                n = int(labels[k])
                if n in wanted:
                    out[n] = pick(data[k])
        return out
    for v in fo.values:
        if v.instance is not None and v.instance.name != inst_name:
            continue
        if v.nodeLabel in wanted:
            out[v.nodeLabel] = pick(v.data)
    return out


def lateral_directions(nodes, node_xyz):
    """{node: unit vector} horizontal and normal to the as-laid pipe axis:
    vertical x tangent, the tangent taken from the neighbouring nodes
    (nodes must be ordered along KP)."""
    iv = {"x": 0, "y": 1, "z": 2}[VERTICAL_AXIS.lower()]
    ez = [0.0, 0.0, 0.0]
    ez[iv] = 1.0
    out, last = {}, None
    for i, n in enumerate(nodes):
        a = node_xyz[nodes[max(i - 1, 0)]]
        b = node_xyz[nodes[min(i + 1, len(nodes) - 1)]]
        t = [b[k] - a[k] for k in range(3)]
        t[iv] = 0.0                                  # horizontal part of the tangent
        lat = [ez[1] * t[2] - ez[2] * t[1],
               ez[2] * t[0] - ez[0] * t[2],
               ez[0] * t[1] - ez[1] * t[0]]
        norm = math.sqrt(sum(c * c for c in lat))
        if norm > 0.0:
            last = [c / norm for c in lat]
        out[n] = last
    first = next((v for v in out.values() if v is not None), [0.0, 0.0, 0.0])
    for n in nodes:
        if out[n] is None:
            out[n] = first
    return out


def read_local_u2(frame, inst_name, nodes, node_xyz, lateral):
    """{node: local U2}: global U projected on the lateral direction."""
    wanted = set(nodes)
    if U2_AS_STORED:
        return read_nodal(frame, "U", "U2", inst_name, wanted)
    vec = read_nodal_vector(frame, "U", inst_name, wanted)
    return dict((n, sum(u[k] * lateral[n][k] for k in range(3))) for n, u in vec.items())


def read_copen(frame, inst_name, wanted, key_filter):
    """Contact opening at the pipe nodes: smallest COPEN over the contact
    output variables whose name contains key_filter (all, if not given)."""
    keys = [k for k in frame.fieldOutputs.keys() if k.upper().startswith("COPEN")
            and (not key_filter or key_filter.lower() in k.lower())]
    out = {}
    for k in keys:
        for n, v in read_nodal(frame, k, None, inst_name, wanted).items():
            if n not in out or v < out[n]:
                out[n] = v
    return out, keys


def write_node_table(path, title_lines, columns, nodes, node_kp):
    """Tab-separated table: KP, then one column per (name, {node: value})."""
    fh = open(path, "w")
    for line in title_lines:
        fh.write(line + "\n")
    fh.write("\t".join(["KP [m]"] + [name for name, _d in columns]) + "\n")
    for n in nodes:
        vals = [d.get(n) for _name, d in columns]
        fh.write("%.4f\t" % node_kp[n]
                 + "\t".join("" if v is None else "%.6E" % v for v in vals) + "\n")
    fh.close()


def buckle_check(odb, group_name, steps, marks, inst_name, node_kp, node_xyz,
                 copen_key, tag):
    """Write local U2 and COPEN along KP for the given steps (frame FRAME), and
    a PNG plot when matplotlib is available. marks = KP of each pair's maximum."""
    nodes = sorted(node_kp, key=lambda n: node_kp[n])
    wanted = set(nodes)
    lateral = lateral_directions(nodes, node_xyz)
    u2, cop, copen_keys = [], [], []
    for name in steps:
        frame = odb.steps[name].frames[FRAME]
        if "U" in frame.fieldOutputs.keys():
            u2.append((name, read_local_u2(frame, inst_name, nodes, node_xyz, lateral)))
        data, copen_keys = read_copen(frame, inst_name, wanted, copen_key)
        if data:
            cop.append((name, data))
    base = os.path.join(OUT_DIR, "CHECK_%s%s" % (group_name, tag))
    files = []
    if u2:
        write_node_table(base + "_U2.rpt",
                         ["Local U2 along KP at frame %d of each step" % FRAME,
                          "Stored U2 component of the ODB, not converted." if U2_AS_STORED else
                          "Global U projected on the horizontal direction normal to the "
                          "as-laid pipe axis; positive to the left looking towards "
                          "increasing KP (vertical axis: global %s)." % VERTICAL_AXIS.upper(),
                          "Units: KP in m, U2 in model length units."],
                         u2, nodes, node_kp)
        files.append(base + "_U2.rpt")
    else:
        print("    no U field output in these steps: U2 check skipped")
    if cop:
        write_node_table(base + "_COPEN.rpt",
                         ["Contact opening COPEN at the pipe nodes along KP at frame %d of each "
                          "step (smallest value over: %s)" % (FRAME, "; ".join(copen_keys)),
                          "Units: KP in m, COPEN in model length units."],
                         cop, nodes, node_kp)
        files.append(base + "_COPEN.rpt")
    else:
        print("    no COPEN field output at the pipe nodes in these steps: "
              "upheaval check skipped")
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except Exception:
        print("    matplotlib is not available in this Abaqus Python: plot the "
              ".rpt files instead")
        return files
    panels = [p for p in ((u2, "Local U2"), (cop, "COPEN")) if p[0]]
    if not panels:
        return files
    fig, axes = plt.subplots(len(panels), 1, figsize=(12, 4.2 * len(panels)), sharex=True)
    if len(panels) == 1:
        axes = [axes]
    for ax, (columns, ylabel) in zip(axes, panels):
        for name, data in columns:
            xs = [node_kp[n] for n in nodes if n in data]
            ys = [data[n] for n in nodes if n in data]
            ax.plot(xs, ys, linewidth=1.2, label=name)
        for x in marks:
            ax.axvline(x, color="0.5", linestyle="--", linewidth=0.8)
        ax.set_ylabel(ylabel)
        ax.grid(True, color="0.85")
        ax.legend(fontsize=7, loc="best")
    axes[0].set_title("%s: buckle shape (U2) and contact opening per step; dashed = KP of "
                      "each pair's maximum stress range" % group_name, fontsize=10)
    axes[-1].set_xlabel("KP [m]")
    fig.tight_layout()
    fig.savefig(base + ".png", dpi=150)
    plt.close(fig)
    files.append(base + ".png")
    return files


def compare_pairs(odb, cases, path, region, inst, kp, node_kp, node_xyz, cache, tol,
                  copen_key, tag):
    """For every phase / load case with more than one step pair: table of each
    pair's maximum stress range and where it occurs. If the maxima differ by
    more than tol percent, write the U2 / COPEN checks for that group."""
    groups, keys = {}, []
    for case in cases:
        if case.get("odb", ODB) != path:
            continue
        key = (case["phase"], case["load_case"])
        if key not in groups:
            groups[key] = []
            keys.append(key)
        for pair in case["pairs"]:
            if pair not in groups[key]:
                groups[key].append(pair)
    lines = ["Maximum stress range of each step pair, per phase and load case",
             "ODB: %s" % os.path.abspath(path),
             "Difference = (largest maximum of the group - this pair's maximum) / largest "
             "maximum. Check tolerance: %g %%" % tol, "",
             "\t".join(["Phase", "Load case", "Step pair", "Max |delta S11| [MPa]",
                        "KP [m]", "Fibre", "Angle", "Difference [%]", "Check"])]
    any_group = False
    names_all = list(odb.steps.keys())
    for key in keys:
        pairs = groups[key]
        if len(pairs) < 2:
            continue
        any_group = True
        rows, steps = [], []
        for pair in pairs:
            deltas, names = case_deltas(odb, {"pairs": [pair]}, region, inst.name, cache)
            if not deltas:
                continue
            k = max(sorted(deltas), key=lambda q: round(abs(deltas[q]), 3))
            rows.append((names[0], abs(deltas[k]) / 1.0e6, kp[k[0]], k[1], k[2]))
            for spec in pair:
                name = resolve_step(odb, spec)
                if name not in steps:
                    steps.append(name)
        if len(rows) < 2:
            continue
        top = max(r[1] for r in rows)
        diffs = [(top - r[1]) / top * 100.0 if top > 0 else 0.0 for r in rows]
        flagged = max(diffs) > tol
        group_name = re.sub(r"[^A-Za-z0-9_.-]+", "_",
                            "_".join(x for x in key if x)) or "pairs"
        print("  compare %s %s: %d pairs, largest difference %.1f %%%s"
              % (key[0] or "-", key[1], len(rows), max(diffs),
                 " > %g %% -> buckle check" % tol if flagged else ""))
        for r, d in zip(rows, diffs):
            print("    %-60s %9.2f MPa at KP %10.2f  %s %g  (%.1f %%)"
                  % (r[0][:60], r[1], r[2], r[3], r[4], d))
            lines.append("\t".join([key[0], key[1], r[0], "%.3f" % r[1], "%.2f" % r[2],
                                    r[3], "%g" % r[4], "%.1f" % d,
                                    "CHECK" if flagged else "ok"]))
        if flagged:
            steps = [n for n in names_all if n in steps]
            files = buckle_check(odb, group_name, steps, [r[2] for r in rows],
                                 inst.name, node_kp, node_xyz, copen_key, tag)
            for f in files:
                print("    -> %s" % os.path.abspath(f))
            lines.append("\t".join([key[0], key[1], "check files: "
                                    + ", ".join(os.path.basename(f) for f in files)]))
        lines.append("")
    if any_group:
        out = os.path.join(OUT_DIR, "pair_comparison%s.txt" % tag)
        fh = open(out, "w")
        fh.write("\n".join(lines) + "\n")
        fh.close()
        print("  pair comparison -> %s" % os.path.abspath(out))


def parse_args(argv):
    import argparse
    ap = argparse.ArgumentParser(
        description="Delta S11 along pipeline KP between step pairs of an Abaqus ODB.")
    ap.add_argument("--odb", help="ODB file (default: ODB in the script)")
    ap.add_argument("--step-pair", nargs="+", metavar="STEP",
                    help="steps in pairs: A1 B1 [A2 B2 ...]; delta = S11(A) - S11(B). "
                         "A step is its number in the ODB or a unique part of its name")
    ap.add_argument("--phase", nargs="+", metavar="PHASE",
                    help="workbook phase of each pair, e.g. \"DESIGN OPERATION\" HYDROTEST")
    ap.add_argument("--load-case", nargs="+", metavar="LC",
                    help="load case of each pair: FCD, HCD, PCD, HYDROTEST1, "
                         "HYDROTEST2 or DESIGN (default FCD)")
    ap.add_argument("--label", nargs="+", metavar="NAME",
                    help="report name of each case (default StepA-StepB, or "
                         "PHASE_LOADCASE_env with --envelope)")
    ap.add_argument("--envelope", action="store_true",
                    help="combine the pairs that share the same --phase and "
                         "--load-case into one case each, keeping the largest "
                         "|delta| at each point (all pairs, if those are not given "
                         "per pair)")
    ap.add_argument("--s11", action="store_true",
                    help="also write S11 of every requested step (frame --frame)")
    ap.add_argument("--u2", action="store_true",
                    help="also write U2_steps.rpt: local U2 along KP of every requested "
                         "step (import it with button 6 of the workbook for the U2 plot)")
    ap.add_argument("--abs-max-s11", nargs="?", const="1", choices=["1", "2", "both"],
                    help="also write the absolute maximum S11 over all frames of the "
                         "first step of each pair (1, heat-up; default), the second "
                         "(2) or both")
    ap.add_argument("--range-from-abs-max", action="store_true",
                    help="stress range = absolute maximum S11 over all frames of the "
                         "first (heat-up) step minus S11 at frame --frame of the "
                         "second (cool-down) step")
    ap.add_argument("--compare-tol", type=float, default=10.0, metavar="PCT",
                    help="when several step pairs share a phase and load case, their "
                         "maximum stress ranges are tabulated; if they differ by more "
                         "than PCT percent (default 10), local U2 and COPEN along KP "
                         "are written and plotted for those steps")
    ap.add_argument("--vertical-axis", choices=["x", "y", "z"],
                    help="global vertical axis, used to convert the global U of the "
                         "ODB to local (lateral) U2 (default %s)" % VERTICAL_AXIS)
    ap.add_argument("--u2-as-stored", action="store_true",
                    help="use the stored U2 component without converting it")
    ap.add_argument("--copen-key", metavar="TEXT",
                    help="only use COPEN output variables whose name contains TEXT "
                         "(default: all, taking the smallest opening)")
    ap.add_argument("--elset", help="pipeline element set (default %s)" % ELSET)
    ap.add_argument("--instance", help="instance name (default: the only instance)")
    ap.add_argument("--frame", type=int, help="frame used in each step (default -1 = last)")
    ap.add_argument("--kp-start", type=float, help="KP [m] at the start of the first element")
    ap.add_argument("--out-dir", help="output folder (default %s)" % OUT_DIR)
    ap.add_argument("--list", action="store_true",
                    help="only list the steps and section points of the ODB")
    return ap.parse_args(argv)


def step_spec(text):
    return int(text) if str(text).lstrip("+").isdigit() else text


def cases_from_args(args):
    """Build the case list from --step-pair, or use CASES from the script."""
    if not args.step_pair:
        return [dict(c) for c in CASES]
    specs = [step_spec(t) for t in args.step_pair]
    if len(specs) % 2:
        raise ValueError("--step-pair needs an even number of steps (A1 B1 A2 B2 ...)")
    pairs = [(specs[i], specs[i + 1]) for i in range(0, len(specs), 2)]
    n = len(pairs)

    def per_pair(values, default, name):
        if not values:
            return [default] * n
        if len(values) == 1:
            return list(values) * n
        if len(values) != n:
            raise ValueError("--%s: give one value or one per step pair (%d pairs)"
                             % (name, n))
        return list(values)

    phases = [v.upper() for v in per_pair(args.phase, "", "phase")]
    load_cases = [v.upper() for v in per_pair(args.load_case, "FCD", "load-case")]

    # Without --envelope every pair is its own case. With it, pairs that share
    # the same phase and load case are combined (largest |delta| at each point).
    groups, index = [], {}
    for i, pair in enumerate(pairs):
        key = (phases[i], load_cases[i]) if args.envelope else i
        if key not in index:
            index[key] = len(groups)
            groups.append({"phase": phases[i], "load_case": load_cases[i], "pairs": []})
        groups[index[key]]["pairs"].append(pair)

    if args.label and len(args.label) != len(groups):
        raise ValueError("--label: give one name per case (%d cases)" % len(groups))
    cases = []
    for i, grp in enumerate(groups):
        if args.label:
            label = args.label[i]
        elif args.envelope and (args.phase or args.load_case):
            label = "_".join(x for x in (grp["phase"], grp["load_case"]) if x) + "_env"
        else:
            label = "_".join("Step%s-Step%s" % p for p in grp["pairs"])
        grp["label"] = re.sub(r"[^A-Za-z0-9_.-]+", "_", label) if not args.label else label
        cases.append(grp)
    return cases


def main():
    global ODB, ELSET, INSTANCE, FRAME, KP_START, OUT_DIR, RANGE_FROM_ABS_MAX
    global VERTICAL_AXIS, U2_AS_STORED
    from odbAccess import openOdb
    args = parse_args(sys.argv[1:])
    if args.odb:
        ODB = args.odb
    if args.elset:
        ELSET = args.elset
    if args.instance:
        INSTANCE = args.instance
    if args.frame is not None:
        FRAME = args.frame
    if args.kp_start is not None:
        KP_START = args.kp_start
    if args.out_dir:
        OUT_DIR = args.out_dir
    if args.range_from_abs_max:
        RANGE_FROM_ABS_MAX = True
    if args.vertical_axis:
        VERTICAL_AXIS = args.vertical_axis
    if args.u2_as_stored:
        U2_AS_STORED = True
    cases = cases_from_args(args)
    labels = [c["label"] for c in cases]
    if len(set(labels)) != len(labels):
        raise ValueError("Case labels must be unique: %s" % labels)

    odb_paths = []
    for case in cases:
        path = case.get("odb", ODB)
        if path not in odb_paths:
            odb_paths.append(path)
    if not args.list and not os.path.isdir(OUT_DIR):
        os.makedirs(OUT_DIR)
    done, ref = {}, None
    for n_odb, path in enumerate(odb_paths):
        print("ODB %s" % path)
        tag = "" if len(odb_paths) == 1 else "_odb%d" % (n_odb + 1)
        odb = openOdb(path=path, readOnly=True)
        try:
            inst, region, elems = get_region(odb)
            prepare_categories(elems)
            if args.list:
                list_odb(odb, elems, region)
                continue
            order, kp, node_kp, node_xyz = build_kp(inst, elems)
            grid = [kp[el] for el in order]
            print("  %d elements, KP %.2f to %.2f m" % (len(grid), grid[0], grid[-1]))
            # The workbook holds one KP grid; a different grid would make its
            # macro clear the stress ranges already imported.
            if ref is None:
                ref = (path, grid)
            elif len(grid) != len(ref[1]) or \
                    max(abs(x - y) for x, y in zip(grid, ref[1])) > 0.001:
                raise RuntimeError("KP grid of %s differs from %s; all ODBs fed to "
                                   "one workbook must share the same mesh."
                                   % (path, ref[0]))
            cache, first_steps, second_steps = {}, [], []
            for i, case in enumerate(cases):
                if case.get("odb", ODB) != path:
                    continue
                print("  case %s" % case["label"])
                deltas, names = case_deltas(odb, case, region, inst.name, cache)
                rpt = os.path.abspath(os.path.join(OUT_DIR, case["label"] + ".rpt"))
                rows, peak, peak_kp = write_report(rpt, case, order, kp, deltas,
                                                   names, os.path.abspath(path))
                print("    %d KP rows, max |delta S11| = %.2f MPa at KP %.2f m -> %s"
                      % (rows, peak / 1.0e6, peak_kp, rpt))
                done[i] = {"label": case["label"], "phase": case["phase"],
                           "load_case": case["load_case"], "report": rpt,
                           "odb": os.path.abspath(path), "step_pairs": names}
                for spec_a, spec_b in case["pairs"]:
                    for spec, bucket in ((spec_a, first_steps), (spec_b, second_steps)):
                        name = resolve_step(odb, spec)
                        if name not in bucket:
                            bucket.append(name)
            compare_pairs(odb, cases, path, region, inst, kp, node_kp, node_xyz, cache,
                          args.compare_tol, args.copen_key, tag)
            names_all = list(odb.steps.keys())
            used = [n for n in names_all if n in first_steps or n in second_steps]
            if args.s11:
                for n in used:
                    last_frame(odb, n, region, inst.name, cache)
                out = os.path.join(OUT_DIR, "S11_steps%s.rpt" % tag)
                write_table(out, ["S11 at frame %d of each requested step" % FRAME,
                                  "ODB: %s" % os.path.abspath(path),
                                  "Units: KP in m, stress in Pa."],
                            [(n, per_element(cache[n])) for n in used], order, kp, "S11")
                print("  S11 of %d steps -> %s" % (len(used), os.path.abspath(out)))
            if args.u2:
                nodes = sorted(node_kp, key=lambda n: node_kp[n])
                lateral = lateral_directions(nodes, node_xyz)
                cols = []
                for n in used:
                    frame = odb.steps[n].frames[FRAME]
                    if "U" in frame.fieldOutputs.keys():
                        cols.append((n, read_local_u2(frame, inst.name, nodes, node_xyz, lateral)))
                if cols:
                    out = os.path.join(OUT_DIR, "U2_steps%s.rpt" % tag)
                    write_node_table(out,
                                     ["Local U2 along KP at frame %d of each requested step" % FRAME,
                                      "ODB: %s" % os.path.abspath(path),
                                      "Units: KP in m, U2 in model length units."],
                                     cols, nodes, node_kp)
                    print("  U2 of %d steps -> %s" % (len(cols), os.path.abspath(out)))
                else:
                    print("  no U field output in the requested steps: U2_steps.rpt not written")
            if args.abs_max_s11:
                want = {"1": first_steps, "2": second_steps,
                        "both": first_steps + second_steps}[args.abs_max_s11]
                want = [n for n in names_all if n in want]
                blocks = []
                for n in want:
                    data = per_element(abs_max_raw(odb, n, region, inst.name))
                    blocks.append((n, data))
                    if data:
                        key = max(data, key=lambda k: abs(data[k]))
                        print("    max |S11| = %.2f MPa at KP %.2f m (%s fibre, angle %g)"
                              % (abs(data[key]) / 1.0e6, kp[key[0]], key[1], key[2]))
                out = os.path.join(OUT_DIR, "ABSMAX_S11%s.rpt" % tag)
                write_table(out, ["Absolute maximum S11 over all frames of each step",
                                  "ODB: %s" % os.path.abspath(path),
                                  "Units: KP in m, stress in Pa. Each value is the S11 "
                                  "of largest magnitude, with its sign."],
                            blocks, order, kp, "ABSMAX_S11")
                print("  -> %s" % os.path.abspath(out))
        finally:
            odb.close()
    if args.list:
        return
    fh = open(os.path.join(OUT_DIR, "manifest.json"), "w")
    json.dump([done[i] for i in sorted(done)], fh, indent=2)
    fh.close()
    print("Wrote %s" % os.path.join(OUT_DIR, "manifest.json"))
    if any(not d["phase"] and not d["load_case"].replace(" ", "").startswith(("HYDROTEST", "DESIGN"))
           for d in done.values()):
        print("Note: cases without --phase cannot be placed in the workbook by "
              "run_fatigue_workbook.py (except load case HYDROTEST1, HYDROTEST2 or DESIGN).")


if __name__ == "__main__":
    main()
