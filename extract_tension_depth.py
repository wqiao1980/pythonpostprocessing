# -*- coding: utf-8 -*-
"""
extract_tension_depth.py
Extract the effective tension (ESF1) and the Z coordinate (water depth) of the
pipeline along KP from an Abaqus ODB, for the steps or step ranges you give,
and write reports that PipelineResultPlots.xlsm imports and plots.

Run with Abaqus Python, in the folder that also holds extract_stress_ranges.py
(this script uses its KP routines):

    abaqus python extract_tension_depth.py --odb model.odb --list
    abaqus python extract_tension_depth.py --odb model.odb --step-range 5 8
    abaqus python extract_tension_depth.py --odb model.odb --step-range 5 8 12 14 --step 22
    abaqus python extract_tension_depth.py -h          # all options

Output, in the output folder (default pipeline_reports):
    EFFTENSION.rpt   effective tension along KP (element mid-length), one column per step
    ZCOORD.rpt       Z coordinate along KP (nodes), one column per step
"""
from __future__ import print_function
import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import extract_stress_ranges as es

OUT_DIR = "pipeline_reports"


def parse_args(argv):
    ap = argparse.ArgumentParser(
        description="Effective tension and Z coordinate along pipeline KP from an Abaqus ODB.")
    ap.add_argument("--odb", required=True, help="ODB file")
    ap.add_argument("--step-range", nargs="+", type=int, metavar="N",
                    help="step numbers in pairs FIRST LAST (inclusive); several ranges "
                         "allowed, e.g. 5 8 12 14 = steps 5 to 8 and 12 to 14")
    ap.add_argument("--step", nargs="+", metavar="STEP",
                    help="single steps: number in the ODB (see --list) or a unique part "
                         "of the name. Default if no step is given: the as-laid step")
    ap.add_argument("--z-step", nargs="+", metavar="STEP",
                    help="steps for the Z coordinate: step numbers / names, or 'all' for "
                         "every requested step (default: the as-laid step, i.e. the step "
                         "whose name contains 'aslaid' / 'as-laid'; else the first "
                         "requested step)")
    ap.add_argument("--force-key", metavar="NAME",
                    help="section force output to use (default ESF1; SF1 is used with a "
                         "warning if the ODB has no ESF1)")
    ap.add_argument("--vertical-axis", choices=["x", "y", "z"], default="z",
                    help="global vertical axis (default z)")
    ap.add_argument("--frame", type=int, default=-1,
                    help="frame used in each step (default -1 = last)")
    ap.add_argument("--elset", help="pipeline element set (default %s)" % es.ELSET)
    ap.add_argument("--instance", help="pipeline instance, if the ODB has more than one")
    ap.add_argument("--kp-start", type=float,
                    help="KP [m] at the start of the first pipeline element (default 0)")
    ap.add_argument("--out-dir", default=OUT_DIR,
                    help="output folder (default %s)" % OUT_DIR)
    ap.add_argument("--list", action="store_true",
                    help="print the steps and the available output variables, then stop")
    return ap.parse_args(argv)


def step_spec(text):
    return int(text) if re.match(r"^\d+$", str(text)) else text


def requested_steps(odb, ranges, singles):
    """Step names from --step-range and --step, in ODB order, without repeats."""
    names = list(odb.steps.keys())
    picked = set()
    if ranges:
        if len(ranges) % 2:
            raise ValueError("--step-range needs pairs of step numbers: FIRST LAST [FIRST LAST ...]")
        for a, b in zip(ranges[0::2], ranges[1::2]):
            if a > b:
                a, b = b, a
            if a < 1 or b > len(names):
                raise ValueError("Step range %d to %d is outside 1..%d" % (a, b, len(names)))
            picked.update(names[a - 1:b])
    for s in singles or []:
        picked.add(es.resolve_step(odb, step_spec(s)))
    return [n for n in names if n in picked]


def aslaid_step(odb):
    hits = [n for n in odb.steps.keys() if "aslaid" in re.sub(r"[^a-z]", "", n.lower())]
    return hits[-1] if hits else None


def read_element_scalar(frame, key, component, region, inst_name):
    """{element label: value of largest magnitude over its integration points}."""
    fo = frame.fieldOutputs[key].getSubset(region=region)
    labels_c = list(fo.componentLabels)
    idx = labels_c.index(component) if component and component in labels_c else 0
    out = {}

    def put(label, row):
        v = float(row[idx]) if hasattr(row, "__len__") else float(row)
        label = int(label)
        if label not in out or abs(v) > abs(out[label]):
            out[label] = v

    try:
        blocks = fo.bulkDataBlocks
    except AttributeError:
        blocks = None
    if blocks is not None:
        for b in blocks:
            if b.instance is not None and b.instance.name != inst_name:
                continue
            labels, data = b.elementLabels, b.data
            for k in range(len(labels)):
                put(labels[k], data[k])
        return out
    for v in fo.values:
        if v.instance is not None and v.instance.name != inst_name:
            continue
        put(v.elementLabel, v.data)
    return out


def force_source(frame, wanted):
    """(field output key, component, label) of the effective tension output."""
    keys = list(frame.fieldOutputs.keys())
    if wanted:
        for k in keys:
            if k.upper() == wanted.upper():
                return k, None, k
        if "SF" in keys and wanted.upper().startswith("SF"):
            return "SF", wanted.upper(), wanted.upper()
        raise KeyError("Output %s not found. Field outputs: %s" % (wanted, ", ".join(keys)))
    if "ESF1" in keys:
        return "ESF1", None, "ESF1"
    if "SF" in keys:
        print("  WARNING: the ODB has no ESF1 output; SF1 is used. SF1 is the axial force "
              "in the pipe wall, not the effective tension. Request ESF1 in *ELEMENT OUTPUT.")
        return "SF", "SF1", "SF1"
    raise KeyError("No ESF1 or SF output in the ODB. Field outputs: %s" % ", ".join(keys))


def read_z(frame, inst_name, nodes, node_xyz, axis):
    """({node: current coordinate along the vertical axis}, source text)."""
    wanted = set(nodes)
    keys = list(frame.fieldOutputs.keys())
    if "COORD" in keys:
        comp = "COOR%d" % (axis + 1)
        return es.read_nodal(frame, "COORD", comp, inst_name, wanted), "COORD output"
    if "U" in keys:
        vec = es.read_nodal_vector(frame, "U", inst_name, wanted)
        return (dict((n, node_xyz[n][axis] + u[axis]) for n, u in vec.items()),
                "initial coordinate + U")
    return dict((n, node_xyz[n][axis]) for n in nodes), "initial coordinate (no U output)"


def write_table(path, title_lines, columns, rows):
    """rows = [(kp, key)]; columns = [(name, {key: value})]."""
    fh = open(path, "w")
    for line in title_lines:
        fh.write(line + "\n")
    fh.write("\t".join(["KP [m]"] + [name for name, _d in columns]) + "\n")
    for kp, key in rows:
        fh.write("%.4f\t" % kp + "\t".join(
            "" if d.get(key) is None else "%.6E" % d[key] for _name, d in columns) + "\n")
    fh.close()


def main():
    args = parse_args(sys.argv[1:])
    if args.elset:
        es.ELSET = args.elset
    if args.instance:
        es.INSTANCE = args.instance
    if args.kp_start is not None:
        es.KP_START = args.kp_start
    axis = {"x": 0, "y": 1, "z": 2}[args.vertical_axis]
    from odbAccess import openOdb
    odb = openOdb(path=args.odb, readOnly=True)
    try:
        if args.list:
            print("Steps in ODB:")
            for i, name in enumerate(odb.steps.keys(), 1):
                print("  %2d  %s  (%d frames)" % (i, name, len(odb.steps[name].frames)))
            print("As-laid step found by name: %s" % (aslaid_step(odb) or "none"))
            last = odb.steps[list(odb.steps.keys())[-1]].frames[args.frame]
            keys = list(last.fieldOutputs.keys())
            for k in ("ESF1", "SF", "COORD", "U"):
                print("  %-6s %s" % (k, "available" if k in keys else "NOT in the ODB"))
            return
        inst, region, elems = es.get_region(odb)
        order, kp, node_kp, node_xyz = es.build_kp(inst, elems)
        steps = requested_steps(odb, args.step_range, args.step)
        laid = aslaid_step(odb)
        if not steps:
            if laid is None:
                raise ValueError("Give --step-range or --step (no as-laid step found by name).")
            steps = [laid]
        if args.z_step and len(args.z_step) == 1 and args.z_step[0].lower() == "all":
            z_steps = list(steps)
        elif args.z_step:
            z_steps = requested_steps(odb, None, args.z_step)
        else:
            z_steps = [laid] if laid else [steps[0]]
        if not os.path.isdir(args.out_dir):
            os.makedirs(args.out_dir)
        print("%d elements, KP %.2f to %.2f m" % (len(order), kp[order[0]], kp[order[-1]]))

        # effective tension
        columns, label = [], None
        for name in steps:
            frame = odb.steps[name].frames[args.frame]
            key, comp, label = force_source(frame, args.force_key)
            data = read_element_scalar(frame, key, comp, region, inst.name)
            columns.append((name, data))
            if data:
                el = max(data, key=lambda e: data[e])
                print("  %s  %s: max %.4g at KP %.2f m, min %.4g"
                      % (label, name, data[el], kp[el], min(data.values())))
        path = os.path.join(args.out_dir, "EFFTENSION.rpt")
        write_table(path,
                    ["EFFECTIVE TENSION along KP at frame %d of each step" % args.frame,
                     "ODB: %s" % os.path.abspath(args.odb),
                     "Output variable: %s, at element mid-length (largest magnitude over "
                     "the integration points)." % label,
                     "Units: KP in m, force in model force units."],
                    columns, [(kp[el], el) for el in order])
        print("  -> %s" % os.path.abspath(path))

        # Z coordinate
        nodes = sorted(node_kp, key=lambda n: node_kp[n])
        columns, source = [], ""
        for name in z_steps:
            data, source = read_z(odb.steps[name].frames[args.frame], inst.name,
                                  nodes, node_xyz, axis)
            columns.append((name, data))
            if data:
                print("  %s  %s: %.2f to %.2f" % (args.vertical_axis.upper(), name,
                                                    min(data.values()), max(data.values())))
        path = os.path.join(args.out_dir, "ZCOORD.rpt")
        write_table(path,
                    ["Z COORDINATE (water depth) along KP at frame %d of each step" % args.frame,
                     "ODB: %s" % os.path.abspath(args.odb),
                     "Global %s coordinate of the pipe nodes (%s)."
                     % (args.vertical_axis.upper(), source),
                     "Units: KP in m, coordinate in model length units."],
                    columns, [(node_kp[n], n) for n in nodes])
        print("  -> %s" % os.path.abspath(path))
    finally:
        odb.close()


if __name__ == "__main__":
    main()
