# -*- coding: utf-8 -*-
"""
extract_contact.py
Extract the contact opening (COPEN) and contact pressure (CPRESS) along the
pipeline KP from an Abaqus ODB, for the whole pipeline or for chosen element
sets (e.g. a DBM middle set and a DBM shoulder set), and write one report per
variable and set. The reports are imported with button 7 of the fatigue
workbook, which plots them along KP with the DBM and curve sections shaded.

Run with Abaqus Python, in the folder that also holds extract_stress_ranges.py
(this script uses its KP and ODB routines):

    abaqus python extract_contact.py --odb model.odb --list
    abaqus python extract_contact.py --odb model.odb --step 22 23
    abaqus python extract_contact.py --odb model.odb --step 22 23 --elset DBM_MIDDLE DBM_SHOULDER
    abaqus python extract_contact.py -h          # all options

Output, in the output folder (default contact_reports):
    COPEN_<set>.rpt    contact opening along KP, one column per step
    CPRESS_<set>.rpt   contact pressure along KP, one column per step
"""
from __future__ import print_function
import argparse
import math
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import extract_stress_ranges as es

VARIABLES = ("COPEN", "CPRESS")
OUT_DIR = "contact_reports"


def parse_args(argv):
    ap = argparse.ArgumentParser(
        description="COPEN / CPRESS along pipeline KP from an Abaqus ODB.")
    ap.add_argument("--odb", required=True, help="ODB file")
    ap.add_argument("--step", nargs="+", metavar="STEP",
                    help="steps to report: number in the ODB (see --list) or a unique "
                         "part of the name (default: the last step)")
    ap.add_argument("--var", nargs="+", choices=VARIABLES, default=list(VARIABLES),
                    help="variables to extract (default: COPEN CPRESS)")
    ap.add_argument("--elset", nargs="+", metavar="NAME",
                    help="element sets to report, one report per set, e.g. DBM_MIDDLE "
                         "DBM_SHOULDER (default: the whole pipeline set)")
    ap.add_argument("--pipe-elset", metavar="NAME",
                    help="pipeline element set that defines KP (default %s)" % es.ELSET)
    ap.add_argument("--contact-key", metavar="TEXT",
                    help="use only the contact output variables whose name contains "
                         "TEXT, e.g. the seabed or DBM surface name (default: all)")
    ap.add_argument("--frame", type=int, default=-1,
                    help="frame used in each step (default -1 = last)")
    ap.add_argument("--instance", help="pipeline instance, if the ODB has more than one")
    ap.add_argument("--kp-start", type=float,
                    help="KP [m] at the start of the first pipeline element (default 0)")
    ap.add_argument("--out-dir", default=OUT_DIR,
                    help="output folder (default %s)" % OUT_DIR)
    ap.add_argument("--list", action="store_true",
                    help="print the steps, element sets and contact output variables, then stop")
    return ap.parse_args(argv)


def step_spec(text):
    return int(text) if re.match(r"^\d+$", text) else text


def find_set(odb, name):
    """[(instance name, element list)] of an element set (instance or assembly level)."""
    asm = odb.rootAssembly
    for iname in asm.instances.keys():
        inst = asm.instances[iname]
        for key in (name, name.upper()):
            if key in inst.elementSets.keys():
                return [(inst.name, list(inst.elementSets[key].elements))]
    for key in (name, name.upper()):
        if key in asm.elementSets.keys():
            region = asm.elementSets[key]
            names = list(region.instanceNames)
            if len(names) == 0:
                return [("", list(region.elements))]
            if len(names) == 1 and (len(region.elements) == 0
                                    or hasattr(region.elements[0], "connectivity")):
                return [(names[0], list(region.elements))]
            return [(names[i], list(region.elements[i])) for i in range(len(names))]
    raise KeyError("Element set %s not found. Run with --list to see the set names." % name)


def set_nodes(odb, parts, pipe_inst, node_kp, node_xyz):
    """{(instance name, node label): KP [m]} for the nodes of an element set.
    Pipeline nodes take their own KP; other nodes (e.g. a DBM surface) take the
    KP of the nearest pipeline node in the as-built model coordinates."""
    out, others = {}, {}
    for iname, elems in parts:
        for e in elems:
            for n in e.connectivity:
                if iname == pipe_inst and n in node_kp:
                    out[(iname, n)] = node_kp[n]
                else:
                    others.setdefault(iname, set()).add(n)
    if not others:
        return out, 0
    pipe = sorted(node_kp, key=lambda n: node_kp[n])
    pipe_xyz = [node_xyz[n] for n in pipe]
    try:
        import numpy as np
        arr = np.array(pipe_xyz)
    except ImportError:
        arr = None
    count = 0
    for iname, labels in others.items():
        nodes = odb.rootAssembly.instances[iname].nodes if iname else odb.rootAssembly.nodes
        for nd in nodes:
            if nd.label not in labels:
                continue
            c = [float(x) for x in nd.coordinates]
            if arr is not None:
                k = int(((arr - np.array(c)) ** 2).sum(axis=1).argmin())
            else:
                k = min(range(len(pipe_xyz)),
                        key=lambda i: sum((pipe_xyz[i][j] - c[j]) ** 2 for j in range(3)))
            out[(iname, nd.label)] = node_kp[pipe[k]]
            count += 1
    return out, count


def contact_keys(frame, var, key_filter):
    return [k for k in frame.fieldOutputs.keys() if k.upper().startswith(var)
            and (not key_filter or key_filter.lower() in k.lower())]


def read_contact(frame, var, wanted, key_filter):
    """{(instance, node): value}: smallest COPEN / largest CPRESS over the
    contact output variables of the frame at the wanted nodes."""
    keys = contact_keys(frame, var, key_filter)
    keep_smaller = (var == "COPEN")
    out = {}

    def put(iname, label, value):
        key = (iname, int(label))
        if key not in wanted:
            return
        value = float(value[0]) if hasattr(value, "__len__") else float(value)
        if key not in out or (value < out[key] if keep_smaller else value > out[key]):
            out[key] = value

    for k in keys:
        fo = frame.fieldOutputs[k]
        try:
            blocks = fo.bulkDataBlocks
        except AttributeError:
            blocks = None
        if blocks is not None:
            for b in blocks:
                iname = b.instance.name if b.instance is not None else ""
                labels, data = b.nodeLabels, b.data
                for i in range(len(labels)):
                    put(iname, labels[i], data[i])
        else:
            for v in fo.values:
                put(v.instance.name if v.instance is not None else "", v.nodeLabel, v.data)
    return out, keys


def by_kp(values, node_kp_map, var):
    """{KP: value}; several nodes at one KP give the smallest COPEN / largest CPRESS."""
    out = {}
    for key, v in values.items():
        kp = round(node_kp_map[key], 4)
        if kp not in out or (v < out[kp] if var == "COPEN" else v > out[kp]):
            out[kp] = v
    return out


def gap_rows(rows):
    """KP values where a blank row is written: the middle of every gap much
    longer than the usual node spacing (5 x the median), so that separate
    pieces of a set are plotted as separate line segments."""
    if len(rows) < 3:
        return []
    steps = sorted(b - a for a, b in zip(rows[:-1], rows[1:]))
    limit = 5.0 * steps[len(steps) // 2]
    return [(a + b) / 2.0 for a, b in zip(rows[:-1], rows[1:]) if b - a > limit > 0.0]


def safe_name(text):
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", text).strip("_")


def list_odb(odb, frame_index):
    asm = odb.rootAssembly
    print("Steps in ODB:")
    for i, name in enumerate(odb.steps.keys(), 1):
        print("  %2d  %s  (%d frames)" % (i, name, len(odb.steps[name].frames)))
    print("Element sets:")
    for iname in asm.instances.keys():
        for s in sorted(asm.instances[iname].elementSets.keys()):
            print("  %-40s (instance %s)" % (s, iname))
    for s in sorted(asm.elementSets.keys()):
        print("  %-40s (assembly)" % s)
    last = odb.steps[list(odb.steps.keys())[-1]].frames[frame_index]
    print("Contact output variables in the last step:")
    found = False
    for var in VARIABLES:
        for k in contact_keys(last, var, None):
            print("  %s" % k)
            found = True
    if not found:
        print("  none - request COPEN and CPRESS in *CONTACT OUTPUT")


def main():
    args = parse_args(sys.argv[1:])
    if args.pipe_elset:
        es.ELSET = args.pipe_elset
    if args.instance:
        es.INSTANCE = args.instance
    if args.kp_start is not None:
        es.KP_START = args.kp_start
    from odbAccess import openOdb
    odb = openOdb(path=args.odb, readOnly=True)
    try:
        if args.list:
            list_odb(odb, args.frame)
            return
        inst, _region, elems = es.get_region(odb)
        _order, _kp, node_kp, node_xyz = es.build_kp(inst, elems)
        names = list(odb.steps.keys())
        steps = [es.resolve_step(odb, step_spec(s)) for s in args.step] if args.step \
            else [names[-1]]
        set_names = args.elset if args.elset else [es.ELSET]
        if not os.path.isdir(args.out_dir):
            os.makedirs(args.out_dir)
        for set_name in set_names:
            if args.elset:
                parts = find_set(odb, set_name)
            else:
                parts = [(inst.name, elems)]
            kp_map, mapped = set_nodes(odb, parts, inst.name, node_kp, node_xyz)
            wanted = set(kp_map)
            print("Set %s: %d nodes, KP %.2f to %.2f m%s"
                  % (set_name, len(wanted), min(kp_map.values()), max(kp_map.values()),
                     "" if not mapped else
                     " (%d nodes are not pipeline nodes: KP of the nearest pipeline node)"
                     % mapped))
            for var in args.var:
                columns, used_keys = [], []
                for name in steps:
                    data, keys = read_contact(odb.steps[name].frames[args.frame], var,
                                              wanted, args.contact_key)
                    for k in keys:
                        if k not in used_keys:
                            used_keys.append(k)
                    columns.append((name, by_kp(data, kp_map, var)))
                rows = sorted(set(kp for _n, d in columns for kp in d))
                if not rows:
                    print("  %s: no output at the nodes of this set in the requested "
                          "steps - report not written" % var)
                    continue
                path = os.path.join(args.out_dir, "%s_%s.rpt" % (var, safe_name(set_name)))
                fh = open(path, "w")
                fh.write("%s along KP at frame %d of each step\n" % (var, args.frame))
                fh.write("Set: %s\n" % set_name)
                fh.write("ODB: %s\n" % os.path.abspath(args.odb))
                fh.write("Contact output: %s (%s value over these at each node)\n"
                         % ("; ".join(used_keys), "smallest" if var == "COPEN" else "largest"))
                fh.write("Units: KP in m, %s in model %s units.\n"
                         % (var, "length" if var == "COPEN" else "pressure"))
                fh.write("\t".join(["KP [m]"] + [n for n, _d in columns]) + "\n")
                for kp in sorted(rows + gap_rows(rows)):
                    fh.write("%.4f\t" % kp + "\t".join(
                        "" if d.get(kp) is None else "%.6E" % d[kp] for _n, d in columns) + "\n")
                fh.close()
                for name, d in columns:
                    if not d:
                        print("  %s  %s: no output" % (var, name))
                        continue
                    if var == "COPEN":
                        kp = max(d, key=lambda k: d[k])
                        print("  COPEN   %s: largest opening %.4g at KP %.2f m" % (name, d[kp], kp))
                    else:
                        kp = max(d, key=lambda k: d[k])
                        print("  CPRESS  %s: largest pressure %.4g at KP %.2f m" % (name, d[kp], kp))
                print("  -> %s (%d KP rows)" % (os.path.abspath(path), len(rows)))
    finally:
        odb.close()


if __name__ == "__main__":
    main()
