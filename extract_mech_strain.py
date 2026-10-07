# -*- coding: utf-8 -*-
"""
extract_mech_strain.py
Extract the maximum and minimum mechanical strain along pipeline KP from an
Abaqus ODB. Mechanical strain = LE11 - THE11, taken over ALL section points
(and integration points) of each pipe element. With --var you can also (or
instead) output LE11 or THE11 themselves. One report per variable, with a MAX
and a MIN column per step, imported and plotted by PipelineResultPlots.xlsm.

Run with Abaqus Python, in the folder that also holds extract_stress_ranges.py
and extract_tension_depth.py (this script uses their routines):

    abaqus python extract_mech_strain.py --odb model.odb --list
    abaqus python extract_mech_strain.py --odb model.odb --step-range 20 23
    abaqus python extract_mech_strain.py --odb model.odb --step 22 23 --all-frames
    abaqus python extract_mech_strain.py --odb model.odb --step 22 23 --var MECH LE11 THE11
    abaqus python extract_mech_strain.py -h          # all options

Output, in the output folder (default pipeline_reports):
    MECHSTRAIN.rpt   KP, then MAX and MIN mechanical strain per step   (--var MECH)
    LE11.rpt         same for the total logarithmic strain LE11        (--var LE11)
    THE11.rpt        same for the thermal strain THE11                 (--var THE11)
"""
from __future__ import print_function
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import extract_stress_ranges as es
import extract_tension_depth as et

OUT_DIR = "pipeline_reports"


def parse_args(argv):
    ap = argparse.ArgumentParser(
        description="Max / min mechanical strain (LE11 - THE11) along pipeline KP.")
    ap.add_argument("--odb", required=True, help="ODB file")
    ap.add_argument("--step-range", nargs="+", type=int, metavar="N",
                    help="step numbers in pairs FIRST LAST (inclusive); several ranges allowed")
    ap.add_argument("--step", nargs="+", metavar="STEP",
                    help="single steps: number in the ODB (see --list) or a unique part "
                         "of the name. Default if no step is given: the last step")
    ap.add_argument("--var", nargs="+", choices=["MECH", "LE11", "THE11"], default=["MECH"],
                    help="what to output: MECH = LE11 - THE11 (default), LE11, THE11; "
                         "several allowed, one report each")
    ap.add_argument("--all-frames", action="store_true",
                    help="maximum and minimum over ALL frames of each step instead of "
                         "one frame")
    ap.add_argument("--frame", type=int, default=-1,
                    help="frame used in each step (default -1 = last)")
    ap.add_argument("--no-thermal", action="store_true",
                    help="use LE11 alone when the ODB has no THE output")
    ap.add_argument("--elset", help="pipeline element set (default %s)" % es.ELSET)
    ap.add_argument("--instance", help="pipeline instance, if the ODB has more than one")
    ap.add_argument("--kp-start", type=float,
                    help="KP [m] at the start of the first pipeline element (default 0)")
    ap.add_argument("--out-dir", default=OUT_DIR,
                    help="output folder (default %s)" % OUT_DIR)
    ap.add_argument("--list", action="store_true",
                    help="print the steps and whether LE and THE are in the ODB, then stop")
    return ap.parse_args(argv)


def read_all_points(frame, key, component, region, inst_name):
    """{(element, integration point, section point number): component value}
    at every section point of the region."""
    fo = frame.fieldOutputs[key].getSubset(region=region)
    idx = list(fo.componentLabels).index(component)
    out = {}
    try:
        blocks = fo.bulkDataBlocks
    except AttributeError:
        blocks = None
    if blocks is not None:
        for b in blocks:
            if b.instance is not None and b.instance.name != inst_name:
                continue
            sp = b.sectionPoint.number if b.sectionPoint is not None else 0
            labels, ips, data = b.elementLabels, b.integrationPoints, b.data
            for k in range(len(labels)):
                row = data[k]
                out[(int(labels[k]), int(ips[k]), sp)] = \
                    float(row[idx]) if hasattr(row, "__len__") else float(row)
        return out
    for v in fo.values:
        if v.instance is not None and v.instance.name != inst_name:
            continue
        sp = v.sectionPoint.number if v.sectionPoint is not None else 0
        row = v.data
        out[(v.elementLabel, v.integrationPoint, sp)] = \
            float(row[idx]) if hasattr(row, "__len__") else float(row)
    return out


def frame_strain(frame, region, inst_name, wanted, no_thermal):
    """({variable: {point key: value}}, number of points without THE11) for the
    wanted variables (MECH = LE11 - THE11, LE11, THE11)."""
    keys = list(frame.fieldOutputs.keys())
    out, missing, le, the = {}, 0, None, None
    if "MECH" in wanted or "LE11" in wanted:
        if "LE" not in keys:
            raise KeyError("No LE output in this frame. Request LE in *ELEMENT OUTPUT.")
        le = read_all_points(frame, "LE", "LE11", region, inst_name)
    need_the = "THE11" in wanted or ("MECH" in wanted and not no_thermal)
    if need_the:
        if "THE" not in keys:
            raise KeyError("No THE output in the ODB. Request THE in *ELEMENT OUTPUT, or "
                           "run with --no-thermal (and without --var THE11) to use LE11 alone.")
        the = read_all_points(frame, "THE", "THE11", region, inst_name)
    if "LE11" in wanted:
        out["LE11"] = le
    if "THE11" in wanted:
        out["THE11"] = the
    if "MECH" in wanted:
        if the is None:
            out["MECH"] = le
        else:
            missing = sum(1 for k in le if k not in the)
            out["MECH"] = dict((k, v - the.get(k, 0.0)) for k, v in le.items())
    return out, missing


def main():
    args = parse_args(sys.argv[1:])
    if args.elset:
        es.ELSET = args.elset
    if args.instance:
        es.INSTANCE = args.instance
    if args.kp_start is not None:
        es.KP_START = args.kp_start
    from odbAccess import openOdb
    odb = openOdb(path=args.odb, readOnly=True)
    try:
        names = list(odb.steps.keys())
        if args.list:
            print("Steps in ODB:")
            for i, name in enumerate(names, 1):
                print("  %2d  %s  (%d frames)" % (i, name, len(odb.steps[name].frames)))
            keys = list(odb.steps[names[-1]].frames[args.frame].fieldOutputs.keys())
            for k in ("LE", "THE"):
                print("  %-4s %s" % (k, "available" if k in keys else "NOT in the ODB"))
            return
        inst, region, elems = es.get_region(odb)
        order, kp, _node_kp, _xyz = es.build_kp(inst, elems)
        steps = et.requested_steps(odb, args.step_range, args.step) or [names[-1]]
        if not os.path.isdir(args.out_dir):
            os.makedirs(args.out_dir)
        print("%d elements, KP %.2f to %.2f m" % (len(order), kp[order[0]], kp[order[-1]]))
        wanted = [v for v in ("MECH", "LE11", "THE11") if v in args.var]
        columns = dict((v, []) for v in wanted)
        points = 0
        for name in steps:
            frames = odb.steps[name].frames
            use = list(frames) if args.all_frames else [frames[args.frame]]
            hi = dict((v, {}) for v in wanted)
            lo = dict((v, {}) for v in wanted)
            where = dict((v, {}) for v in wanted)
            for frame in use:
                if args.all_frames and "LE" not in frame.fieldOutputs.keys():
                    continue
                data, missing = frame_strain(frame, region, inst.name, wanted, args.no_thermal)
                if missing:
                    print("  WARNING %s: %d points have LE11 but no THE11 (THE11 taken as 0)"
                          % (name, missing))
                for var in wanted:
                    h, l, w = hi[var], lo[var], where[var]
                    points = max(points, len(set(k[2] for k in data[var])))
                    for (el, _ip, sp), v in data[var].items():
                        if el not in h or v > h[el]:
                            h[el] = v
                            w[el] = sp
                        if el not in l or v < l[el]:
                            l[el] = v
            for var in wanted:
                h, l = hi[var], lo[var]
                columns[var].append(("MAX " + name, h))
                columns[var].append(("MIN " + name, l))
                if h:
                    e1 = max(h, key=lambda e: h[e])
                    e2 = min(l, key=lambda e: l[e])
                    print("  %-5s %s: max %.4f %% at KP %.2f m (section point %d), "
                          "min %.4f %% at KP %.2f m"
                          % (var, name, 100.0 * h[e1], kp[e1], where[var][e1],
                             100.0 * l[e2], kp[e2]))
        titles = {"MECH": ("MECHSTRAIN.rpt", "MECHANICAL STRAIN along KP: maximum and minimum "
                           "of LE11 - THE11" + (" (THE11 not subtracted, --no-thermal)"
                                                if args.no_thermal else "")),
                  "LE11": ("LE11.rpt", "LE11 along KP: maximum and minimum of the total "
                           "logarithmic strain LE11"),
                  "THE11": ("THE11.rpt", "THE11 along KP: maximum and minimum of the thermal "
                            "strain THE11")}
        for var in wanted:
            path = os.path.join(args.out_dir, titles[var][0])
            et.write_table(path,
                           [titles[var][1] + " over all section points (%d per element "
                            "found) and integration points" % points,
                            "ODB: %s" % os.path.abspath(args.odb),
                            "Frames: %s." % ("all frames of each step" if args.all_frames
                                             else "frame %d of each step" % args.frame),
                            "Units: KP in m, strain as a fraction (0.001 = 0.1 %)."],
                           columns[var], [(kp[el], el) for el in order])
            print("  -> %s" % os.path.abspath(path))
    finally:
        odb.close()


if __name__ == "__main__":
    main()
