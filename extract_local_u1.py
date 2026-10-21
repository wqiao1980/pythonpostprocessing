# -*- coding: utf-8 -*-
"""
extract_local_u1.py
Extract the local U1 (axial displacement) of the pipeline along KP from an
Abaqus ODB, for the steps or step ranges you give.

The ODB stores displacements in the global system. Local U1 is the global
displacement projected on the as-laid pipe axis at each node (unit tangent
from the neighbouring nodes), positive towards increasing KP. With
--relative-to STEP the U1 of that step is subtracted, e.g. to see the axial
displacement since the as-laid step or the walking per cycle.

Run with Abaqus Python, in the folder that also holds extract_stress_ranges.py
and extract_tension_depth.py (this script uses their routines):

    abaqus python extract_local_u1.py --odb model.odb --list
    abaqus python extract_local_u1.py --odb model.odb --step 22 23
    abaqus python extract_local_u1.py --odb model.odb --step-range 20 23 --step 14
    abaqus python extract_local_u1.py --odb model.odb --step-range 20 23 --relative-to 5
    abaqus python extract_local_u1.py -h          # all options

Output, in the output folder (default pipeline_reports):
    LOCALU1.rpt   local U1 along KP (nodes), one column per step
"""
from __future__ import print_function
import argparse
import math
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import extract_stress_ranges as es
import extract_tension_depth as et

OUT_DIR = "pipeline_reports"


def parse_args(argv):
    ap = argparse.ArgumentParser(description="Local U1 (axial) along pipeline KP from an Abaqus ODB.")
    ap.add_argument("--odb", required=True, help="ODB file")
    ap.add_argument("--step-range", nargs="+", type=int, metavar="N",
                    help="step numbers in pairs FIRST LAST (inclusive); several ranges allowed")
    ap.add_argument("--step", nargs="+", metavar="STEP",
                    help="single steps: number in the ODB (see --list) or a unique part "
                         "of the name. Default if no step is given: the last step")
    ap.add_argument("--relative-to", metavar="STEP",
                    help="subtract the U1 of this step (number or part of the name), e.g. "
                         "the as-laid step")
    ap.add_argument("--u1-as-stored", action="store_true",
                    help="use the stored U1 component without converting it")
    ap.add_argument("--frame", type=int, default=-1,
                    help="frame used in each step (default -1 = last)")
    ap.add_argument("--elset", help="pipeline element set (default %s)" % es.ELSET)
    ap.add_argument("--instance", help="pipeline instance, if the ODB has more than one")
    ap.add_argument("--kp-start", type=float,
                    help="KP [m] at the start of the first pipeline element (default 0)")
    ap.add_argument("--out-dir", default=OUT_DIR,
                    help="output folder (default %s)" % OUT_DIR)
    ap.add_argument("--list", action="store_true",
                    help="print the steps and whether U is in the ODB, then stop")
    return ap.parse_args(argv)


def axial_directions(nodes, node_xyz):
    """{node: unit vector} along the as-laid pipe axis, towards increasing KP,
    from the neighbouring nodes (nodes ordered along KP)."""
    out = {}
    for i, n in enumerate(nodes):
        a = node_xyz[nodes[max(i - 1, 0)]]
        b = node_xyz[nodes[min(i + 1, len(nodes) - 1)]]
        t = [b[k] - a[k] for k in range(3)]
        norm = math.sqrt(sum(c * c for c in t))
        out[n] = [c / norm for c in t] if norm > 0.0 else [1.0, 0.0, 0.0]
    return out


def read_local_u1(frame, inst_name, nodes, axial, as_stored):
    """{node: local U1}: global U projected on the as-laid pipe axis."""
    wanted = set(nodes)
    if as_stored:
        return es.read_nodal(frame, "U", "U1", inst_name, wanted)
    vec = es.read_nodal_vector(frame, "U", inst_name, wanted)
    return dict((n, sum(u[k] * axial[n][k] for k in range(3))) for n, u in vec.items())


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
            print("  U    %s" % ("available" if "U" in keys else "NOT in the ODB"))
            return
        inst, _region, elems = es.get_region(odb)
        _order, _kp, node_kp, node_xyz = es.build_kp(inst, elems)
        steps = et.requested_steps(odb, args.step_range, args.step) or [names[-1]]
        nodes = sorted(node_kp, key=lambda n: node_kp[n])
        axial = axial_directions(nodes, node_xyz)
        ref, ref_name = None, None
        if args.relative_to:
            ref_name = es.resolve_step(odb, et.step_spec(args.relative_to))
            ref = read_local_u1(odb.steps[ref_name].frames[args.frame], inst.name, nodes,
                                axial, args.u1_as_stored)
            print("  U1 relative to %s" % ref_name)
        print("%d nodes, KP %.2f to %.2f m" % (len(nodes), node_kp[nodes[0]], node_kp[nodes[-1]]))
        columns = []
        for name in steps:
            frame = odb.steps[name].frames[args.frame]
            if "U" not in frame.fieldOutputs.keys():
                print("  %s: no U output - skipped" % name)
                continue
            data = read_local_u1(frame, inst.name, nodes, axial, args.u1_as_stored)
            if ref is not None:
                data = dict((n, v - ref[n]) for n, v in data.items() if n in ref)
            columns.append((name, data))
            if data:
                n1 = max(data, key=lambda n: abs(data[n]))
                print("  %s: largest |U1| = %.4g at KP %.2f m (U1 = %.4g)"
                      % (name, abs(data[n1]), node_kp[n1], data[n1]))
        if not columns:
            raise RuntimeError("No U field output in the requested steps.")
        if not os.path.isdir(args.out_dir):
            os.makedirs(args.out_dir)
        path = os.path.join(args.out_dir, "LOCALU1.rpt")
        es.write_node_table(
            path,
            ["LOCAL U1 along KP at frame %d of each step" % args.frame,
             "ODB: %s" % os.path.abspath(args.odb),
             "Stored U1 component of the ODB, not converted." if args.u1_as_stored else
             "Global U projected on the as-laid pipe axis; positive towards increasing KP.",
             "Relative to step: %s (its U1 is subtracted)." % ref_name if ref_name else
             "Total displacement (not relative to another step).",
             "Units: KP in m, U1 in model length units."],
            columns, nodes, node_kp)
        print("  -> %s" % os.path.abspath(path))
    finally:
        odb.close()


if __name__ == "__main__":
    main()
