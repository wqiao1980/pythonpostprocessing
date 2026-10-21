# -*- coding: utf-8 -*-
"""
extract_local_u2.py
Extract the local U2 (lateral displacement) of the pipeline along KP from an
Abaqus ODB, for the steps or step ranges you give, and write a report that
PipelineResultPlots.xlsm imports and plots.

The ODB stores displacements in the global system. Local U2 is the global
displacement projected on the horizontal direction normal to the pipe axis at
each node (positive to the left when looking towards increasing KP). The axis
is the original mesh, or the deformed shape at the step given with
--aslaid-step. With --relative-to STEP the U2 of that step is subtracted.

Run with Abaqus Python, in the folder that also holds extract_stress_ranges.py
and extract_tension_depth.py (this script uses their routines):

    abaqus python extract_local_u2.py --odb model.odb --list
    abaqus python extract_local_u2.py --odb model.odb --step 22 23
    abaqus python extract_local_u2.py --odb model.odb --step-range 20 23 --step 14
    abaqus python extract_local_u2.py --odb model.odb --step 22 23 --aslaid-step 5 --relative-to 5
    abaqus python extract_local_u2.py -h          # all options

Output, in the output folder (default pipeline_reports):
    LOCALU2.rpt   local U2 along KP (nodes), one column per step
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
    ap = argparse.ArgumentParser(description="Local U2 along pipeline KP from an Abaqus ODB.")
    ap.add_argument("--odb", required=True, help="ODB file")
    ap.add_argument("--step-range", nargs="+", type=int, metavar="N",
                    help="step numbers in pairs FIRST LAST (inclusive); several ranges allowed")
    ap.add_argument("--step", nargs="+", metavar="STEP",
                    help="single steps: number in the ODB (see --list) or a unique part "
                         "of the name. Default if no step is given: the last step")
    ap.add_argument("--vertical-axis", choices=["x", "y", "z"],
                    help="global vertical axis used for the conversion to local U2 "
                         "(default %s)" % es.VERTICAL_AXIS)
    ap.add_argument("--relative-to", metavar="STEP",
                    help="subtract the U2 of this step (number or part of the name), e.g. "
                         "the as-laid step")
    ap.add_argument("--u2-as-stored", action="store_true",
                    help="use the stored U2 component without converting it")
    ap.add_argument("--aslaid-step", metavar="STEP",
                    help="as-laid step (number or part of the name): the pipe axis that "
                         "defines the local directions is taken from the deformed shape at "
                         "this step. Default: the original mesh")
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


def axis_coordinates(odb, spec, frame_index, inst_name, nodes, node_xyz):
    """(node coordinates that define the pipe axis, description). spec = the
    as-laid step: its deformed shape (original coordinates + U) is used.
    No spec: the original mesh."""
    if not spec:
        return node_xyz, "original mesh"
    name = es.resolve_step(odb, et.step_spec(spec))
    frame = odb.steps[name].frames[frame_index]
    if "U" not in frame.fieldOutputs.keys():
        raise KeyError("No U output in the as-laid step %s" % name)
    vec = es.read_nodal_vector(frame, "U", inst_name, set(nodes))
    xyz = dict(node_xyz)
    for n, u in vec.items():
        xyz[n] = [node_xyz[n][k] + u[k] for k in range(3)]
    return xyz, "deformed shape at the as-laid step %s" % name


def main():
    args = parse_args(sys.argv[1:])
    if args.elset:
        es.ELSET = args.elset
    if args.instance:
        es.INSTANCE = args.instance
    if args.kp_start is not None:
        es.KP_START = args.kp_start
    if args.vertical_axis:
        es.VERTICAL_AXIS = args.vertical_axis
    if args.u2_as_stored:
        es.U2_AS_STORED = True
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
        axis_xyz, axis_text = axis_coordinates(odb, args.aslaid_step, args.frame, inst.name,
                                               nodes, node_xyz)
        print("  pipe axis: %s" % axis_text)
        lateral = es.lateral_directions(nodes, axis_xyz)
        ref, ref_name = None, None
        if args.relative_to:
            ref_name = es.resolve_step(odb, et.step_spec(args.relative_to))
            ref = es.read_local_u2(odb.steps[ref_name].frames[args.frame], inst.name, nodes,
                                   axis_xyz, lateral)
            print("  U2 relative to %s" % ref_name)
        print("%d nodes, KP %.2f to %.2f m" % (len(nodes), node_kp[nodes[0]], node_kp[nodes[-1]]))
        columns = []
        for name in steps:
            frame = odb.steps[name].frames[args.frame]
            if "U" not in frame.fieldOutputs.keys():
                print("  %s: no U output - skipped" % name)
                continue
            data = es.read_local_u2(frame, inst.name, nodes, axis_xyz, lateral)
            if ref is not None:
                data = dict((n, v - ref[n]) for n, v in data.items() if n in ref)
            columns.append((name, data))
            if data:
                n1 = max(data, key=lambda n: abs(data[n]))
                print("  %s: largest |U2| = %.4g at KP %.2f m (U2 = %.4g)"
                      % (name, abs(data[n1]), node_kp[n1], data[n1]))
        if not columns:
            raise RuntimeError("No U field output in the requested steps.")
        if not os.path.isdir(args.out_dir):
            os.makedirs(args.out_dir)
        path = os.path.join(args.out_dir, "LOCALU2.rpt")
        es.write_node_table(
            path,
            ["LOCAL U2 along KP at frame %d of each step" % args.frame,
             "ODB: %s" % os.path.abspath(args.odb),
             "Stored U2 component of the ODB, not converted." if es.U2_AS_STORED else
             "Global U projected on the horizontal direction normal to the pipe axis (%s); "
             "positive to the left looking towards increasing KP (vertical axis: global %s)."
             % (axis_text, es.VERTICAL_AXIS.upper()),
             "Relative to step: %s (its U2 is subtracted)." % ref_name if ref_name else
             "Total displacement (not relative to another step).",
             "Units: KP in m, U2 in model length units."],
            columns, nodes, node_kp)
        print("  -> %s" % os.path.abspath(path))
    finally:
        odb.close()


if __name__ == "__main__":
    main()
