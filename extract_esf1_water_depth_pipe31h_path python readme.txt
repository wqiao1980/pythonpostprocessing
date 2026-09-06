PIPE31H ESF1 and Final As-Laid Water Depth - User Instructions
==============================================================

Purpose
-------

extract_esf1_water_depth_pipe31h_path.py outputs:

- cumulative pipeline path distance;
- node label;
- Water Depth Z at the final frame of a user-selected as-laid step; and
- ESF1 for every selected step, or every ODB step by default.

The path runs from one start node to one end node through connected PIPE31H
elements. The script is self-contained and does not require another Python
script. Version 2026-09-05-r3 creates both a text report and an Excel workbook
with native, editable charts.


Required as-laid step input
---------------------------

Every extraction run must include --aslaid-step. Supply either the exact step
name or its 1-based position in the ODB.

Exact as-laid step name:

abaqus python extract_esf1_water_depth_pipe31h_path.py --aslaid-step "Final As-Laid"

As-laid step by 1-based position, for example step 7:

abaqus python extract_esf1_water_depth_pipe31h_path.py --aslaid-step 7

The water depth is always taken from the final frame of this as-laid step. The
--frame-index option does not change the as-laid frame; it applies only to the
ESF1 steps.

For each output node, the script first uses COORD3 from the final as-laid frame.
If COORD is unavailable for that node, it uses original coordinate Z + U3 from
the same final frame. The Z sign and model length units are preserved. The
script reports which source was used.


Default behavior
----------------

When --steps and --step-range are both omitted, the script extracts ESF1 from
all ODB steps that contain ESF1. When no element selection is supplied, it
outputs all PIPE31H elements on the resolved start-to-end route.

For ESF1, the default is the last frame containing ESF1 in each step. Path
distance is cumulative 3D distance calculated from the original ODB nodal
coordinates. Water Depth Z comes only from the final frame of --aslaid-step.


Run the extraction
------------------

Run from an Abaqus Command Prompt. Abaqus/CAE or Viewer does not need to be
opened.

One or all ODBs in the script folder, all ESF1 steps:

abaqus python extract_esf1_water_depth_pipe31h_path.py --aslaid-step "Final As-Laid"

Process every ODB in another folder, all ESF1 steps:

abaqus python "C:\python_aba\takeoutSFSM\extract_esf1_water_depth_pipe31h_path.py" --input-dir "C:\Data\BPTiberFL6" --output-dir "C:\Data\BPTiberFL6\02_Results" --aslaid-step "Final As-Laid"

Process one ODB:

abaqus python extract_esf1_water_depth_pipe31h_path.py --input-dir "C:\Data\BPTiberFL6" --odb "model.odb" --aslaid-step "Final As-Laid"

The supplied as-laid name or position must be valid in every ODB processed by
the same command. If the ODBs use different as-laid step names or positions,
run them separately with the appropriate --aslaid-step value.


List steps
----------

Listing does not require --aslaid-step:

abaqus python extract_esf1_water_depth_pipe31h_path.py --list-steps

For one ODB:

abaqus python extract_esf1_water_depth_pipe31h_path.py --input-dir "C:\Data\BPTiberFL6" --odb "model.odb" --list-steps


Select ESF1 steps
-----------------

If neither option below is supplied, all ODB steps are considered for ESF1.
The as-laid step input is independent of this ESF1 selection.

Exact ESF1 step names:

abaqus python extract_esf1_water_depth_pipe31h_path.py --aslaid-step "Final As-Laid" --steps "Step-2" "Step-5"

Inclusive 1-based ESF1 step positions 3 through 7:

abaqus python extract_esf1_water_depth_pipe31h_path.py --aslaid-step "Final As-Laid" --step-range 3 7

Inclusive ESF1 range using exact first and last step names:

abaqus python extract_esf1_water_depth_pipe31h_path.py --aslaid-step "Final As-Laid" --step-range "Preload" "Operation"


Start and end nodes
-------------------

The script resolves each endpoint in this order:

1. An exact node label supplied by --start-node or --end-node.
2. An exact instance node-set name supplied by --start-node-set or
   --end-node-set.
3. A unique instance node-set name containing START or END, matched without
   regard to letter case.
4. If one endpoint is still missing, the farthest reachable graph endpoint is
   used.
5. If both are missing and the graph has exactly two connected endpoints,
   those endpoints are used, with the lower node label as the start.

If several node-set names contain START or END, specify the exact desired set.

Exact node labels:

abaqus python extract_esf1_water_depth_pipe31h_path.py --aslaid-step "Final As-Laid" --start-node 1 --end-node 5001

Exact node-set names:

abaqus python extract_esf1_water_depth_pipe31h_path.py --aslaid-step "Final As-Laid" --start-node-set "PIPE_START_NODE" --end-node-set "PIPE_END_NODE"

List possible endpoint node sets; listing does not require --aslaid-step:

abaqus python extract_esf1_water_depth_pipe31h_path.py --list-endpoint-sets


List PIPE31H element sets
-------------------------

Only instance- and assembly-level sets containing at least one PIPE31H element
in the selected instance are listed. Listing does not require --aslaid-step:

abaqus python extract_esf1_water_depth_pipe31h_path.py --instance "PART-1-1" --list-element-sets


Select element sets or labels
-----------------------------

One element set:

abaqus python extract_esf1_water_depth_pipe31h_path.py --aslaid-step "Final As-Laid" --element-set "SET_A"

Repeat --element-set to form a union:

abaqus python extract_esf1_water_depth_pipe31h_path.py --aslaid-step "Final As-Laid" --element-set "SET_A" --element-set "SET_B"

Both instance- and assembly-level element sets are supported. Set-name matching
is not case-sensitive.

One or more exact element labels:

abaqus python extract_esf1_water_depth_pipe31h_path.py --aslaid-step "Final As-Laid" --element 1001 1002 1005

An inclusive element-label range:

abaqus python extract_esf1_water_depth_pipe31h_path.py --aslaid-step "Final As-Laid" --element-range 1001 1200

Several ranges:

abaqus python extract_esf1_water_depth_pipe31h_path.py --aslaid-step "Final As-Laid" --element-range 1001 1200 --element-range 2001 2200

Element sets, exact labels, and ranges form a union:

abaqus python extract_esf1_water_depth_pipe31h_path.py --aslaid-step "Final As-Laid" --element-set "SET_A" --element 1500 --element-range 2001 2200 --steps "Step-2" "Step-5"

Only selected PIPE31H elements lying on the resolved start-to-end route are
reported. Non-PIPE31H elements and elements outside that route are omitted with
a note. Filtering does not reset path distance: selected segments retain their
distance measured from the full-route start node.


Instance and ESF1 frame options
-------------------------------

The default instance is PART-1-1. Select another with:

--instance "PIPELINE-1"

The default ESF1 frame selection is the last frame containing ESF1:

--frame-index -1

Use a particular zero-based ESF1 frame index:

--frame-index 1

Again, Water Depth Z always uses the final frame of --aslaid-step regardless of
--frame-index.


Output files
------------

For each ODB, the default output names are:

model.odb -> model_PIPE31H_ESF1_WATER_DEPTH_PATH.rpt
model.odb -> model_PIPE31H_ESF1_WATER_DEPTH_PATH.xlsx

The table contains Path Distance, Node Label, final as-laid Water Depth Z, and
one ESF1 column per selected/default step. A blank ESF1 cell means that no
selected element-nodal value was available at that node and frame.

Reports are written beside the ODB unless --output-dir is supplied. For a
single --odb run, use a custom text-report name with:

--output-name "custom_esf1_depth.rpt"

The Excel workbook keeps its ODB-based filename when --output-name is used.
Existing .rpt and .xlsx outputs from this script with the same names are
overwritten.


Excel workbook and plots
------------------------

The Path Data worksheet contains:

- Path Distance;
- Node Label;
- Water Depth Z from the final frame of the selected as-laid step; and
- one ESF1 column for every selected step, or every usable step by default.

It contains two native Excel scatter charts:

1. A combined dual-axis chart. All selected/default ESF1 step curves use the
   left axis, and final as-laid Water Depth Z uses the right axis.
2. A separate final as-laid Water Depth Z versus Path Distance chart.

Both plots use the selected element locations while retaining full-route
distance from the start node. They are native Excel objects, so users can
change chart types, colors, line styles, markers, titles, axes, legend, size,
and layout. Excel does not need to be installed or open while Abaqus creates
the workbook.


Diagnostic log and console capture
----------------------------------

Every processing run creates:

extract_esf1_water_depth_pipe31h_path.log

Capture complete console output with:

abaqus python "C:\python_aba\takeoutSFSM\extract_esf1_water_depth_pipe31h_path.py" --input-dir "C:\Data\BPTiberFL6" --output-dir "C:\Data\BPTiberFL6\02_Results" --aslaid-step "Final As-Laid" > "C:\Data\BPTiberFL6\02_Results\esf1_depth_console.txt" 2>&1

The > operator overwrites the console file. Use >> to append.


Troubleshooting
---------------

1. --aslaid-step is required or the as-laid step is not found

   Run --list-steps, then supply the exact step name or 1-based position. For a
   folder run, the same input must resolve in every ODB.

2. Final as-laid water depth cannot be obtained

   The final frame needs COORD output or U output for every reported path node.
   Request COORD or U field output in the analysis and rerun the ODB if neither
   is present. COORD3 is preferred; otherwise the script uses original Z + U3.

3. More than one START or END set is found

   Run --list-endpoint-sets, then specify the exact set with
   --start-node-set and/or --end-node-set.

4. No connected path exists between the endpoints

   Confirm both nodes belong to the same connected PIPE31H pipeline and that
   the correct --instance is selected.

5. No PIPE31H elements are found

   This script intentionally uses the exact Abaqus element type PIPE31H for
   path construction. Confirm the model element type and instance.

6. A selected set or range produces no output

   Run --list-element-sets or confirm the requested labels are PIPE31H elements
   lying on the resolved start-to-end route.

7. ESF1 cells are blank

   Confirm ESF1 was requested as field output for the selected PIPE31H elements
   and selected frame. The console reports the number of blank step/node cells.
