PIPE31H ESF1 and Water Depth Along Path - User Instructions
===========================================================

Purpose
-------

extract_esf1_water_depth_pipe31h_path.py outputs:

- cumulative pipeline path distance;
- node label;
- Water Depth Z (the original ODB nodal coordinate Z); and
- ESF1 for every selected step.

The path runs from one start node to one end node through connected PIPE31H
elements. The script is self-contained and does not require any other Python
script.


Default behavior
----------------

When no step or element selection is supplied, the script:

- processes every step containing ESF1;
- uses the last frame containing ESF1 in each step; and
- outputs all PIPE31H elements on the resolved start-to-end route.

Water Depth Z is copied directly from the original ODB nodal coordinates. Its
sign and model length units are not changed. Path distance is cumulative 3D
distance calculated from the same original nodal coordinates.


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

abaqus python extract_esf1_water_depth_pipe31h_path.py --start-node 1 --end-node 5001

Exact node-set names:

abaqus python extract_esf1_water_depth_pipe31h_path.py --start-node-set "PIPE_START_NODE" --end-node-set "PIPE_END_NODE"

List possible endpoint node sets:

abaqus python extract_esf1_water_depth_pipe31h_path.py --list-endpoint-sets


Run the default extraction
--------------------------

Run from an Abaqus Command Prompt:

abaqus python extract_esf1_water_depth_pipe31h_path.py

The Abaqus/CAE or Viewer window does not need to be opened.

Process every ODB in another folder:

abaqus python "C:\python_aba\takeoutSFSM\extract_esf1_water_depth_pipe31h_path.py" --input-dir "C:\Data\BPTiberFL6" --output-dir "C:\Data\BPTiberFL6\02_Results"


List steps
----------

abaqus python extract_esf1_water_depth_pipe31h_path.py --list-steps

For one ODB:

abaqus python extract_esf1_water_depth_pipe31h_path.py --input-dir "C:\Data\BPTiberFL6" --odb "model.odb" --list-steps


Select steps
------------

Exact step names:

abaqus python extract_esf1_water_depth_pipe31h_path.py --steps "Step-2" "Step-5"

Inclusive 1-based step positions 3 through 7:

abaqus python extract_esf1_water_depth_pipe31h_path.py --step-range 3 7

Inclusive range using exact first and last step names:

abaqus python extract_esf1_water_depth_pipe31h_path.py --step-range "Preload" "Operation"


List PIPE31H element sets
-------------------------

Only instance- and assembly-level sets containing at least one PIPE31H
element in the selected instance are listed:

abaqus python extract_esf1_water_depth_pipe31h_path.py --instance "PART-1-1" --list-element-sets


Select element sets
-------------------

abaqus python extract_esf1_water_depth_pipe31h_path.py --element-set "SET_A"

Repeat the option to form a union:

abaqus python extract_esf1_water_depth_pipe31h_path.py --element-set "SET_A" --element-set "SET_B"

Both instance- and assembly-level element sets are supported. Set-name
matching is not case-sensitive.


Select exact element labels
---------------------------

abaqus python extract_esf1_water_depth_pipe31h_path.py --element 1001

Several labels may follow one option:

abaqus python extract_esf1_water_depth_pipe31h_path.py --element 1001 1002 1005

The option may also be repeated.


Select element-label ranges
---------------------------

The first and last labels are both included:

abaqus python extract_esf1_water_depth_pipe31h_path.py --element-range 1001 1200

Repeat the option for several ranges:

abaqus python extract_esf1_water_depth_pipe31h_path.py --element-range 1001 1200 --element-range 2001 2200


Combine element selections
--------------------------

Element sets, exact labels, and ranges form a union:

abaqus python extract_esf1_water_depth_pipe31h_path.py --element-set "SET_A" --element 1500 --element-range 2001 2200 --steps "Step-2" "Step-5"

Only selected PIPE31H elements lying on the resolved start-to-end route are
reported. Non-PIPE31H elements and PIPE31H elements outside that route are
omitted with a note. Filtering does not reset path distance: selected segments
retain their distance measured from the full-route start node.


Instance and frame options
--------------------------

The default instance is PART-1-1. Use another instance with:

--instance "PIPELINE-1"

The default frame selection is the last frame containing ESF1:

--frame-index -1

Use a particular zero-based frame index:

--frame-index 1


Output files
------------

For each ODB, the default report name is:

model.odb -> model_PIPE31H_ESF1_WATER_DEPTH_PATH.rpt

The table contains one ESF1 column per selected step. A blank ESF1 cell means
that no selected element-nodal value was available at that node and frame.

Reports are written beside the ODB unless --output-dir is supplied. For a
single --odb run, use a custom report name with:

--output-name "custom_esf1_depth.rpt"


Diagnostic log and console capture
----------------------------------

Every processing run creates:

extract_esf1_water_depth_pipe31h_path.log

Capture complete console output with:

abaqus python "C:\python_aba\takeoutSFSM\extract_esf1_water_depth_pipe31h_path.py" --input-dir "C:\Data\BPTiberFL6" --output-dir "C:\Data\BPTiberFL6\02_Results" > "C:\Data\BPTiberFL6\02_Results\esf1_depth_console.txt" 2>&1

The > operator overwrites the console file. Use >> to append.


Troubleshooting
---------------

1. More than one START or END set is found

   Run --list-endpoint-sets, then specify the exact set with
   --start-node-set and/or --end-node-set.

2. No connected path exists between the endpoints

   Confirm both nodes belong to the same connected PIPE31H pipeline and that
   the correct --instance is selected.

3. No PIPE31H elements are found

   This script intentionally uses the exact Abaqus element type PIPE31H for
   path construction. Confirm the model element type and instance.

4. A selected set/range produces no output

   Run --list-element-sets or confirm the requested labels are PIPE31H elements
   lying on the resolved start-to-end route.

5. ESF1 cells are blank

   Confirm ESF1 was requested as field output for the selected PIPE31H elements
   and selected frame. The console reports the number of blank step/node cells.
