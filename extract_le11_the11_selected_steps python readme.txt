Abaqus LE11/THE11/MECH11 Pipeline Report - User Instructions
============================================================

Required file
-------------

Keep this Python file in the working folder:

extract_le11_the11_selected_steps.py

Run the commands below from an Abaqus Command Prompt. Use "abaqus python";
the Abaqus/CAE or Viewer window does not need to be opened.


Results produced by the script
------------------------------

For every processed ODB, the script creates two files:

1. model_LE11_THE11_MECH11.rpt
2. model_MAX_MECH11_along_path.xlsx

The report contains results along the complete pipeline path for every
selected step:

- LE11
- THE11
- MECH11 = LE11 - THE11
- MAX_MECH11, the maximum MECH11 among all section points at each path node

By default, every section point is preserved. Each section point has three
adjacent report columns: LE11, THE11, and MECH11.

The Excel workbook contains the path distance, node label, and one
MAX_MECH11 column for every selected step. A native Excel scatter chart plots
all selected steps along the pipeline path. Each disconnected path has its
own worksheet and chart.

The Excel chart is editable. In Excel, users can change its chart type,
colors, line styles, markers, titles, axes, legend, size, and layout. The
chart is linked to the visible worksheet data, so changing a value updates
the chart. Excel does not need to be installed or open when Abaqus creates
the workbook.


Process all ODBs and all steps
------------------------------

The following command processes every ODB in the current folder and every
analysis step containing the required LE and THE field outputs:

abaqus python extract_le11_the11_selected_steps.py

To process every ODB in another folder:

abaqus python "C:\python_aba\takeoutSFSM\extract_le11_the11_selected_steps.py" --input-dir "C:\Data\P1-2\00_NewFA\12inTRF"


List the available steps
------------------------

This command lists the ordered steps in every ODB in the current folder. It
does not create reports or Excel workbooks.

abaqus python extract_le11_the11_selected_steps.py --list-steps

To list the steps in ODBs located in another folder:

abaqus python "C:\python_aba\takeoutSFSM\extract_le11_the11_selected_steps.py" --input-dir "C:\Data\P1-2\00_NewFA\12inTRF" --list-steps


Output selected steps by name
-----------------------------

List each exact step name after --steps. Put quotation marks around step
names that contain spaces.

abaqus python extract_le11_the11_selected_steps.py --steps "Step-2" "Step-5" "Step-8"

Step-name matching is not case-sensitive. A requested step that is not in an
ODB is omitted, and the script prints a note.


Output a range using step positions
-----------------------------------

Step positions are 1-based and the range is inclusive. The following example
outputs steps 3, 4, 5, 6, and 7:

abaqus python extract_le11_the11_selected_steps.py --step-range 3 7

If an ODB has fewer steps than the requested ending position, the script uses
the last available step and prints a note.


Output a range using step names
-------------------------------

The first and last named steps are both included:

abaqus python extract_le11_the11_selected_steps.py --step-range "Preload" "Operation"


Process selected steps in the 12inTRF folder
--------------------------------------------

Example: output steps 3 through 7 from every ODB in 12inTRF:

abaqus python "C:\python_aba\takeoutSFSM\extract_le11_the11_selected_steps.py" --input-dir "C:\Data\P1-2\00_NewFA\12inTRF" --step-range 3 7

Example: output selected named steps from every ODB in 12inTRF:

abaqus python "C:\python_aba\takeoutSFSM\extract_le11_the11_selected_steps.py" --input-dir "C:\Data\P1-2\00_NewFA\12inTRF" --steps "Step-2" "Step-5"


Process one ODB only
--------------------

Add --odb followed by the ODB filename:

abaqus python extract_le11_the11_selected_steps.py --odb "model.odb" --step-range 3 7

If the ODB is in another folder, combine --input-dir and --odb:

abaqus python "C:\python_aba\takeoutSFSM\extract_le11_the11_selected_steps.py" --input-dir "C:\Data\P1-2\00_NewFA\12inTRF" --odb "model.odb" --steps "Step-2" "Step-5"

The --odb option may be repeated to process several named ODB files.


Path distance and starting node
-------------------------------

Path distance is the cumulative distance along the undeformed pipeline
line-element mesh. It is calculated from the nodal coordinates in the ODB,
so its units are the model length units.

The default pipeline instance is:

PART-1-1

The default path origin is obtained from the instance node set START. If this
set is unavailable, the script selects an endpoint automatically and prints a
path note.

Use another instance:

--instance "PIPELINE-1"

Use another instance-level starting node set:

--start-node-set "PIPE_START"

Use a particular node label as distance zero. This option overrides
--start-node-set:

--start-node 1001

If the pipeline contains disconnected components, each component receives a
separate Path ID. The report identifies each path origin, and the Excel
workbook uses a separate worksheet for each path.


Section-point handling
----------------------

Default behavior preserves every section point and outputs one adjacent
LE11/THE11/MECH11 column group for each section point.

To average values through all section points before calculating MECH11, add:

--average-section-points

At a matching node and section point, contributions from connected elements
are averaged. The report and console output identify this behavior.


Frame selection
---------------

By default, the script searches backward in each selected step and uses the
last frame containing both LE and THE:

--frame-index -1

To use a particular zero-based frame index, for example frame 1:

--frame-index 1

A step is skipped if the requested frame is unavailable or does not contain
both LE and THE. The reason is printed in the console.


Output location and custom report name
--------------------------------------

Reports and Excel workbooks are written into the input ODB folder unless
--output-dir is used:

--output-dir "C:\path\to\results"

When processing exactly one ODB, --output-name can assign a custom name to
the text report:

--output-name "custom_pipeline_report.rpt"

The Excel workbook continues to use the ODB-based name:

model_MAX_MECH11_along_path.xlsx

Existing output files with the same names are overwritten.


Example console output
----------------------

Selected steps:
  Step-2 (frame index 10)
  Step-5 (frame index 18)
Wrote:      C:\Data\Results\model_LE11_THE11_MECH11.rpt
Excel plot: C:\Data\Results\model_MAX_MECH11_along_path.xlsx
Completed: 1 succeeded, 0 failed.


Important command syntax
------------------------

Correct:

--odb "model.odb"

Incorrect:

--model.odb

Use quotation marks around paths or names containing spaces.
