Abaqus Local-Nodal U2 Pipeline Report - User Instructions
==========================================================

Required files
--------------

Keep these two Python files together in the same folder:

1. extract_u2_local_selected_steps.py
2. extract_le11_the11_selected_steps.py

The U2 script reuses helper functions from the LE11/THE11 script.

Run the commands below from an Abaqus Command Prompt. The Abaqus/CAE window
does not need to be opened, but this script must use Abaqus Viewer in noGUI
mode because Abaqus Viewer performs the local nodal-coordinate transformation:

abaqus viewer noGUI=extract_u2_local_selected_steps.py -- [script options]

The separate -- between the script filename and the script options is
required.


Script dependencies
-------------------

The U2 script is not self-contained. It requires this helper file in the same
folder:

extract_le11_the11_selected_steps.py

The helper supplies ODB selection, step selection, pipeline-distance, and
Excel-writing utilities. Importing it does not run an LE11/THE11 extraction.

The dependency is one-way: the U2 script needs the original LE11/THE11
script, but the original LE11/THE11 script does not need the U2 script.


Meaning of local-nodal U2
-------------------------

The script applies the Abaqus NODAL result transformation. At nodes that have
a local nodal coordinate system defined by *TRANSFORM, U2 is the displacement
in local direction 2.

At nodes without *TRANSFORM, Abaqus leaves the result in global coordinates,
so the reported U2 is global U2 for those nodes.

U2 and path distance use the length units of the Abaqus model.


Results produced by the script
------------------------------

For every processed ODB, the script creates two files:

1. model_U2_LOCAL.rpt
2. model_U2_LOCAL_along_path.xlsx

The text report contains one table for every selected step. The path is
constructed only from Abaqus element types beginning with PIPE, such as
PIPE31 and PIPE32. Non-PIPE elements are excluded. Each row is a node
belonging to the resulting PIPE path and includes:

- Path ID
- Path Distance
- Node Label
- X, Y, and Z coordinates
- U2 LOCAL

The Excel workbook contains path distance, node label, and one local U2
column for every selected step. A native Excel scatter chart plots all
selected steps along the pipeline path. Each disconnected pipeline path has
its own worksheet and chart.

The Excel chart is editable. In Excel, users can change its chart type,
colors, line styles, markers, titles, axes, legend, size, and layout. The
chart is linked to the visible worksheet data, so changing a value updates
the chart. Excel does not need to be installed or open when Abaqus creates
the workbook.


Process all ODBs and all steps
------------------------------

The following command processes every ODB in the current folder and every
analysis step containing U2 field output:

abaqus viewer noGUI=extract_u2_local_selected_steps.py

To process every ODB in another folder:

abaqus viewer noGUI="C:\python_aba\takeoutSFSM\extract_u2_local_selected_steps.py" -- --input-dir "C:\Data\P1-2\00_NewFA\12inTRF"


List the available steps
------------------------

This command lists the ordered steps in every ODB in the current folder. It
does not create reports or Excel workbooks.

abaqus viewer noGUI=extract_u2_local_selected_steps.py -- --list-steps

To list the steps in ODBs located in another folder:

abaqus viewer noGUI="C:\python_aba\takeoutSFSM\extract_u2_local_selected_steps.py" -- --input-dir "C:\Data\P1-2\00_NewFA\12inTRF" --list-steps


Output selected steps by name
-----------------------------

List each exact step name after --steps. Put quotation marks around step
names that contain spaces.

abaqus viewer noGUI=extract_u2_local_selected_steps.py -- --steps "Step-2" "Step-5" "Step-8"

Step-name matching is not case-sensitive. A requested step that is not in an
ODB is omitted, and the script prints a note.


Output a range using step positions
-----------------------------------

Step positions are 1-based and the range is inclusive. The following example
outputs steps 3, 4, 5, 6, and 7:

abaqus viewer noGUI=extract_u2_local_selected_steps.py -- --step-range 3 7

If an ODB has fewer steps than the requested ending position, the script uses
the last available step and prints a note.


Output a range using step names
-------------------------------

The first and last named steps are both included:

abaqus viewer noGUI=extract_u2_local_selected_steps.py -- --step-range "Preload" "Operation"


Process selected steps in the 12inTRF folder
--------------------------------------------

Example: output steps 3 through 7 from every ODB in 12inTRF:

abaqus viewer noGUI="C:\python_aba\takeoutSFSM\extract_u2_local_selected_steps.py" -- --input-dir "C:\Data\P1-2\00_NewFA\12inTRF" --step-range 3 7

Example: output selected named steps from every ODB in 12inTRF:

abaqus viewer noGUI="C:\python_aba\takeoutSFSM\extract_u2_local_selected_steps.py" -- --input-dir "C:\Data\P1-2\00_NewFA\12inTRF" --steps "Step-2" "Step-5"


Process one ODB only
--------------------

Add --odb followed by the ODB filename:

abaqus viewer noGUI=extract_u2_local_selected_steps.py -- --odb "model.odb" --step-range 3 7

If the ODB is in another folder, combine --input-dir and --odb:

abaqus viewer noGUI="C:\python_aba\takeoutSFSM\extract_u2_local_selected_steps.py" -- --input-dir "C:\Data\P1-2\00_NewFA\12inTRF" --odb "model.odb" --steps "Step-2" "Step-5"

The --odb option may be repeated to process several named ODB files.


Path distance and starting node
-------------------------------

Path distance is the cumulative distance along the undeformed PIPE-element
mesh. It is calculated from the nodal coordinates in the ODB. Elements whose
type does not begin with PIPE are not used to construct the path.

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


Frame selection
---------------

By default, the script searches backward in each selected step and uses the
last frame containing U2:

--frame-index -1

To use a particular zero-based frame index, for example frame 1:

--frame-index 1

A step is skipped if the requested frame is unavailable or does not contain
U2. The reason is printed in the console.


Output location and custom report name
--------------------------------------

Reports and Excel workbooks are written into the input ODB folder unless
--output-dir is used:

--output-dir "C:\path\to\results"

When processing exactly one ODB, --output-name can assign a custom name to
the text report:

--output-name "custom_U2_report.rpt"

The Excel workbook continues to use the ODB-based name:

model_U2_LOCAL_along_path.xlsx

Existing output files with the same names are overwritten.


Example console output
----------------------

Selected steps:
  Step-2 (frame index 10)
  Step-5 (frame index 18)
Wrote:      C:\Data\Results\model_U2_LOCAL.rpt
Excel plot: C:\Data\Results\model_U2_LOCAL_along_path.xlsx
Completed: 1 succeeded, 0 failed.


Important command syntax
------------------------

Correct:

abaqus viewer noGUI=extract_u2_local_selected_steps.py -- --odb "model.odb"

Incorrect:

abaqus viewer noGUI=extract_u2_local_selected_steps.py --model.odb

The first standalone -- separates Abaqus launcher arguments from the Python
script arguments. The second --odb is the actual script option.

Use quotation marks around paths or names containing spaces.


Debug logs and console capture
------------------------------

The script automatically creates this diagnostic log in --output-dir on
every processing run:

extract_u2_local_selected_steps.log

For example, when the output directory is:

C:\Data\BPTiberFL6\02_Results

the automatic log is:

C:\Data\BPTiberFL6\02_Results\extract_u2_local_selected_steps.log

The log records the script version, complete sys.argv received from Viewer,
input and output directories, number of ODB jobs, and a full traceback for
every failed ODB.

To also capture all Abaqus launcher and Python console output, add normal
Windows Command Prompt redirection to the end of the command:

> "C:\Data\BPTiberFL6\02_Results\u2_console.txt" 2>&1

Complete example:

abaqus viewer noGUI="C:\Data\BPTiberFL6\00_TestCode\extract_u2_local_selected_steps.py" -- --input-dir "C:\Data\BPTiberFL6" --output-dir "C:\Data\BPTiberFL6\02_Results" > "C:\Data\BPTiberFL6\02_Results\u2_console.txt" 2>&1

The > operator overwrites u2_console.txt on each run. Use >> instead of > to
append new console output to the existing file.

When reporting an error, provide both files:

1. extract_u2_local_selected_steps.log
2. u2_console.txt


Troubleshooting
---------------

1. ImportError for extract_le11_the11_selected_steps

   Keep extract_le11_the11_selected_steps.py in the same folder as
   extract_u2_local_selected_steps.py.

2. No parsable local U2 data

   Confirm that displacement field output U includes component U2 for the
   requested frame and that the pipeline instance contains the required
   nodes. Run the script in a fresh Abaqus Viewer noGUI session.

3. U2 appears to be in global coordinates

   Confirm that the pipeline nodes have a nodal coordinate transformation
   defined in the model with *TRANSFORM. Nodes without *TRANSFORM use global
   directions.

4. Abaqus Viewer license error

   The script requires an available Abaqus Viewer license because Viewer
   performs the NODAL transformation.

5. NameError: name '__file__' is not defined

   This occurred with an earlier version of the U2 script under some Abaqus
   Viewer releases. Use the current extract_u2_local_selected_steps.py, which
   determines its location from the Abaqus noGUI execution information when
   __file__ is unavailable.

6. Viewer exits with code 2 and no usage message

   This occurred with an earlier version when Viewer placed the first user
   option at sys.argv[0]. Use the current script, which accepts both known
   Abaqus Viewer argument layouts. The current console output begins with:

   extract_u2_local_selected_steps.py version 2026-09-05-r5

   It then prints the script arguments received from Viewer.

7. Viewer reports exit code 1 but hides the Python traceback

   The current script writes complete diagnostics directly to:

   extract_u2_local_selected_steps.log

   This file is placed in --output-dir. Open or share this log to identify the
   exact ODB and Abaqus API call that failed.

8. Session has no attribute defaultOdbDisplay

   This occurred with an earlier version in Abaqus Viewer 2022 noGUI mode.
   The current script configures the ODB display on a real viewport and
   creates a viewport automatically if Viewer did not create one.

9. Non-pipe elements appear in the pipeline instance

   Version 2026-09-05-r5 and later constructs the U2 path only from element
   types beginning with PIPE. Nodes that belong only to beams, connectors,
   solids, or other non-PIPE elements are excluded. A node shared by a PIPE
   element and another element remains included because it belongs to the
   PIPE path.
