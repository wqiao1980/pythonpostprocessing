Abaqus Pipe, Connector, and Spring History - User Instructions
===============================================================

Purpose
-------

extract_pipe_connector_spring_history.py is a self-contained Abaqus Python
script. It extracts frame-by-frame histories for user-selected elements:

- ESF1 for pipe elements;
- CTF1 for connector elements;
- S11 for spring elements;
- E11 for spring elements; and
- CTF1 + S11 for each selected connector/spring pair sharing exactly one node.

It creates a tab-delimited text report and an Excel workbook with five native,
editable charts. Extraction is quiet by default: the script does not print
results or progress to the screen. Abaqus/CAE or Viewer does not need to be
opened.


Selection behavior
------------------

Element selection is separated by element family:

- --pipe-element-set, --pipe-element, --pipe-element-range
- --connector-element-set, --connector-element, --connector-element-range
- --spring-element-set, --spring-element, --spring-element-range

At least one element selector is required for an extraction. A family whose
selectors are omitted is not extracted; its Excel chart remains empty. Within
one family, sets, exact labels, and ranges form a union.

Element types are identified by these Abaqus type-name prefixes:

- PIPE for pipes, such as PIPE31H;
- CONN for connectors, such as CONN3D2; and
- SPRING for springs, such as SPRINGA, SPRING1, or SPRING2.

If a selected element set or range also contains other element types, those
other types are omitted and recorded in the text report. An exact label
supplied under the wrong family produces an error.


Quick example
-------------

Run from an Abaqus Command Prompt:

abaqus python extract_pipe_connector_spring_history.py --pipe-element-set "PIPE_SET" --connector-element-set "CONNECTOR_SET" --spring-element-set "SPRING_SET"

Process all ODBs in another directory and write results elsewhere:

abaqus python "C:\python_aba\takeoutSFSM\extract_pipe_connector_spring_history.py" --input-dir "C:\Data\BPTiberFL6" --output-dir "C:\Data\BPTiberFL6\02_Results" --pipe-element-set "PIPE_SET" --connector-element-set "CONNECTOR_SET" --spring-element-set "SPRING_SET"

Process one ODB:

abaqus python extract_pipe_connector_spring_history.py --input-dir "C:\Data\BPTiberFL6" --odb "model.odb" --pipe-element-set "PIPE_SET" --connector-element-set "CONNECTOR_SET" --spring-element-set "SPRING_SET"


List steps and element sets
---------------------------

List ordered steps and frame counts:

abaqus python extract_pipe_connector_spring_history.py --list-steps

List instance- and assembly-level sets containing supported elements:

abaqus python extract_pipe_connector_spring_history.py --instance "PART-1-1" --list-element-sets

The element-set listing shows separate PIPE, CONN, and SPRING counts. Listing
commands do not require element selectors.


Select steps
------------

All ODB steps and all their frames are included by default.

Select exact step names:

abaqus python extract_pipe_connector_spring_history.py --pipe-element-set "PIPE_SET" --steps "Step-2" "Operation"

Select an inclusive range using 1-based positions:

abaqus python extract_pipe_connector_spring_history.py --pipe-element-set "PIPE_SET" --step-range 3 7

Select an inclusive range using exact first and last names:

abaqus python extract_pipe_connector_spring_history.py --pipe-element-set "PIPE_SET" --step-range "Preload" "Operation"

Every frame of every selected step is output. Total Time is calculated as the
ODB step starting total time plus the frame's Step Time. Both time values are
included in the report and workbook.


Select pipe elements for ESF1
-----------------------------

One or more element sets:

--pipe-element-set "PIPE_SET_A" --pipe-element-set "PIPE_SET_B"

Exact element numbers:

--pipe-element 1001 1002 1005

Inclusive ranges:

--pipe-element-range 1001 1200 --pipe-element-range 2001 2200

These pipe selections control the ESF1 columns and ESF1 chart only.


Select connector elements for CTF1
----------------------------------

One or more element sets:

--connector-element-set "CONNECTOR_SET_A" --connector-element-set "CONNECTOR_SET_B"

Exact element numbers:

--connector-element 5001 5002

Inclusive ranges:

--connector-element-range 5001 5100

These connector selections control the CTF1 columns and CTF1 chart.


Select spring elements for S11 and E11
--------------------------------------

One or more element sets:

--spring-element-set "SPRING_SET_A" --spring-element-set "SPRING_SET_B"

Exact element numbers:

--spring-element 7001 7002

Inclusive ranges:

--spring-element-range 7001 7100

The same spring selection is used for both S11 and E11.


CTF1 + S11 shared-node pairing
------------------------------

The script compares connectivity for all selected connector and spring
elements. A sum series is created when one connector element and one spring
element have exactly one common node. Its column and legend name look like:

NODE 100: CONN 5001 CTF1 + SPRING 7001 S11

If several valid element pairs exist, each pair receives its own column and
curve. If no pair shares exactly one node, the CTF1 + S11 chart is present but
empty and this condition is recorded in the text report.

The sum is direct signed addition of the values reported in each element's
local-1 convention. The script does not reverse either sign or transform local
directions. Confirm that the connector CTF1 and spring S11 signs and units are
compatible before using the sum as a combined physical load.


How values are obtained
-----------------------

The script first looks for an exact field-output key, for example CTF1. It also
supports component output stored under a parent field:

- ESF.ESF1
- CTF.CTF1
- S.S11
- E.E11

If Abaqus provides more than one integration-point, section-point, or other
field value for the same selected element in a frame, the script averages the
finite values for that element. Blank cells mean that the requested result was
not available for that element and frame. NaN and infinite values are ignored.


Output files and workbook layout
--------------------------------

For model.odb, the default outputs are:

model_PIPE_CONNECTOR_SPRING_HISTORY.rpt
model_PIPE_CONNECTOR_SPRING_HISTORY.xlsx

The History Data worksheet contains:

- Total Time;
- Step;
- Frame;
- Step Time;
- one ESF1 column per selected pipe element;
- one CTF1 column per selected connector element;
- one S11 column per selected spring element;
- one E11 column per selected spring element; and
- one CTF1 + S11 column per shared-node element pair.

The workbook contains these native Excel scatter charts:

1. PIPE ESF1 History
2. Connector CTF1 History
3. Spring S11 History
4. Spring E11 History
5. CTF1 + S11 at Shared Nodes

Each selected element is a separate curve. Total Time is the horizontal axis.
The charts are ordinary Excel objects, so users can modify their colors, line
styles, axes, titles, legends, sizes, and chart types.

Use another output directory with --output-dir. For one --odb run, set a custom
report name with:

--output-name "custom_history.rpt"

The corresponding workbook is custom_history.xlsx. Existing files with the
same output names are overwritten.


Instance option
---------------

The default instance is PART-1-1. Use another instance with:

--instance "PIPELINE-1"

All exact labels and element sets are resolved for that instance. Both
instance-level and assembly-level element sets are supported.


Quiet screen output, diagnostic log, and verbose mode
-----------------------------------------------------

Normal extraction commands are quiet. Numeric results are written only to the
.rpt text report, and plots are written only to the .xlsx Excel workbook.
Warnings and tracebacks are recorded in:

extract_pipe_connector_spring_history.log

The Abaqus command launcher can still display its own license-manager messages;
those messages are outside this Python script.

To display optional progress, output paths, and warning summaries, add:

--verbose

Example verbose run:

abaqus python "C:\python_aba\takeoutSFSM\extract_pipe_connector_spring_history.py" --input-dir "C:\Data\BPTiberFL6" --output-dir "C:\Data\BPTiberFL6\02_Results" --pipe-element-set "PIPE_SET" --connector-element-set "CONNECTOR_SET" --spring-element-set "SPRING_SET" --verbose

Even in verbose mode, detailed exceptions and lists of available element sets
are not printed to the screen. If an element-set name is incorrect, the screen
shows only that the ODB failed and the diagnostic log location. The log records
the incorrect name, traceback, and complete list of available element sets.

The explicit --list-steps and --list-element-sets commands still print their
requested lists to the screen. They do not perform result extraction.


Terminate a running extraction
------------------------------

To stop a run from the Abaqus Command Prompt:

1. Click the command window so it has keyboard focus.
2. Press Ctrl+C.
3. If Windows asks "Terminate batch job (Y/N)?", type Y and press Enter.
4. If Ctrl+C does not respond, try Ctrl+Break and wait briefly.

Some ODB field-reading operations run inside Abaqus code and may not react to
Ctrl+C immediately. As a last resort, open Windows Task Manager, identify the
Abaqus/Python process started by this command, and select End task. Take care
not to terminate another Abaqus analysis or an unrelated Python process.

The script opens every ODB read-only, so terminating postprocessing does not
modify or damage the ODB. However, the current .rpt, .xlsx, or .log file may be
missing, incomplete, or unreadable if the run stops while that file is being
written. Rerun the script to overwrite those partial outputs. The Abaqus
license should be released after the stopped process exits, although the
license server may take a short time to show it as available.


Independence
------------

This script does not import or run any of the other postprocessing scripts in
the folder. Only the .py file is required on the Abaqus computer.


Troubleshooting
---------------

1. No values appear for one variable

   Confirm that variable was requested as field output in the analysis and is
   available for the selected element type. The text report states which exact
   or parent field names were found.

2. An exact element label has the wrong element type

   Supply that label under the correct --pipe-element, --connector-element, or
   --spring-element option. Run --list-element-sets to inspect suitable sets.

3. An element-set name is incorrect

   Open extract_pipe_connector_spring_history.log. The error entry contains the
   requested name and all available instance- and assembly-level element-set
   names. They are intentionally not printed during extraction, including when
   --verbose is used.

4. The CTF1 + S11 chart is empty

   Confirm that both connector and spring elements were selected and that their
   connectivity has exactly one common node. Pair details are written into the
   text report.

5. A sum has the unexpected sign

   Inspect the local-1 directions and sign conventions of both elements. The
   script adds reported CTF1 and S11 directly and does not reorient them.

6. A step has blank rows

   All frames of selected steps are retained for a continuous history. A frame
   can therefore be present even when one or more requested fields are absent.
