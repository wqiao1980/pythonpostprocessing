Abaqus Pipe S11 Delta Between Step Pairs - User Instructions
=============================================================

Purpose
-------

extract_pipe_s11_delta_step_pairs.py is a self-contained Abaqus Python script.
For every user-defined pair of steps, it:

- reads S11 for pipe elements along the pipeline path;
- reads every frame and every available section point;
- finds both the maximum and minimum S11 over all frames independently in each
  step at every element/node/section-point location;
- calculates MAX Delta S11 = maximum S11 in the first step minus maximum S11
  in the second step, and MIN Delta S11 = minimum S11 in the first step minus
  minimum S11 in the second step, at each matching location;
- retains the greatest signed MAX Delta S11 and the smallest signed MIN Delta
  S11 separately at -90, 0, 90, and 180 degrees for the inner, middle, and
  outer radii at each path node;
- writes a wide tab-delimited .rpt report containing one pipeline-distance
  column followed by twenty-four MAX/MIN radius/angle columns for every
  requested step pair; and
- writes the same final data into three editable Excel .xlsx workbooks, one
  each for the inner, middle, and outer fiber, with one native editable chart
  for every requested step pair.

Raw frame-by-frame S11 values are not written unless --write-intermediate is
entered. The per-location step envelopes used for Delta S11 are always handled
in memory and are not included in the intermediate report. The three final
Excel workbooks are always created; intermediate values remain in .rpt format
only when requested.

Abaqus/CAE or Viewer does not need to be opened. Run the script with
"abaqus python" from an Abaqus Command Prompt.


Important calculation definitions
---------------------------------

Step-pair order matters. At each matching element/node/section-point location:

MAX Delta S11 = maximum S11 over all first-step frames
                  - maximum S11 over all second-step frames

MIN Delta S11 = minimum S11 over all first-step frames
                  - minimum S11 over all second-step frames

For example:

--step-pair "AsLaid" "Operational"

means:

MAX Delta S11 = max-frame S11(AsLaid) - max-frame S11(Operational)
MIN Delta S11 = min-frame S11(AsLaid) - min-frame S11(Operational)

Frames are not paired. Each step is scanned and enveloped independently. The
two steps may have different frame counts, frame times, and time increments.

For example, at one matching section point:

- first-step S11 across its frames: 10, 25, 20; maximum = 25 and minimum = 10;
- second-step S11 across its frames: 5, 30; maximum = 30 and minimum = 5;
- MAX Delta S11 = 25 - 30 = -5; and
- MIN Delta S11 = 10 - 5 = 5.

At one path node, the script reports twenty-four values for each step pair.
For every radius/angle section point, it reports the greatest signed MAX Delta
S11 and the smallest signed MIN Delta S11 among all selected pipe elements
contributing at that node:

- inner radius at -90, 0, 90, and 180 degrees;
- middle radius at -90, 0, 90, and 180 degrees; and
- outer radius at -90, 0, 90, and 180 degrees.

Positive and negative radius section points are not combined.

These are numerical signed envelopes, not absolute magnitudes. For example,
the reported MAX of -20, -5, and 3 is 3, while the reported MIN is -20.


Quick examples
--------------

Process all ODB files in the current folder using step names:

abaqus python extract_pipe_s11_delta_step_pairs.py --step-pair "AsLaid" "Operational"

Process all ODBs in another folder and write results elsewhere:

abaqus python "C:\python_aba\takeoutSFSM\extract_pipe_s11_delta_step_pairs.py" --input-dir "C:\Data\BPTiberFL6" --output-dir "C:\Data\BPTiberFL6\02_Results" --step-pair "AsLaid" "Operational"

Process one ODB only:

abaqus python "C:\python_aba\takeoutSFSM\extract_pipe_s11_delta_step_pairs.py" --input-dir "C:\Data\BPTiberFL6" --odb "model.odb" --output-dir "C:\Data\BPTiberFL6\02_Results" --step-pair "AsLaid" "Operational"


Select step pairs
-----------------

At least one --step-pair FIRST SECOND is required. Repeat the option to request
more than one pair.

Use exact step names:

abaqus python extract_pipe_s11_delta_step_pairs.py --step-pair "AsLaid" "Hydrotest" --step-pair "Hydrotest" "Operation"

Use 1-based step positions instead of whole names:

abaqus python extract_pipe_s11_delta_step_pairs.py --step-pair 3 7 --step-pair 7 10

Names and positions may be mixed:

abaqus python extract_pipe_s11_delta_step_pairs.py --step-pair 3 "Operation"

Step positions are based on the ODB step order and are 1-based. Step position 1
means the first ODB step, not frame 1. Run --list-steps first if the positions
are not known.

List ordered steps and frame counts:

abaqus python extract_pipe_s11_delta_step_pairs.py --input-dir "C:\Data\BPTiberFL6" --odb "model.odb" --list-steps

Listing does not require --step-pair and does not extract results.


Pipe element selection
----------------------

If no pipe element selector is entered, all supported pipe elements on the
resolved START-to-END path are processed.

Restrict the calculation to one or more element sets:

--pipe-element-set "PIPE_SET_A" --pipe-element-set "PIPE_SET_B"

Restrict it to exact element labels:

--pipe-element 1001 1002 1005

Restrict it to inclusive label ranges:

--pipe-element-range 1001 1200 --pipe-element-range 2001 2200

Sets, exact labels, and ranges may be combined; their pipe elements form a
union. Selected pipe elements outside the resolved START-to-END path are
omitted and recorded in the log.

The script recognizes element types whose Abaqus type name begins with PIPE,
including PIPE31H. Two-node and three-node pipe connectivity are supported.

List only element sets that contain pipe elements:

abaqus python extract_pipe_s11_delta_step_pairs.py --input-dir "C:\Data\BPTiberFL6" --odb "model.odb" --instance "PART-1-1" --list-pipe-element-sets


Pipeline path and distance
--------------------------

The default pipeline instance is PART-1-1. Use another instance with:

--instance "PIPELINE-1"

The script builds a connected graph from supported pipe elements. It determines
the path start and end in this order:

1. Explicit --start-node and --end-node labels, if supplied.
2. Explicit --start-node-set and --end-node-set names, if supplied.
3. A unique instance node-set name containing START and a unique instance
   node-set name containing END. The names do not need to equal START or END.
4. If no endpoint sets are found, the two endpoints of a single unbranched pipe
   graph.

Examples:

--start-node 1001 --end-node 2500

--start-node-set "PIPE_START_NODE" --end-node-set "PIPE_END_NODE"

List node sets whose names contain START or END:

abaqus python extract_pipe_s11_delta_step_pairs.py --input-dir "C:\Data\BPTiberFL6" --odb "model.odb" --instance "PART-1-1" --list-endpoint-sets

Pipeline distance is the cumulative three-dimensional distance between the
original ODB node coordinates along the resolved pipe route. The start node has
distance zero.


Default final report
--------------------

The final report is always written because it is the requested result. For:

model.odb

the report is:

model_MAX_DELTA_S11_PATH.rpt

It contains path and section-point mapping metadata followed by a wide table.
For one step pair, the column pattern is:

Pipeline Distance
DELTA_S11 MAX INNER FIBER ANGLE -90 DEG [AsLaid - Hydrotest]
DELTA_S11 MIN INNER FIBER ANGLE -90 DEG [AsLaid - Hydrotest]
DELTA_S11 MAX INNER FIBER ANGLE 0 DEG [AsLaid - Hydrotest]
DELTA_S11 MIN INNER FIBER ANGLE 0 DEG [AsLaid - Hydrotest]
DELTA_S11 MAX INNER FIBER ANGLE 90 DEG [AsLaid - Hydrotest]
DELTA_S11 MIN INNER FIBER ANGLE 90 DEG [AsLaid - Hydrotest]
DELTA_S11 MAX INNER FIBER ANGLE 180 DEG [AsLaid - Hydrotest]
DELTA_S11 MIN INNER FIBER ANGLE 180 DEG [AsLaid - Hydrotest]
DELTA_S11 MAX MIDDLE FIBER ANGLE -90 DEG [AsLaid - Hydrotest]
DELTA_S11 MIN MIDDLE FIBER ANGLE -90 DEG [AsLaid - Hydrotest]
DELTA_S11 MAX MIDDLE FIBER ANGLE 0 DEG [AsLaid - Hydrotest]
DELTA_S11 MIN MIDDLE FIBER ANGLE 0 DEG [AsLaid - Hydrotest]
DELTA_S11 MAX MIDDLE FIBER ANGLE 90 DEG [AsLaid - Hydrotest]
DELTA_S11 MIN MIDDLE FIBER ANGLE 90 DEG [AsLaid - Hydrotest]
DELTA_S11 MAX MIDDLE FIBER ANGLE 180 DEG [AsLaid - Hydrotest]
DELTA_S11 MIN MIDDLE FIBER ANGLE 180 DEG [AsLaid - Hydrotest]
DELTA_S11 MAX OUTER FIBER ANGLE -90 DEG [AsLaid - Hydrotest]
DELTA_S11 MIN OUTER FIBER ANGLE -90 DEG [AsLaid - Hydrotest]
DELTA_S11 MAX OUTER FIBER ANGLE 0 DEG [AsLaid - Hydrotest]
DELTA_S11 MIN OUTER FIBER ANGLE 0 DEG [AsLaid - Hydrotest]
DELTA_S11 MAX OUTER FIBER ANGLE 90 DEG [AsLaid - Hydrotest]
DELTA_S11 MIN OUTER FIBER ANGLE 90 DEG [AsLaid - Hydrotest]
DELTA_S11 MAX OUTER FIBER ANGLE 180 DEG [AsLaid - Hydrotest]
DELTA_S11 MIN OUTER FIBER ANGLE 180 DEG [AsLaid - Hydrotest]

These headings appear on one tab-delimited header row. They are shown one per
line above only for readability.

There is only one location column: Pipeline Distance. A node-label column is
not written. Each --step-pair adds exactly twenty-four adjacent result columns.
MAX and MIN are adjacent for each angle. The order is INNER at all four angles,
MIDDLE at all four angles, then OUTER at all four angles. Additional pairs
repeat this twenty-four-column group to the right in command-line order. Blank
cells mean no matching finite S11 value was available for that radius, angle,
path location, and pair.

The .rpt output is tab-delimited even though the example above uses spaces for
readability. It can be opened in a text editor or imported into Excel manually.


Three final Excel workbooks
---------------------------

The script also writes three separate editable Excel workbooks for every ODB:

model_MAX_DELTA_S11_INNER_FIBER.xlsx
model_MAX_DELTA_S11_MIDDLE_FIBER.xlsx
model_MAX_DELTA_S11_OUTER_FIBER.xlsx

Each workbook contains only one fiber. Its first data column is Pipeline
Distance. For every requested step pair, the next eight columns contain MAX
then MIN Delta S11 at -90, 0, 90, and 180 degrees. Additional step pairs repeat
this eight-column group to the right in the same order as the command line.

Each workbook contains one native editable XY scatter/line chart for every
requested step pair. A chart plots all eight series against the numeric
Pipeline Distance values: four solid MAX lines and four dashed MIN lines. The
same color identifies the same angle in both envelopes. Charts are placed below
the data table so they do not cover results. Users can change chart type, line
style, colors, axes, legend, title, and other formatting directly in Excel.

The cells are numeric and can be reformatted, filtered, plotted, or used in
Excel formulas. The title rows record the ODB name, calculation definition,
and section-point mapping. The header row is frozen together with the Pipeline
Distance column. Blank cells have the same meaning as blank cells in the final
.rpt report.

The three .xlsx files and the wide final .rpt file are written to --output-dir,
or to --input-dir when --output-dir is omitted. Existing files with the same
names are overwritten on a new run.


Optional intermediate report
----------------------------

By default, the script does not write raw S11 to a file. It scans one step at a
time and retains only the maximum and minimum S11 at each element/node/section-
point location and the signed MAX/MIN Delta S11 envelopes needed for the final
report.

To request all intermediate values, add:

--write-intermediate

Example:

abaqus python extract_pipe_s11_delta_step_pairs.py --step-pair "AsLaid" "Operational" --write-intermediate

For each ODB and each requested pair, this writes a file whose name looks like:

model_Pair001_AsLaid_TO_Operational_S11_INTERMEDIATE.rpt

The intermediate report contains one wide tab-delimited table. Its first five
columns are always:

Pipe Distance
Node Number(s)
Element Number
Step Number
Frame Number

Every remaining column contains S11 for one section point. Each heading gives
the output thick-pipe angle, signed radius, and Abaqus section-point number. For
example:

S11 ANGLE -90 DEG RADIUS +0.67771 [SP 1]
S11 ANGLE -90 DEG RADIUS +0.838855 [SP 2]
S11 ANGLE -90 DEG RADIUS +1 [SP 3]
S11 ANGLE 0 DEG RADIUS +0.67771 [SP 13]

Columns are ordered first by angle, then from the smallest absolute radius to
the largest absolute radius. Positive and negative radius values remain
separate. Angles outside -90, 0, 90, and 180 degrees are also retained in the
intermediate table and follow those four standard angles.

Step Number is the 1-based position of the step in the ODB. The report header
maps each number to its complete step name. Frame Number is the zero-based
Abaqus frame index. Thus frame 0 is the first frame of that step.

Each data row represents one start-to-end path element for one step and frame.
Pipe Distance is the midpoint station of that element, calculated as one half
of the sum of the route distances at its first and last path nodes. Node
Number(s) lists the nodes belonging to that element; for example, 101,102 means
that the element connects nodes 101 and 102. Element Number contains exactly
one pipe element label.

For every section point, the intermediate S11 is the arithmetic average of the
available element-nodal S11 values belonging to that same element. S11 is never
averaged with an adjacent pipe element. A blank S11 cell means that the section-
point value was not available for that element and frame.

Rows are ordered as follows:

1. first selected step, frame 0, element midpoint distance from START to END;
2. first selected step, frame 1, element midpoint distance from START to END;
3. remaining first-step frames in increasing frame order;
4. second selected step, frame 0, pipe distance from START to END; and
5. remaining second-step frames in increasing frame order.

Thus the ordered element-distance sequence restarts for each new frame. Because
distances are located at element midpoints, the first value is normally half
the first element length rather than zero, and the last value is normally half
the last element length before the total route length. When pipe element
selectors are used, only selected elements on the START-to-END route appear.

The intermediate report does not pair frames and does not contain a per-frame
Delta S11. The final calculation still independently creates maximum and
minimum envelopes from all frames in each step, then subtracts each matching
second-step envelope from its first-step envelope.

The intermediate report can be very large because it includes every frame,
section point, contributing element, and path node. It is tab-delimited and is
not an Abaqus/CAE XY report. Intermediate data is not copied into the three
final Excel workbooks.


How S11 is obtained and matched
-------------------------------

The script first looks for an exact S11 field-output key. It also supports S11
stored as the S11 component of the parent S field.

S11 is requested at ELEMENT_NODAL position so every result is associated with
a pipe path node. Abaqus may extrapolate integration-point results when it
creates this read-only subset. The ODB itself is not modified.

Each step is first enveloped independently over all of its frames to obtain
both a maximum and a minimum at every location. Locations from the two step
envelopes are then compared only when node label, element label, and section
point match. Section points are matched by section-point number when that
number is available; their descriptions are used only when no number is
available. A controlling frame in the first step does not have to equal the
corresponding controlling frame in the second step. Duplicate finite values at
an identical location within one frame are averaged. NaN and infinite values
are ignored.

For the optional intermediate wide table only, the element-nodal S11 values at
the nodes of one element are averaged to give that element's section-point S11.
Adjacent elements remain separate rows at separate element-midpoint distances
and are never averaged together. Both final Delta S11 calculations remain
element/node/section-point specific before their documented final node
envelopes, so this intermediate presentation change does not alter the final
.rpt or Excel results.


Inner, middle, and outer fiber identification
---------------------------------------------

The final report always contains twenty-four MAX/MIN radius/angle columns per
step pair. The script reads both the output thick-pipe section radius and
output thick-pipe section angle from each Abaqus section-point description.

Some Abaqus ODBs write an explicit radius, for example:

Angle = -67.5000, Radius = 0.6777

At the principal angles -90, 0, 90, and 180 degrees, Abaqus can instead write
two section-coordinate fractions, for example:

Angle = -90.0000, (1-fraction = 0.000000, 2-fraction = -0.677710)

When Radius is absent, the script calculates its nonnegative magnitude as:

radius = sqrt((1-fraction)^2 + (2-fraction)^2)

Thus the example above maps to radius 0.677710. The angle retains the
circumferential direction. Explicit signed Radius values remain signed and are
not combined with a different signed-radius section point at the same angle.

The required mapping is:

- smallest absolute radius = inner fiber;
- middle absolute radius = middle fiber; and
- absolute radius 1.0 = outer fiber.

The script requires exactly three distinct absolute radius magnitudes and
requires the largest magnitude to equal 1.0 within numerical tolerance. It then
keeps a separate section-point group for each radius at these four angles:

- -90 degrees;
- 0 degrees;
- 90 degrees; and
- 180 degrees.

Equivalent angle forms are normalized; for example, 270 degrees is treated as
-90 degrees and -180 degrees is treated as 180 degrees. Other angles are
omitted from the final path table and recorded in the report mapping metadata.

Positive and negative radius points are never put into the same result column.
If more than one signed-radius section point maps to the same fiber and angle,
the script stops for that ODB instead of combining them. All twelve required
radius/angle combinations must be present.

If a radius or angle cannot be read, or the required combinations are not
present, the ODB fails instead of silently guessing. The diagnostic log lists
the detected section-point descriptions, radii, angles, missing combinations,
or duplicates. The selected radius/angle mapping and section-point labels are
also written in the header of the final report.


Quiet operation and diagnostic log
----------------------------------

Normal extraction does not print numeric results or progress to the screen.
The Abaqus command launcher can still display its own license-manager messages.

Each run writes this diagnostic log in the output directory:

extract_pipe_s11_delta_step_pairs.log

The log records selected paths, step pairs, frame counts, matching statistics,
S11 source fields, both MAX and MIN calculation definitions, the final .rpt
path, all three Excel paths, warnings, errors, and tracebacks.

To print live frame-by-frame progress and output paths, add:

--verbose

Example:

abaqus python extract_pipe_s11_delta_step_pairs.py --odb "model.odb" --step-pair 22 23 --verbose

The script prints a completion line after every processed frame, for example:

Pair 1/1, first step 'Step-22': frame 15/86 complete (17.4%); finite samples=24576, cumulative=368640
Pair 1/1, second step 'Step-23': frame 21/103 complete (20.4%); finite samples=24576, cumulative=516096

Each progress line identifies the step pair, first or second step, completed
frame count, percentage, finite samples in that frame, and cumulative samples
for that step. Output is flushed immediately, so the lines also appear promptly
when the command output is redirected to a console text file. A single large
frame can still take time before its completion line appears.

Even in verbose mode, detailed errors and complete lists of available element
sets are written to the log instead of being printed. Explicit listing commands
such as --list-steps and --list-pipe-element-sets intentionally print their
requested lists and then exit.


Terminate a running extraction
------------------------------

To stop a run from the Abaqus Command Prompt:

1. Click the command window so it has keyboard focus.
2. Press Ctrl+C.
3. If Windows asks "Terminate batch job (Y/N)?", type Y and press Enter.
4. If Ctrl+C does not respond, try Ctrl+Break and wait briefly.

Some ODB field-reading operations run inside Abaqus code and may not respond to
Ctrl+C immediately. As a last resort, open Windows Task Manager, identify the
Abaqus/Python process started by this command, and select End task. Do not stop
another Abaqus analysis or unrelated Python process.

The ODB is opened read-only, so terminating postprocessing does not modify it.
A final .rpt, .xlsx, intermediate, or log file may be incomplete if the program
is stopped while writing that file. Rerun the command to overwrite partial
outputs. A temporary file named s11_intermediate_*.tmp can remain in the output
directory after a forced termination; it may be deleted after confirming that
the Abaqus postprocessing process has stopped.


Independence
------------

This script does not import or run any other postprocessing script in the
folder. Only extract_pipe_s11_delta_step_pairs.py is required on the Abaqus
computer.


Troubleshooting
---------------

1. The run fails because a step was not found

   Run --list-steps on the same ODB. Use an exact displayed name or its 1-based
   position. A step pair must resolve to two different steps.

2. An element-set name is wrong

   Open extract_pipe_s11_delta_step_pairs.log. It records the requested name,
   traceback, and available instance- and assembly-level element-set names.
   Use --list-pipe-element-sets if a screen listing is intentionally wanted.

3. Final cells are blank

   Confirm S11 was requested as field output for the relevant steps and pipe
   elements. Also confirm that the independently calculated step envelopes
   contain matching node, element, and section-point locations. Matching counts
   and source fields are recorded in the log.

4. The steps have different numbers of frames

   This is allowed. All frames in each step are used to calculate that step's
   S11 maximum and minimum. Frames are not paired and the two frame counts do
   not have to match. The log records both frame and sample counts.

5. The selected elements produce no output

   Confirm that the selection contains PIPE elements on the resolved START-to-
   END route. The log records selected elements that were outside the route.

6. The intermediate report is too large or the run is slow

   Omit --write-intermediate for the normal fast mode. Restrict the run with
   --odb and, if appropriate, pipe element sets, labels, or label ranges.

7. Excel reports that a workbook is damaged

   Delete the partial .xlsx file and rerun the command. This can occur if the
   Abaqus process was terminated while the workbook was being written. The
   diagnostic log identifies the intended Excel output paths.
