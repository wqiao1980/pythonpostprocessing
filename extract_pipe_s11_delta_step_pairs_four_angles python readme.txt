Abaqus Pipe S11 Heat-Up/Cool-Down Delta Four-Angle Fast Variant
===============================================================

Purpose
-------

extract_pipe_s11_delta_step_pairs_four_angles.py is a self-contained Abaqus Python script.
For every user-defined pair of steps, it:

- reads S11 for pipe elements along the pipeline path;
- treats the first step in every pair as HEAT-UP and reads all of its frames,
  retaining S11 only at -90, 0, 90, and 180 degrees;
- treats the second step as COOL-DOWN and uses only its last frame in the
  final Delta S11 calculation;
- calculates MAX Delta S11 = maximum S11 over all heat-up frames minus S11 at
  the last cool-down frame, and MIN Delta S11 = minimum S11 over all heat-up
  frames minus S11 at the same last cool-down frame;
- retains the greatest signed MAX Delta S11 and the smallest signed MIN Delta
  S11 separately at -90, 0, 90, and 180 degrees for the inner, middle, and
  outer radii at each path node;
- writes a wide tab-delimited .rpt report containing one pipeline-distance
  column followed by twenty-four MAX/MIN radius/angle columns for every
  requested step pair; and
- writes the same final data into three editable Excel .xlsx workbooks, one
  each for the inner, middle, and outer fiber, with one native editable chart
  for every requested step pair.

Raw frame-by-frame S11 values are never written. When --write-intermediate is
entered, the script writes only the heat-up maximum envelope, heat-up minimum
envelope, and last-frame cool-down S11 at -90, 0, 90, and 180 degrees. The
three final Excel workbooks are always created; intermediate envelope values
remain in .rpt format only when requested.

This is a separate optimized variant. It does not import, change, or overwrite
extract_pipe_s11_delta_step_pairs.py. Its Python script, log, final report,
Excel workbooks, and optional intermediate reports all have distinct names.

Abaqus/CAE or Viewer does not need to be opened. Run the script with
"abaqus python" from an Abaqus Command Prompt.


Why this variant is faster
--------------------------

The original script retains every available pipe section point while scanning
a frame and selects the four requested angles later. This variant checks the
section-point angle first and retrieves, stores, sorts, and envelopes numeric
S11 only when the angle is equivalent to -90, 0, 90, or 180 degrees. The angle
decision is cached by section-point identity within each frame.

For a thick pipe with sixteen angles and three radii, this reduces the retained
locations from 48 to 12 section points per element/node. For example:

15,955 elements x 2 nodes x 48 section points = 1,531,680 original samples
15,955 elements x 2 nodes x 12 section points =   382,920 four-angle samples

That is a 75 percent reduction in numeric samples retained and enveloped. The
Abaqus field-output collection must still be traversed to identify each value's
section point, so the actual elapsed-time improvement depends on ODB storage,
disk speed, element count, and frame count. Both scripts scan only the final
cool-down frame. This variant also skips the per-frame location sort.


Important calculation definitions
---------------------------------

Step-pair order matters. Always enter HEAT-UP first and COOL-DOWN second. At
each matching element/node/section-point location:

MAX Delta S11 = maximum S11 over all HEAT-UP frames
                  - S11 at the last COOL-DOWN frame

MIN Delta S11 = minimum S11 over all HEAT-UP frames
                  - S11 at the last COOL-DOWN frame

For example:

--step-pair "HeatUp" "CoolDown"

means:

MAX Delta S11 = max(S11 over every HeatUp frame) - S11(last CoolDown frame)
MIN Delta S11 = min(S11 over every HeatUp frame) - S11(last CoolDown frame)

Frames are not paired. The heat-up step is enveloped over all frames. The
cool-down step may have any number of frames, but only its last frame is used.

For example, at one matching section point:

- heat-up S11 across its frames: 10, 25, 20; maximum = 25 and minimum = 10;
- cool-down S11 across its frames: 5, 30; last-frame S11 = 30;
- MAX Delta S11 = 25 - 30 = -5; and
- MIN Delta S11 = 10 - 30 = -20.

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

abaqus python extract_pipe_s11_delta_step_pairs_four_angles.py --step-pair "HeatUp" "CoolDown"

Process all ODBs in another folder and write results elsewhere:

abaqus python "C:\python_aba\takeoutSFSM\extract_pipe_s11_delta_step_pairs_four_angles.py" --input-dir "C:\Data\BPTiberFL6" --output-dir "C:\Data\BPTiberFL6\02_Results" --step-pair "HeatUp" "CoolDown"

Process one ODB only:

abaqus python "C:\python_aba\takeoutSFSM\extract_pipe_s11_delta_step_pairs_four_angles.py" --input-dir "C:\Data\BPTiberFL6" --odb "model.odb" --output-dir "C:\Data\BPTiberFL6\02_Results" --step-pair "HeatUp" "CoolDown"

Process several pairs using one compact option:

abaqus python extract_pipe_s11_delta_step_pairs_four_angles.py --step-pairs 10 11 14 15 22 23


Select step pairs
-----------------

At least one pair is required. In every pair, the first entry must be the
heat-up step and the second entry must be the cool-down step.

For one pair, or for the original repeatable syntax, use --step-pair:

Use exact step names:

abaqus python extract_pipe_s11_delta_step_pairs_four_angles.py --step-pair "HeatUp-1" "CoolDown-1" --step-pair "HeatUp-2" "CoolDown-2"

For a long list, use --step-pairs once and enter consecutive references:

abaqus python extract_pipe_s11_delta_step_pairs_four_angles.py --step-pairs 10 11 14 15 22 23

This is interpreted as the three pairs 10->11, 14->15, and 22->23. You do not
repeat --step-pairs before each pair. An even number of references is required.
Exact step names work the same way; quote each name if it contains spaces:

abaqus python extract_pipe_s11_delta_step_pairs_four_angles.py --step-pairs "HeatUp-1" "CoolDown-1" "HeatUp-2" "CoolDown-2"

Use 1-based step positions instead of whole names:

abaqus python extract_pipe_s11_delta_step_pairs_four_angles.py --step-pair 3 7 --step-pair 7 10

Names and positions may be mixed:

abaqus python extract_pipe_s11_delta_step_pairs_four_angles.py --step-pair 3 "Operation"

Step positions are based on the ODB step order and are 1-based. Step position 1
means the first ODB step, not frame 1. Run --list-steps first if the positions
are not known.

List ordered steps and frame counts:

abaqus python extract_pipe_s11_delta_step_pairs_four_angles.py --input-dir "C:\Data\BPTiberFL6" --odb "model.odb" --list-steps

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

abaqus python extract_pipe_s11_delta_step_pairs_four_angles.py --input-dir "C:\Data\BPTiberFL6" --odb "model.odb" --instance "PART-1-1" --list-pipe-element-sets


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

abaqus python extract_pipe_s11_delta_step_pairs_four_angles.py --input-dir "C:\Data\BPTiberFL6" --odb "model.odb" --instance "PART-1-1" --list-endpoint-sets

Pipeline distance is the cumulative three-dimensional distance between the
original ODB node coordinates along the resolved pipe route. The start node has
distance zero.


Default final report
--------------------

The final report is always written because it is the requested result. For:

model.odb

the report is:

model_MAX_DELTA_S11_4ANGLES_PATH.rpt

It contains path and section-point mapping metadata followed by a wide table.
For one step pair, the column pattern is:

Pipeline Distance
DELTA_S11 MAX INNER FIBER ANGLE -90 DEG [HeatUp - CoolDown]
DELTA_S11 MIN INNER FIBER ANGLE -90 DEG [HeatUp - CoolDown]
DELTA_S11 MAX INNER FIBER ANGLE 0 DEG [HeatUp - CoolDown]
DELTA_S11 MIN INNER FIBER ANGLE 0 DEG [HeatUp - CoolDown]
DELTA_S11 MAX INNER FIBER ANGLE 90 DEG [HeatUp - CoolDown]
DELTA_S11 MIN INNER FIBER ANGLE 90 DEG [HeatUp - CoolDown]
DELTA_S11 MAX INNER FIBER ANGLE 180 DEG [HeatUp - CoolDown]
DELTA_S11 MIN INNER FIBER ANGLE 180 DEG [HeatUp - CoolDown]
DELTA_S11 MAX MIDDLE FIBER ANGLE -90 DEG [HeatUp - CoolDown]
DELTA_S11 MIN MIDDLE FIBER ANGLE -90 DEG [HeatUp - CoolDown]
DELTA_S11 MAX MIDDLE FIBER ANGLE 0 DEG [HeatUp - CoolDown]
DELTA_S11 MIN MIDDLE FIBER ANGLE 0 DEG [HeatUp - CoolDown]
DELTA_S11 MAX MIDDLE FIBER ANGLE 90 DEG [HeatUp - CoolDown]
DELTA_S11 MIN MIDDLE FIBER ANGLE 90 DEG [HeatUp - CoolDown]
DELTA_S11 MAX MIDDLE FIBER ANGLE 180 DEG [HeatUp - CoolDown]
DELTA_S11 MIN MIDDLE FIBER ANGLE 180 DEG [HeatUp - CoolDown]
DELTA_S11 MAX OUTER FIBER ANGLE -90 DEG [HeatUp - CoolDown]
DELTA_S11 MIN OUTER FIBER ANGLE -90 DEG [HeatUp - CoolDown]
DELTA_S11 MAX OUTER FIBER ANGLE 0 DEG [HeatUp - CoolDown]
DELTA_S11 MIN OUTER FIBER ANGLE 0 DEG [HeatUp - CoolDown]
DELTA_S11 MAX OUTER FIBER ANGLE 90 DEG [HeatUp - CoolDown]
DELTA_S11 MIN OUTER FIBER ANGLE 90 DEG [HeatUp - CoolDown]
DELTA_S11 MAX OUTER FIBER ANGLE 180 DEG [HeatUp - CoolDown]
DELTA_S11 MIN OUTER FIBER ANGLE 180 DEG [HeatUp - CoolDown]

These headings appear on one tab-delimited header row. They are shown one per
line above only for readability.

There is only one location column: Pipeline Distance. A node-label column is
not written. Each selected pair adds exactly twenty-four adjacent result columns.
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

model_MAX_DELTA_S11_4ANGLES_INNER_FIBER.xlsx
model_MAX_DELTA_S11_4ANGLES_MIDDLE_FIBER.xlsx
model_MAX_DELTA_S11_4ANGLES_OUTER_FIBER.xlsx

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

By default, the script does not write an intermediate S11 file. To request the
compact envelope summary, add:

--write-intermediate

Example:

abaqus python extract_pipe_s11_delta_step_pairs_four_angles.py --step-pair "HeatUp" "CoolDown" --write-intermediate

For each ODB and each requested pair, this writes a file whose name looks like:

model_Pair001_HeatUp_TO_CoolDown_S11_4ANGLES_INTERMEDIATE.rpt

The intermediate report contains one wide tab-delimited table. Its first three
columns are:

Pipe Distance
Node Number
Element Number

The remaining thirty-six columns are three groups of twelve:

1. HEATUP MAX S11: maximum over all heat-up frames;
2. HEATUP MIN S11: minimum over all heat-up frames; and
3. COOLDOWN LAST S11: S11 from the final cool-down frame only.

Within each group, columns cover inner, middle, and outer fibers at -90, 0, 90,
and 180 degrees. No other angles are written or retained by this optimized
variant. The report header records both step names, their 1-based positions and
frame counts, and the zero-based cool-down frame used.

Each row is one exact element-node location. Pipe Distance is that node's
distance along the START-to-END path. If two adjacent pipe elements share a
node, they appear in separate rows with the same node distance; their S11
values are never averaged together. Blank cells mean that a matching result
was unavailable. Rows are ordered by pipe distance, node number, and element
number.

No raw heat-up or cool-down frame rows are written. The compact report is
normally far smaller than the former frame-by-frame intermediate report and is
not copied into the three final Excel workbooks.


How S11 is obtained and matched
-------------------------------

The script first looks for an exact S11 field-output key. It also supports S11
stored as the S11 component of the parent S field.

S11 is requested at ELEMENT_NODAL position so every result is associated with
a pipe path node. Abaqus may extrapolate integration-point results when it
creates this read-only subset. The ODB itself is not modified.

The heat-up step is enveloped over all of its frames to obtain both a maximum
and a minimum at every location. Those heat-up envelopes are compared with the
cool-down step's last-frame values only when node label, element label, and
section point match. Section points are matched by section-point number when
that number is available; their descriptions are used only when no number is
available. The controlling heat-up frame may be any heat-up frame; the
cool-down control is always its zero-based last frame. Duplicate finite values
at an identical location within one frame are averaged. NaN and infinite
values are ignored.

The optional intermediate table keeps every matching element-node location
separate. Adjacent elements sharing a node therefore appear as separate rows
at the same node distance and are never averaged together. This presentation
does not alter the final .rpt or Excel Delta S11 results.


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
discarded during every frame scan and are not included in the envelopes,
intermediate report, final path table, or Excel workbooks.

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

extract_pipe_s11_delta_step_pairs_four_angles.log

The log records selected paths, step pairs, frame counts, matching statistics,
S11 source fields, both MAX and MIN calculation definitions, the final .rpt
path, all three Excel paths, warnings, errors, and tracebacks.

To print live frame-by-frame progress and output paths, add:

--verbose

Example:

abaqus python extract_pipe_s11_delta_step_pairs_four_angles.py --odb "model.odb" --step-pair 22 23 --verbose

The script prints a completion line after every processed frame, for example:

Pair 1/1, heat-up step 'Step-22' (all frames): scan 15/86 complete (17.4%); ODB frame=14; finite samples=24576, cumulative scanned=368640
Pair 1/1, cool-down step 'Step-23' (last frame used): scan 1/1 complete (100.0%); ODB frame=102; finite samples=24576, cumulative scanned=24576

Each progress line identifies the step pair, heat-up or cool-down role, scan
count, actual zero-based ODB frame, percentage, finite samples in that frame,
and cumulative scanned samples. All heat-up frames and only the final
cool-down frame are scanned whether or not --write-intermediate is used.
Output is flushed immediately. A single large frame can still take time before
its completion line appears.

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
outputs.


Independence
------------

This script does not import or run any other postprocessing script in the
folder. Only extract_pipe_s11_delta_step_pairs_four_angles.py is required on the Abaqus
computer.


Troubleshooting
---------------

1. The run fails because a step was not found

   Run --list-steps on the same ODB. Use an exact displayed name or its 1-based
   position. A step pair must resolve to two different steps.

2. An element-set name is wrong

   Open extract_pipe_s11_delta_step_pairs_four_angles.log. It records the requested name,
   traceback, and available instance- and assembly-level element-set names.
   Use --list-pipe-element-sets if a screen listing is intentionally wanted.

3. Final cells are blank

   Confirm S11 was requested as field output for the relevant steps and pipe
   elements. Also confirm that the heat-up envelope and cool-down last frame
   contain matching node, element, and section-point locations. Matching counts
   and source fields are recorded in the log.

4. The steps have different numbers of frames

   This is allowed. All heat-up frames are used for the heat-up maximum and
   minimum; only the cool-down last frame is used. Frames are not paired and
   the two frame counts do not have to match. The log records both frame counts
   and the exact zero-based cool-down frame used.

5. The selected elements produce no output

   Confirm that the selection contains PIPE elements on the resolved START-to-
   END route. The log records selected elements that were outside the route.

6. The intermediate report is larger than needed or the run is slow

   The report already contains only three four-angle S11 summaries per pair.
   Restrict the run with --odb and, if appropriate, pipe element sets, labels,
   or label ranges.

7. Excel reports that a workbook is damaged

   Delete the partial .xlsx file and rerun the command. This can occur if the
   Abaqus process was terminated while the workbook was being written. The
   diagnostic log identifies the intended Excel output paths.
