LCCDCC Pipe Input Extractor - User Instructions
================================================

Purpose
-------

extract_lccdcc_pipe_input.py creates LCCDCC input tables from Abaqus ODB pipe
results. It is self-contained and runs with abaqus python without opening
Abaqus/CAE or Viewer.

One tab-delimited .txt file is created for every selected step in every
processed ODB. The file contains exactly these columns:

KP    KP    WD    ESF1    Axial Strain    THE    SK2    SK1    BMY    BM-Z
(m)   (m)   (m)   (N)     1                       1/m    1/m    N-m    N-m

The actual file uses tab characters between fields. The second unit line keeps
the THE unit cell blank to match the requested LCCDCC format.


Column mapping
--------------

- First KP: cumulative pipeline path distance, fixed-point format.
- Second KP: the same cumulative path distance, scientific format.
- WD: COORD3 at the final frame of the user-defined as-laid step.
- ESF1: pipe element ESF1.
- Axial Strain: pipe element SE1.
- THE: pipe element THE11.
- SK2: pipe element SK2.
- SK1: pipe element SK1.
- BMY: pipe element SM2.
- BM-Z: pipe element SM1.

Only elements whose Abaqus element type starts with PIPE are used for the path
and result extraction.

The script writes the unit labels shown above but does not convert units. The
ODB must use the intended consistent units, such as metres, newtons, and
newton-metres, before the files are used for LCCDCC.


Section-point envelope
----------------------

At each pipeline path node, the script gathers values from:

- every section point;
- every extrapolated element-nodal result at that node; and
- every selected pipe element contributing to that node.

Each output column is enveloped independently. The output is the signed value
having the largest absolute magnitude. For example, values 10, -20, and 15
produce an envelope of -20. Ties in absolute magnitude select the more positive
value.

ESF1, SE1, THE11, SK2, SK1, SM2, and SM1 do not have to come from the same
section point. Each result receives its own controlling section-point envelope.


Default extraction
------------------

By default, the script:

- processes every .odb file in the script directory;
- uses every ODB step for the pipe result columns;
- uses the final frame in every result step;
- uses COORD3 from the final frame of the required --aslaid-step for WD;
- uses every supported pipe element on the resolved start-to-end path;
- calculates KP from the original three-dimensional nodal coordinates;
- writes one LCCDCC .txt file per step; and
- prints no script-generated result or progress text on screen.

Run from an Abaqus Command Prompt:

abaqus python extract_lccdcc_pipe_input.py --aslaid-step "As-Laid"

Process all ODBs in another directory:

abaqus python "C:\python_aba\takeoutSFSM\extract_lccdcc_pipe_input.py" --input-dir "C:\Data\BPTiberFL6" --output-dir "C:\Data\BPTiberFL6\02_Results" --aslaid-step "As-Laid"

Process one ODB:

abaqus python extract_lccdcc_pipe_input.py --input-dir "C:\Data\BPTiberFL6" --odb "model.odb" --output-dir "C:\Data\BPTiberFL6\02_Results" --aslaid-step "As-Laid"


Define the as-laid water-depth step
----------------------------------

Every extraction command must include --aslaid-step. Enter its exact step name
or 1-based ODB step position.

Exact name:

--aslaid-step "As-Laid"

Position, for example the fourth ODB step:

--aslaid-step 4

WD is read once from COORD3 in the final frame of this as-laid step. The same
WD profile is then written into every selected result-step file. The as-laid
step does not have to be included in --steps or --step-range.

For a command that processes several ODBs, the supplied name or position must
resolve to the intended as-laid step in every ODB. Run ODBs separately if they
use different as-laid names or positions.


Select steps
------------

Exact step names:

abaqus python extract_lccdcc_pipe_input.py --aslaid-step "As-Laid" --steps "As-Laid" "Operation"

Inclusive 1-based step positions 3 through 7:

abaqus python extract_lccdcc_pipe_input.py --aslaid-step "As-Laid" --step-range 3 7

Inclusive range using exact first and last step names:

abaqus python extract_lccdcc_pipe_input.py --aslaid-step "As-Laid" --step-range "Preload" "Operation"

If neither --steps nor --step-range is entered, all ODB steps are processed.

List step names and positions:

abaqus python extract_lccdcc_pipe_input.py --list-steps


Select the frame
----------------

The default is the final frame in every selected step:

--frame-index -1

Select a zero-based frame index, for example frame 2:

--frame-index 2

This option controls only ESF1, SE1, THE11, SK2, SK1, SM2, and SM1. WD always
uses the final frame of --aslaid-step. If the requested result frame is not
available in a selected step, that result step is skipped and the condition is
written to the log.


Water depth
-----------

WD is taken strictly from COORD3 in the final frame of --aslaid-step. No U3
fallback is used. If COORD output or a required node's COORD3 value is missing,
processing that ODB fails and the details are written to the log. The Z sign is
preserved; the script does not convert depth to a positive magnitude.


Select pipe elements
--------------------

If no pipe element selector is entered, every pipe element on the resolved
start-to-end path is used.

One or more element sets:

--pipe-element-set "PIPE_SET_A" --pipe-element-set "PIPE_SET_B"

Exact element numbers:

--pipe-element 1001 1002 1005

Inclusive element-number ranges:

--pipe-element-range 1001 1200 --pipe-element-range 2001 2200

Sets, exact labels, and ranges form a union. Only selected pipe elements lying
on the resolved start-to-end path are output. Selected pipe elements outside
the path and non-pipe elements inside a selected set/range are omitted and
recorded in the log. Path filtering does not reset KP.

List only sets containing pipe elements:

abaqus python extract_lccdcc_pipe_input.py --instance "PART-1-1" --list-pipe-element-sets


Start and end of the pipeline path
----------------------------------

The script resolves endpoints in this order:

1. --start-node and --end-node exact labels.
2. --start-node-set and --end-node-set exact instance node-set names.
3. Unique instance node-set names containing START and END, without regard to
   letter case.
4. A missing endpoint is inferred as the farthest reachable graph endpoint.
5. If both are missing and the pipe graph has exactly two endpoints, those
   endpoints are used and the lower node label is the start.

Exact labels:

--start-node 1 --end-node 5001

Exact endpoint sets:

--start-node-set "PIPE_START_NODE" --end-node-set "PIPE_END_NODE"

List possible endpoint sets:

abaqus python extract_lccdcc_pipe_input.py --list-endpoint-sets


Output names and format
-----------------------

For model.odb and ODB step position 4 named As-Laid, the output name is:

model_Step004_As-Laid_LCCDCC_INPUT.txt

Step position is included so different step names cannot easily overwrite one
another. Characters unsuitable for Windows filenames are replaced with an
underscore.

Example with --precision 2:

KP	KP	WD	ESF1	Axial Strain	THE	SK2	SK1	BMY	BM-Z
(m)	(m)	(m)	(N)	1		1/m	1/m	N-m	N-m
0.00	0.00E+00	-1.21E+03	-6.55E+05	1.83E-03	1.80E-03	8.08E-03	-6.28E-06	2.22E+05	-1.72E+02

The default scientific-format precision is six digits after the decimal. Use
another value from 1 through 12 with:

--precision 2

Blank data cells mean that the corresponding ODB result was unavailable at
that node and frame. The first KP column always uses two digits after the
decimal; the second KP column preserves additional precision in scientific
format.


ODB result names
----------------

The script supports exact scalar output names and parent-field components:

- ESF1 or ESF.ESF1
- SE1 or SE.SE1
- THE11 or THE.THE11
- SK2/SK1 or SK.SK2/SK.SK1
- SM2/SM1 or SM.SM2/SM.SM1

Field results are requested at ELEMENT_NODAL position so integration-point
results can be extrapolated to path nodes before the section-point envelope is
formed.


Log file and quiet operation
----------------------------

Every extraction run writes:

extract_lccdcc_pipe_input.log

The log records endpoint resolution, pipe selections, the as-laid WD step and
final frame, selected result frames, field sources, missing values, output
filenames, and complete failure tracebacks.
Incorrect element-set names and available element-set lists are written to the
log rather than printed on screen.

Normal extraction is quiet. Add --verbose for short progress and output-file
messages. Detailed errors and available-set lists remain in the log. Explicit
--list-steps, --list-pipe-element-sets, and --list-endpoint-sets commands print
their requested lists because they do not perform result extraction.

The Abaqus launcher can still display its own license-manager messages; those
messages are outside this Python script.


Terminate a running extraction
------------------------------

1. Click the Abaqus Command Prompt so it has keyboard focus.
2. Press Ctrl+C.
3. If Windows asks "Terminate batch job (Y/N)?", type Y and press Enter.
4. If Ctrl+C does not respond, try Ctrl+Break and wait briefly.

Some ODB reads may not react immediately. As a last resort, use Windows Task
Manager to end only the Abaqus/Python process started by this command. Do not
terminate another analysis or an unrelated Python process.

ODBs are opened read-only and will not be damaged. A .txt or .log file being
written may be incomplete; rerun the command to overwrite partial output. The
Abaqus license should be released after the process exits.


Independence
------------

This script does not import or run another local postprocessing script. Only
extract_lccdcc_pipe_input.py is required on the Abaqus computer.


Troubleshooting
---------------

1. WD is blank

   A successful output cannot contain a blank WD. Confirm COORD field output
   exists for every path node in the final frame of --aslaid-step. The script
   intentionally does not substitute U3.

2. A result column is blank

   Confirm the requested variable was written for the pipe elements. Open the
   log to see which exact or parent field names were found and how many nodes
   received values.

3. The pipe path cannot be resolved

   Run --list-endpoint-sets and specify exact start/end nodes or node sets.
   Confirm that they belong to one connected pipe-element route.

4. An element set is not found

   Open extract_lccdcc_pipe_input.log for the requested name and all available
   set names, or explicitly run --list-pipe-element-sets.

5. Envelope signs are unexpected

   The envelope preserves the sign of the largest absolute value. Inspect the
   pipe local directions and section-point conventions in Abaqus.
