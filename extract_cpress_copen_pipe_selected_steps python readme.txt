Abaqus PIPE-Only CPRESS/COPEN Pipeline Report - User Instructions
=================================================================

Required files
--------------

Keep these two Python files together in the same folder:

1. extract_cpress_copen_pipe_selected_steps.py
2. extract_le11_the11_selected_steps_pipe_only.py

Run the commands below from an Abaqus Command Prompt using "abaqus python".
The Abaqus/CAE or Viewer window does not need to be opened.


Script dependencies
-------------------

The CPRESS/COPEN script is not self-contained. It requires this helper file
in the same folder:

extract_le11_the11_selected_steps_pipe_only.py

The helper supplies ODB selection, step selection, PIPE-only path-distance,
and Excel-writing utilities. Importing it does not run an LE11/THE11
extraction.

The dependency is one-way: the CPRESS/COPEN script needs the PIPE-only LE11
script, but the PIPE-only LE11 script does not need the CPRESS/COPEN script.


PIPE-only behavior
------------------

The pipeline path is constructed only from element types whose names begin
with PIPE, such as PIPE31 and PIPE32. Non-PIPE elements are excluded from the
path.

CPRESS and COPEN values are retained only for nodes on this PIPE path.

Important: contact field output can be stored at contact/element-nodal
locations. Its elementLabel may identify a contact or surface representation
rather than the underlying PIPE mesh element. Version 2026-09-05-r5 and later
therefore restricts results by PIPE-path node membership and does not reject a
valid contact value merely because its contact element label is not a PIPE
element label.

Users can further restrict output by element set, inclusive element-label
range, or a union of multiple sets and ranges. If no element selection option
is provided, all PIPE elements are processed.

Version 2026-09-05-r6 changes explicit element selection for speed. When
--element-set or --element-range is supplied, the selected elements are
trusted and their element types are not checked again. The user must ensure
that every requested set/range contains the intended pipeline elements. The
complete PIPE path is still identified once to preserve the original
full-pipeline distances.

For nodal contact output, Abaqus may not provide an element label. Such a
value is retained when its node belongs to the PIPE path. Therefore, a node
shared by a PIPE element and another element remains a PIPE-path node.


Contact-value handling
----------------------

A node can have more than one contact result because of multiple contact
surfaces or contributions. For each selected step and PIPE-path node:

- CPRESS MAX is the maximum available CPRESS value.
- COPEN MIN is the minimum available COPEN value.
- P Samples is the number of retained CPRESS contributions.
- O Samples is the number of retained COPEN contributions.

Missing contact output is left blank. The script does not replace a missing
value with zero.

Abaqus can store a very large finite placeholder, commonly with a magnitude
near 1E38, when a contact result is undefined or unavailable. Script version
2026-09-05-r4 and later treats CPRESS or COPEN values whose absolute magnitude
is at least 1.0E30 as undefined. Version 2026-09-05-r5 also corrects the
contact element-label filtering that could discard a real COPEN value at the
same PIPE node. Undefined values are left blank rather than reported as
physical results. The text report and console show how many such values were
discarded.

The default limit can be changed if necessary. For example:

abaqus python extract_cpress_copen_pipe_selected_steps.py --undefined-value-limit 1.0E35

The limit applies to both positive and negative values and to both CPRESS and
COPEN. It should remain well above every physically possible result in the
model.

CPRESS uses the model stress units. COPEN and path distance use the model
length units.


Results produced by the script
------------------------------

For every processed ODB, the script creates:

1. model_PIPE_ONLY_CPRESS_COPEN.rpt
2. model_PIPE_ONLY_CPRESS_COPEN_along_path.xlsx

The text report contains one table for every selected step, with these
columns:

- Path ID
- Path Distance
- Node Label
- X, Y, and Z coordinates
- CPRESS MAX
- COPEN MIN
- P Samples
- O Samples

It also reports the maximum CPRESS and minimum COPEN location for each step.

The Excel workbook has separate CPRESS and COPEN worksheets for every
disconnected PIPE path. Each worksheet contains path distance, node label,
and one column for every selected step. Each worksheet also contains a native
Excel scatter chart.

The charts are editable in Excel. Users can change chart type, colors, line
styles, markers, titles, axes, legend, size, and layout. Excel does not need
to be installed or open when Abaqus creates the workbook.


Process all ODBs and all steps
------------------------------

This command processes every ODB in the script folder and every step
containing both CPRESS and COPEN:

abaqus python extract_cpress_copen_pipe_selected_steps.py

To process every ODB in another folder:

abaqus python "C:\python_aba\takeoutSFSM\extract_cpress_copen_pipe_selected_steps.py" --input-dir "C:\Data\BPTiberFL6" --output-dir "C:\Data\BPTiberFL6\02_Results"


List the available steps
------------------------

List the ordered steps without creating output:

abaqus python extract_cpress_copen_pipe_selected_steps.py --list-steps

List steps in another folder:

abaqus python "C:\python_aba\takeoutSFSM\extract_cpress_copen_pipe_selected_steps.py" --input-dir "C:\Data\BPTiberFL6" --list-steps


List the available element sets
-------------------------------

List only instance- and assembly-level element sets that contain at least one
PIPE element in the selected pipeline instance. The displayed count is the
number of PIPE elements in each listed set:

abaqus python extract_cpress_copen_pipe_selected_steps.py --list-element-sets

For ODBs in another folder:

abaqus python "C:\python_aba\takeoutSFSM\extract_cpress_copen_pipe_selected_steps.py" --input-dir "C:\Data\BPTiberFL6" --instance "PART-1-1" --list-element-sets

This operation does not create reports or Excel workbooks.

Version 2026-09-05-r6 skips every set that contains no PIPE elements. It
caches the PIPE labels once per ODB and uses Abaqus bulk label access when
available, making the inventory faster and its console output shorter for
large models with many element sets.

Without --odb, every ODB in --input-dir is still opened and inventoried. For
the fastest check of one model, specify its filename:

abaqus python extract_cpress_copen_pipe_selected_steps.py --input-dir "C:\Data\BPTiberFL6" --odb "model.odb" --instance "PART-1-1" --list-element-sets


Restrict output by element set
------------------------------

Use --element-set followed by an exact instance- or assembly-level element
set name:

abaqus python extract_cpress_copen_pipe_selected_steps.py --element-set "CONTACT_ZONE"

Repeat --element-set to combine several sets:

abaqus python extract_cpress_copen_pipe_selected_steps.py --element-set "ZONE_A" --element-set "ZONE_B"

Set-name matching is not case-sensitive. All elements contained in a
requested set are accepted without checking element.type. Make sure the set
contains the intended pipeline elements.


Restrict output by element-number range
---------------------------------------

Use --element-range FIRST LAST. Both labels are included:

abaqus python extract_cpress_copen_pipe_selected_steps.py --element-range 1001 1200

Repeat the option to select multiple ranges:

abaqus python extract_cpress_copen_pipe_selected_steps.py --element-range 1001 1200 --element-range 2001 2200

All existing element labels inside the requested ranges are accepted without
checking element.type. Make sure the ranges contain the intended pipeline
elements.


Combine element sets and ranges
-------------------------------

Element sets and ranges are combined as a union. This example selects all
elements in CONTACT_ZONE plus elements numbered 3001 through 3200:

abaqus python extract_cpress_copen_pipe_selected_steps.py --element-set "CONTACT_ZONE" --element-range 3001 3200

Element selection can be combined with step selection:

abaqus python extract_cpress_copen_pipe_selected_steps.py --element-set "CONTACT_ZONE" --element-range 3001 3200 --step-range 3 7

Complete example for another folder:

abaqus python "C:\python_aba\takeoutSFSM\extract_cpress_copen_pipe_selected_steps.py" --input-dir "C:\Data\BPTiberFL6" --output-dir "C:\Data\BPTiberFL6\02_Results" --element-set "CONTACT_ZONE" --element-range 3001 3200 --steps "Step-2" "Step-5"


Output selected steps by name
-----------------------------

List each exact step name after --steps. Use quotation marks around names
that contain spaces:

abaqus python extract_cpress_copen_pipe_selected_steps.py --steps "Step-2" "Step-5" "Step-8"

Step-name matching is not case-sensitive. A requested step that is missing
from an ODB is omitted, and the script prints a note.


Output a range using step positions
-----------------------------------

Step positions are 1-based and the range is inclusive. This example outputs
steps 3, 4, 5, 6, and 7:

abaqus python extract_cpress_copen_pipe_selected_steps.py --step-range 3 7

If an ODB has fewer steps than the requested ending position, the script uses
the last available step and prints a note.


Output a range using step names
-------------------------------

The first and last named steps are both included:

abaqus python extract_cpress_copen_pipe_selected_steps.py --step-range "Preload" "Operation"


Process one ODB only
--------------------

Add --odb followed by the ODB filename:

abaqus python extract_cpress_copen_pipe_selected_steps.py --odb "model.odb" --step-range 3 7

For an ODB in another folder, combine --input-dir and --odb:

abaqus python "C:\python_aba\takeoutSFSM\extract_cpress_copen_pipe_selected_steps.py" --input-dir "C:\Data\BPTiberFL6" --odb "model.odb" --steps "Step-2" "Step-5"

The --odb option may be repeated to process several named ODB files.


Path distance and starting node
-------------------------------

Path distance is first calculated along the complete undeformed PIPE-element
mesh using the ODB nodal coordinates. Element-set/range filtering is then
applied. Therefore, selected segments retain their original full-pipeline
distance instead of restarting at zero at the first selected element.

The default pipeline instance is:

PART-1-1

The default path origin is obtained from the instance node set START. If this
set is unavailable, the script selects an endpoint automatically.

Use another instance:

--instance "PIPELINE-1"

Use another starting node set:

--start-node-set "PIPE_START"

Use a particular node label as distance zero. This overrides
--start-node-set:

--start-node 1001

If the PIPE mesh contains disconnected components, each component receives a
separate Path ID and separate CPRESS/COPEN Excel worksheets.


Frame selection
---------------

By default, the script searches backward in each selected step and uses the
last frame containing both CPRESS and COPEN:

--frame-index -1

Use a particular zero-based frame index, for example frame 1:

--frame-index 1

A step is skipped if the requested frame is unavailable or does not contain
both fields. The reason is printed in the console.


Output location and custom report name
--------------------------------------

Output is written to the input ODB folder unless --output-dir is used:

--output-dir "C:\path\to\results"

When processing exactly one ODB, assign a custom text-report name with:

--output-name "custom_contact_report.rpt"

The Excel workbook continues to use the ODB-based name:

model_PIPE_ONLY_CPRESS_COPEN_along_path.xlsx

Existing output files with the same names are overwritten.


Diagnostic log and console capture
----------------------------------

Every processing run creates this diagnostic log in --output-dir:

extract_cpress_copen_pipe_selected_steps.log

To capture the complete console output as well:

abaqus python "C:\python_aba\takeoutSFSM\extract_cpress_copen_pipe_selected_steps.py" --input-dir "C:\Data\BPTiberFL6" --output-dir "C:\Data\BPTiberFL6\02_Results" > "C:\Data\BPTiberFL6\02_Results\contact_console.txt" 2>&1

The > operator overwrites contact_console.txt. Use >> to append instead.


Example console output
----------------------

extract_cpress_copen_pipe_selected_steps.py version 2026-09-05-r6
Opening: C:\Data\Results\model.odb
Selected steps:
  Step-2 (frame index 10)
  Step-5 (frame index 18)
Undefined/non-finite values discarded: CPRESS=0, COPEN=28
Wrote:      C:\Data\Results\model_PIPE_ONLY_CPRESS_COPEN.rpt
Excel plot: C:\Data\Results\model_PIPE_ONLY_CPRESS_COPEN_along_path.xlsx
Completed: 1 succeeded, 0 failed.


Troubleshooting
---------------

1. ImportError for extract_le11_the11_selected_steps_pipe_only

   Keep extract_le11_the11_selected_steps_pipe_only.py in the same folder as
   the CPRESS/COPEN script.

2. No frame contains both CPRESS and COPEN

   Confirm that CPRESS and COPEN were requested as contact field output in
   the analysis. Different variables may be written at different frame
   frequencies.

3. A PIPE node has a blank CPRESS or COPEN cell

   A blank means that the ODB contains no retained value at that node in the
   selected frame. It is not automatically equivalent to zero contact
   pressure or zero opening.

4. No connected PIPE elements were found

   Confirm the correct instance name with --instance and confirm that its
   element types begin with PIPE.

5. Only some open gaps have COPEN

   Abaqus may provide COPEN only for nodes in contact or sufficiently close
   to contact, depending on the contact formulation and tracking settings.

6. Requested sets or ranges produce no pipeline-path nodes

   Run --list-element-sets and confirm the intended set is listed. For ranges,
   confirm that the inclusive labels exist in the selected instance and are
   the intended pipeline elements. Explicit selections are trusted and their
   element types are not checked.

7. A boundary node also belongs to an unselected element

   Selection is element-based, but CPRESS/COPEN output is organized along
   nodes. Every node belonging to a selected PIPE element is retained,
   including an end node shared with an adjacent unselected element.

8. COPEN is approximately 1E38 or -1E38

   This is an undefined/unavailable Abaqus placeholder, not a physical gap.
   Version 2026-09-05-r4 and later removes it automatically using the default
   absolute-value limit of 1.0E30. Version 2026-09-05-r5 also retains valid
   contact/element-nodal values whose contact element label differs from the
   PIPE mesh element label. Rerun the extraction to replace workbooks created
   by an older script. If legitimate model results can exceed the default
   limit, set a higher value with --undefined-value-limit.
