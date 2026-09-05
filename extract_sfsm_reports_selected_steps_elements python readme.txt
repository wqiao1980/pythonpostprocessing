Abaqus SFSM Selected-Elements Report - User Instructions
========================================================

Purpose
-------

extract_sfsm_reports_selected_steps_elements.py outputs ESF1, SF, and SM for
user-selected element sets, exact element labels, and/or inclusive element
label ranges. It can process one ODB or every ODB in a folder and can output
all steps, named steps, or an inclusive range of steps.

This is a new, self-contained script. It does not import, modify, overwrite,
or require either original SFSM Python script.


How selected-element results are calculated
--------------------------------------------

ESF1, SF, and SM are extrapolated to element nodes. At a node shared by two
or more selected elements, only contributions from the selected elements are
arithmetically averaged. Contributions from adjacent unselected elements are
not included.

The report remains nodal, like the original SFSM report. Its columns are:

- Node Label
- ESF1
- SF.SF1, SF.SF2, and SF.SF3
- SM.SM1, SM.SM2, and SM.SM3

Element sets, individual labels, and ranges are combined as a union.


Default behavior
----------------

If no --element-set, --element, or --element-range option is provided, the
script retains the original SFSM behavior and processes the node set:

PART-1-1.START

In this fallback mode, element-nodal contributions from every connected
element can be included in the nodal average, matching the original script.

If no step option is provided, every ODB step is processed. The last frame
containing ESF1, SF, and SM is used in each selected step.


Required software and command
-----------------------------

Run the script from an Abaqus Command Prompt with "abaqus python". Abaqus/CAE
or Viewer does not need to be opened.

The script is self-contained; no companion Python file is required.


List steps
----------

abaqus python extract_sfsm_reports_selected_steps_elements.py --list-steps

For ODBs in another folder:

abaqus python "C:\python_aba\takeoutSFSM\extract_sfsm_reports_selected_steps_elements.py" --input-dir "C:\Data\BPTiberFL6" --list-steps


List element sets
-----------------

List the instance- and assembly-level element sets containing elements from
the selected instance:

abaqus python extract_sfsm_reports_selected_steps_elements.py --instance "PART-1-1" --list-element-sets

For one ODB only:

abaqus python extract_sfsm_reports_selected_steps_elements.py --input-dir "C:\Data\BPTiberFL6" --odb "model.odb" --instance "PART-1-1" --list-element-sets

This operation does not create a report.


Select element sets
-------------------

abaqus python extract_sfsm_reports_selected_steps_elements.py --element-set "SET_A"

Repeat the option to form a union:

abaqus python extract_sfsm_reports_selected_steps_elements.py --element-set "SET_A" --element-set "SET_B"

Both instance- and assembly-level element sets are supported. Set-name
matching is not case-sensitive.


Select exact element labels
---------------------------

One label:

abaqus python extract_sfsm_reports_selected_steps_elements.py --element 1001

Several labels:

abaqus python extract_sfsm_reports_selected_steps_elements.py --element 1001 1002 1005

The option may also be repeated:

abaqus python extract_sfsm_reports_selected_steps_elements.py --element 1001 --element 1002 1005


Select inclusive element-label ranges
-------------------------------------

abaqus python extract_sfsm_reports_selected_steps_elements.py --element-range 1001 1200

Repeat the option to select multiple ranges:

abaqus python extract_sfsm_reports_selected_steps_elements.py --element-range 1001 1200 --element-range 2001 2200


Combine sets, labels, and ranges
--------------------------------

This example processes the union of SET_A, element 1500, and elements 2001
through 2200:

abaqus python extract_sfsm_reports_selected_steps_elements.py --element-set "SET_A" --element 1500 --element-range 2001 2200

The default instance for explicit element selection is PART-1-1. Use another
instance with:

--instance "PIPELINE-1"


Select steps
------------

Selected exact step names:

abaqus python extract_sfsm_reports_selected_steps_elements.py --steps "Step-2" "Step-5"

Inclusive 1-based step positions 3 through 7:

abaqus python extract_sfsm_reports_selected_steps_elements.py --step-range 3 7

Inclusive range by exact first and last step names:

abaqus python extract_sfsm_reports_selected_steps_elements.py --step-range "Preload" "Operation"

Element and step selections can be combined:

abaqus python extract_sfsm_reports_selected_steps_elements.py --element-set "SET_A" --element-range 2001 2200 --step-range 3 7


Process one or all ODBs
-----------------------

Without --odb, every ODB directly inside --input-dir is processed:

abaqus python "C:\python_aba\takeoutSFSM\extract_sfsm_reports_selected_steps_elements.py" --input-dir "C:\Data\BPTiberFL6" --output-dir "C:\Data\BPTiberFL6\02_Results" --element-set "SET_A"

Process one ODB:

abaqus python extract_sfsm_reports_selected_steps_elements.py --input-dir "C:\Data\BPTiberFL6" --odb "model.odb" --element 1001 1002

The --odb option may be repeated.


Output files
------------

The default report name is deliberately different from the original SFSM
report so the original output cannot be overwritten accidentally:

model.odb -> model_SFSM_SELECTED_ELEMENTS.rpt

Reports are written beside the input ODB unless --output-dir is provided.
Existing reports from this new script with the same name are overwritten.

For exactly one --odb, a custom report name can be provided:

--output-name "custom_sfsm_elements.rpt"


Fallback node-set selection
---------------------------

When no element selection is supplied, use a different qualified node set
with:

--node-set "INSTANCE-NAME.NODE-SET-NAME"

The default is PART-1-1.START. The complete qualified name is required, and
matching is not case-sensitive.


Frame selection
---------------

The default is the last frame in each selected step containing all three
required fields:

--frame-index -1

Use a particular zero-based frame index:

--frame-index 1


Diagnostic log and console capture
----------------------------------

Every processing run creates this file in --output-dir:

extract_sfsm_reports_selected_steps_elements.log

Capture complete console output with:

abaqus python "C:\python_aba\takeoutSFSM\extract_sfsm_reports_selected_steps_elements.py" --input-dir "C:\Data\BPTiberFL6" --output-dir "C:\Data\BPTiberFL6\02_Results" --element-set "SET_A" > "C:\Data\BPTiberFL6\02_Results\sfsm_elements_console.txt" 2>&1

The > operator overwrites the console file. Use >> to append instead.


Troubleshooting
---------------

1. An element set is not found

   Use --list-element-sets with the correct --instance. Set matching is not
   case-sensitive, but the complete set name is required.

2. An exact element label is missing

   Exact labels supplied with --element must exist in --instance. The script
   stops and lists missing labels rather than silently ignoring them.

3. A range selects no elements

   Confirm the inclusive first/last labels and the selected instance. A range
   containing no labels contributes nothing to the union. If the complete
   union is empty, the script stops.

4. No element-nodal result exists at a selected node

   Confirm ESF1, SF, and SM were requested for the selected elements and
   frame. The diagnostic error lists the first affected node labels.

5. The START node set is missing

   This matters only when no explicit element selection is supplied. Use a
   qualified --node-set or provide --element-set, --element, or
   --element-range.

6. Report values differ at a shared node

   In explicit-element mode, the average includes only selected-element
   contributions. In fallback node-set mode, contributions from all connected
   elements can be included, matching the original SFSM script.
