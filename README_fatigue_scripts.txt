FATIGUE DAMAGE ALONG PIPELINE KP FROM ABAQUS ODB FILES
======================================================

Files in C:\SeabedAnalysis

  extract_stress_ranges.py     Step 1. Reads the ODB(s), writes delta-S11 reports.
  extract_contact.py           Optional. COPEN / CPRESS along KP for the plots.
  extract_tension_depth.py     Optional. Effective tension and Z coordinate along KP.
  extract_mech_strain.py       Optional. Max / min mechanical strain along KP.
  PipelineResultPlots.xlsm     Separate workbook that plots those two.
  run_fatigue_workbook.py      Step 2. Feeds the reports to the Excel workbook.
  FatigueDamageCal_ManualWorkflows_INPUT.xlsm
                               The fatigue workbook with the INPUT tab (never
                               modified by the scripts). All fatigue inputs and
                               the DBM KP ranges are entered on its INPUT tab.
  FatigueDamageCal_ManualWorkflows_Hydrotest_Design.xlsm
                               The earlier workbook, unchanged.
  FatigueDamageManualWorkflow.bas
                               Copy of the macro module in the INPUT workbook.

Status: the extraction logic was tested on a fake ODB only, and step 2 has not
been run in Excel yet. Check the first results by hand.


WHAT YOU NEED
-------------
  Step 1: Abaqus (uses "abaqus python", no extra packages).
  Step 2: Windows, Excel, Python 3 and pywin32  (pip install pywin32).
  ODB:    field output S at the inner and outer section points of the pipe
          elements at angles -90, 0, 90 and 180, for every step used in a pair.


STEP 1 - EXTRACT STRESS RANGES
------------------------------
Everything is given on the command line; the script does not need editing.
Open an Abaqus command window in C:\SeabedAnalysis.

1. Check the ODB first:

     abaqus python extract_stress_ranges.py --odb model.odb --list

   This prints the step numbers and names, and the section points of the pipe
   elements. Points marked with * are the ones the script will use; you
   should see INNER and OUTER at -90, 0, 90 and 180. If none are marked, fill
   in SECTION_POINT_MAP at the top of the script, for example
     SECTION_POINT_MAP = {4: ("INNER", -90), 6: ("OUTER", -90), ...}

2. Run the extraction, for example:

     abaqus python extract_stress_ranges.py --odb model.odb --step-pair 22 23 14 16 --phase "DESIGN OPERATION" HYDROTEST --s11 --abs-max-s11

   Options
     --odb FILE          ODB file.
     --step-pair A1 B1 [A2 B2 ...]
                         Steps in pairs. Stress range = S11(A) - S11(B), so
                         "22 23 14 16" gives the two ranges 22-23 and 14-16.
                         A step is its number in the ODB (see --list) or a
                         unique part of its name in quotes ("Step 014").
     --phase P1 [P2 ...] Workbook phase of each pair, in the same order:
                         "EARLY UNDRAINED", "EARLY DRAINED", "MIDDLE DRAINED",
                         "LATE DRAINED", HYDROTEST or "DESIGN OPERATION".
                         Needed for step 2; one value applies to all pairs.
     --load-case L1 [L2 ...]
                         FCD, HCD, PCD, HYDROTEST1, HYDROTEST2 or DESIGN for
                         each pair (default FCD). The last three go to the
                         hydrotest and design blocks whatever the phase.
     --label N1 [N2 ...] Report name of each case (default StepA-StepB). Must
                         not contain INNER, OUTER, ANGLE, DISTANCE or
                         square brackets.
     --envelope          Combine the pairs that share the same phase and load
                         case into one report each, keeping the larger
                         magnitude at each KP point and position. Example:
                           --step-pair 19 20 21 20 14 15 16 17
                           --load-case FCD FCD HYDROTEST HYDROTEST
                           --phase "EARLY UNDRAINED" --envelope
                         gives EARLY_UNDRAINED_FCD_env.rpt (pairs 19-20 and
                         21-20) and EARLY_UNDRAINED_HYDROTEST_env.rpt (pairs
                         14-15 and 16-17). If phase and load case are not
                         given per pair, all pairs are combined into one.
     --s11               Also write S11 of every requested step.
     --u2                Also write U2_steps.rpt: local U2 along KP at the
                         nodes for every requested step (same conversion as
                         described under --compare-tol). Import it with
                         button 6 of the workbook to get the U2 plot.
     --abs-max-s11 [1|2|both]
                         Also write the absolute maximum S11 over ALL frames
                         of the first step of each pair (1 = heat-up step, the
                         default), the second step (2) or both.
     --range-from-abs-max
                         Stress range = absolute maximum S11 over all frames
                         of the first (heat-up) step, with its sign, minus S11
                         at the last frame of the second (cool-down) step.
                         Without it, the last frame of both steps is used.
     --compare-tol PCT   When several step pairs share the same phase and load
                         case, pair_comparison.txt lists each pair's maximum
                         stress range with its KP, fibre and angle, and the
                         difference from the largest of the group. If that
                         difference exceeds PCT percent (default 10), the
                         script also writes, for the steps of that group:
                           CHECK_<phase>_<load case>_U2.rpt     local U2 along KP
                           CHECK_<phase>_<load case>_COPEN.rpt  contact opening
                           CHECK_<phase>_<load case>.png        both plotted
                         U2 shows how the buckle shape changes per cycle and
                         COPEN shows upheaval (pipe lifting off). The PNG needs
                         matplotlib in the Abaqus Python; without it only the
                         .rpt files are written.
                         Local U2: the ODB stores displacements in the global
                         system, so the script projects them on the horizontal
                         direction normal to the as-laid pipe axis at each
                         node (positive to the left looking towards increasing
                         KP).
     --vertical-axis A   Global vertical axis for that conversion: x, y or z
                         (default z).
     --u2-as-stored      Skip the conversion and use the stored U2 component.
     --copen-key TEXT    Use only the COPEN output variables whose name
                         contains TEXT (e.g. the seabed surface). Default: all
                         of them, taking the smallest opening at each node.
     --frame N           Frame used for the stress range and --s11
                         (default -1 = last frame of the step).
     --elset NAME        Pipeline element set (default PIPE_ELEMENTS).
     --instance NAME     Instance, if the ODB has more than one.
     --kp-start X        KP in metres at the start of the first element (default 0).
     --out-dir DIR       Output folder (default fatigue_reports).
     -h                  Show all options.

   Output, in the output folder:
     <label>.rpt      one tab-separated report per pair: KP in m, then delta S11
                      in Pa at INNER/OUTER fibre, angles -90, 0, 90, 180
     manifest.json    list of the reports with phase and load case (read by step 2)
     S11_steps.rpt    with --s11: S11 in Pa of each requested step, same
                      fibres and angles, eight columns per step
     U2_steps.rpt     with --u2: local U2 of each requested step along KP
     ABSMAX_S11.rpt   with --abs-max-s11: for each heat-up step, the S11 of
                      largest magnitude over all frames (sign kept), in Pa

   The screen shows the number of KP rows and the largest value of each
   case, with its KP.

3. Several ODBs, or cases you want to keep: instead of --step-pair, list them
   in CASES at the top of the script (each case can name its own "odb") and
   run the script without --step-pair. Unit factors LENGTH_TO_M and
   STRESS_TO_PA are also there; 1.0 for a model in metres and pascals.

Notes
  - KP is the cumulative element length from the first element, taken at the
    middle of each element. The DBM KP ranges in the workbook (sheet
    StressRangeID, columns BN to BP) must use the same origin.
  - All ODBs used for one workbook must have the same mesh. The script stops
    if their KP grids differ.


COPEN / CPRESS ALONG KP (extract_contact.py)
--------------------------------------------
Keep extract_contact.py in the same folder as extract_stress_ranges.py (it
uses its KP routines). The ODB needs COPEN and CPRESS in *CONTACT OUTPUT.

     abaqus python extract_contact.py --odb model.odb --list
     abaqus python extract_contact.py --odb model.odb --step 22 23
     abaqus python extract_contact.py --odb model.odb --step 22 23 --elset DBM_MIDDLE DBM_SHOULDER

   Options
     --odb FILE          ODB file.
     --step S1 [S2 ...]  Steps to report (number or part of the name; default
                         the last step). One column per step.
     --var COPEN CPRESS  Variables to extract (default both).
     --elset N1 [N2 ...] Element sets to report, one report per set (default
                         the whole pipeline). Nodes of a set that are not
                         pipeline nodes, e.g. a DBM surface, get the KP of the
                         nearest pipeline node; several nodes at one KP give
                         the smallest COPEN / largest CPRESS.
     --pipe-elset NAME   Pipeline set that defines KP (default PIPE_ELEMENTS).
     --contact-key TEXT  Use only the contact output variables whose name
                         contains TEXT (e.g. the DBM surface name). Default:
                         all, taking the smallest COPEN / largest CPRESS.
     --frame N           Frame in each step (default -1 = last).
     --instance NAME     Pipeline instance, if the ODB has more than one.
     --kp-start X        KP in metres at the start of the first element.
     --out-dir DIR       Output folder (default contact_reports).
     --list              Print steps, element sets and contact output names.

   Output: COPEN_<set>.rpt and CPRESS_<set>.rpt (KP in m, then one column per
   step, in model units). Import them with button 7 of the workbook.


EFFECTIVE TENSION, Z COORDINATE AND MECHANICAL STRAIN (PipelineResultPlots.xlsm)
--------------------------------------------------------------------------------
Keep both scripts in the same folder as extract_stress_ranges.py.

1. Effective tension and Z coordinate (needs ESF1 in *ELEMENT OUTPUT):

     abaqus python extract_tension_depth.py --odb model.odb --list
     abaqus python extract_tension_depth.py --odb model.odb --step-range 5 8
     abaqus python extract_tension_depth.py --odb model.odb --step-range 5 8 12 14 --step 22

     --step-range A B [A B ...]  Step numbers, first and last of each range.
     --step S1 [S2 ...]          Single steps (number or part of the name).
                                 No step given: the as-laid step (found by
                                 "aslaid" / "as-laid" in the step name).
     --z-step S1 [...] | all     Steps for the Z coordinate. Default: the
                                 as-laid step only; "all" = every requested step.
     --force-key NAME            Output to use instead of ESF1. Without ESF1
                                 in the ODB the script uses SF1 and warns that
                                 SF1 is the wall force, not effective tension.
     --vertical-axis x|y|z       Global vertical axis (default z).
     --frame, --elset, --instance, --kp-start, --out-dir as in the other scripts.

   Output in pipeline_reports: EFFTENSION.rpt (at element mid-length) and
   ZCOORD.rpt (at the nodes; COORD output if present, else initial coordinate
   + U). One column per step, model units.

2. Mechanical strain = LE11 - THE11 (needs LE and THE in *ELEMENT OUTPUT at
   the section points):

     abaqus python extract_mech_strain.py --odb model.odb --step-range 20 23
     abaqus python extract_mech_strain.py --odb model.odb --step 22 23 --var MECH LE11 THE11

     --step-range / --step       As above. No step given: the last step.
     --var MECH LE11 THE11       What to output: MECH = LE11 - THE11 (default),
                                 LE11, THE11; several allowed.
     --all-frames                Maximum / minimum over all frames of each
                                 step instead of the last frame.
     --no-thermal                Use LE11 alone if the ODB has no THE output.

   For each element the script takes the maximum and the minimum over ALL
   section points and integration points. Output in pipeline_reports:
   MECHSTRAIN.rpt, LE11.rpt, THE11.rpt, each with a MAX and a MIN column per
   step, strain as a fraction.

3. Plot: open PipelineResultPlots.xlsm.
   - INPUT tab: DBM KP ranges and curve-section KP ranges. Button 3 copies
     both tables from the INPUT tab of the fatigue workbook.
   - Button 1: select the .rpt files (hold Ctrl for several). The EffTension
     tab gets the effective tension and Z coordinate charts, the MechStrain
     tab the strain charts, with amber DBM bands and blue curve bands. The
     chart data are on the same tabs, from column Q.
   - Button 2 redraws the charts after the KP ranges change.
   - A new import replaces the same kind of report; up to 77 data columns
     per report.
   PipelineResultPlots.bas is a copy of the workbook's macro module.


STEP 2 - RUN THE FATIGUE WORKBOOK
---------------------------------
1. Before the first run, check the INPUT tab of the workbook: cycle counts,
   SCF, wall thickness, S-N curves and the DBM KP ranges.
   Optional: in the "User S-N curve KP ranges" table on the INPUT tab, enter
   a KP start and end with m, C2 and KDF for the INNER and/or OUTER fibre.
   KP inside such a range uses that curve instead of the DBM/RB curve on
   every fatigue tab. Run button 4 again after changing the KP ranges.
   Optional: in the "Result KP ranges" table on the INPUT tab, enter a KP
   start and end. Button 5 then lists, on the Fatigue_KPRangeSUM tab, the
   maximum total damage, UC and KP at the maximum inside each range.
   Plots: button 5 also rebuilds the Plots tab with charts along KP for the
   stress range (one line per imported phase / load case), the total fatigue
   damage and the UC, for the inner and the outer fibre. Each chart shows
   the DBM KP ranges of the INPUT tab as amber bands from KP start to KP end.
   U2 plot (optional): button 6 on "Workflow Controls" asks for a U2 report
   (U2_steps.rpt from --u2, or a CHECK_..._U2.rpt). It copies the report to
   the U2 tab and adds a chart of local U2 along KP, one line per step, with
   the DBM bands, below the UC charts on the Plots tab. The chart is kept
   when button 5 rebuilds the Plots tab.
   COPEN / CPRESS plots (optional): button 7 asks for the reports written by
   extract_contact.py (hold Ctrl to pick several, e.g. one per element set).
   They go to the COPEN and CPRESS tabs, one block of columns per set, and
   each variable gets a chart on the Plots tab with one line per set and
   step. Button 7 replaces what was imported before, so select all the
   reports you want to see together.
   Curve sections (optional): in the "Curve section KP ranges" table on the
   INPUT tab (columns Y:AA) enter a KP start and end per route curve. They
   are drawn as blue bands on every chart, next to the amber DBM bands. Press
   button 5 (or import a report with button 6 or 7) to redraw the charts.
2. Check the three paths at the top of run_fatigue_workbook.py
   (WORKBOOK, OUTPUT, MANIFEST). Close the OUTPUT workbook if it is open.
3. Run:

     python run_fatigue_workbook.py

   The script copies the workbook to OUTPUT, then for each case sets the
   phase and load case on the "Workflow Controls" sheet and runs the
   workbook's own macro (import, fatigue damage, summary). It prints OK or
   the macro's error message per case, then the maximum damage per fatigue
   sheet for DBM and rough-bottom sections, and saves OUTPUT.

   Open OUTPUT (FatigueDamageCal_Results.xlsm) for the damage along KP,
   the summaries and the charts.


MANUAL ALTERNATIVE TO STEP 2
----------------------------
A report can also be imported by hand: on "Workflow Controls" set the target
phase and load case, press button 1 and pick the .rpt file, then buttons 4
and 5 (delta S11 reports skip buttons 2 and 3). Load case HYDROTEST 1,
HYDROTEST 2 or DESIGN sends the report to that block whatever the target phase.

Where the stress ranges land on StressRangeID / StressRangeOD (four angle
columns each, data from row 15):
  Hydrotest 1  D:G      Hydrotest 2  H:K      Design  L:O
  Early undrained   PCD P:S     HCD T:W     FCD AB:AE
  Early drained     PCD AF:AI   HCD AJ:AM   FCD AR:AU
  Middle drained    PCD AV:AY   HCD AZ:BC   FCD BH:BK
  Late drained      PCD CQ:CT   HCD CU:CX   FCD DC:DF
Hydrotest 1 and Hydrotest 2 each have their own cycle count on the INPUT tab
(rows 5 and 6 of the hydrotest columns); the hydrotest damage is their sum.
Button 4 also writes, next to each DBM KP range on the INPUT tab, the first
and last stress-range row inside that range.


COMMON MESSAGES
---------------
  "Step '...' matches N steps"     Use a longer part of the step name or its number.
  "Element set ... not found"      Set ELSET to one of the listed set names.
  "No S11 output at ..."           The ODB has no S output at those section
                                   points; add them to the output request.
  "KP grid of ... differs"         The ODBs do not share the same mesh.
  "... requires the FCD cycle count in B1" (from the workbook)
                                   Fill in that cell in the workbook and rerun step 2.
