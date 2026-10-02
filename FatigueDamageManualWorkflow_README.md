# Manual fatigue damage workflow

Open `FatigueDamageCal_ManualWorkflows_Hydrotest_Design.xlsm` in desktop Microsoft Excel and enable its VBA macros to use the workflow buttons. `FatigueDamageManualWorkflow.bas` is the matching editable VBA source.

## Workflow

1. In **Workflow Controls**, select the input option (AUTO, S11, or DELTA S11), target phase, load case, and relevant report/step settings.
2. Import the report with button 1. For raw S11 input, use buttons 2 and 3 to separate ID/OD and calculate stress ranges.
3. Set life-cycle counts and DBM/RB S-N parameters on the calculation tabs.
4. Use button 4 to calculate fatigue damage, then button 5 to update summaries and charts. The run-all macro performs the workflow together.

Hydrotest and Design Operation each have ID and OD calculation tabs, using fixed total cycle counts of 2 and 1, respectively. Their contributions are shown separately and included in the all-life total damage and UC summaries. Charts show damage and UC versus KP for ID/OD and early-profile/all-life results.

In this completed version, select Hydrotest or Design Operation under **Target phase**; the Load case selection is ignored for those phases. Moving these options into the Load case dropdown was requested later but is not implemented in this version.

The workbook was previously checked using both direct delta-S11 and raw S11 reports, including DBM/RB selection, summaries, and charts. The GitHub upload preserves that completed workbook without recalculating or modifying it.
