FORMAT ALL ODB EXCEL CHARTS - README
====================================

Script
------
format_all_odb_excel_charts.py

Purpose
-------
This standalone script batch-formats native Excel charts that were generated
by the Abaqus ODB postprocessing scripts. It does not open or read an ODB and
does not extract new result data. Run the extraction script first, then run
this formatter on the folder containing the resulting .xlsx workbooks.

The default appearance is based on Cpressformat.png:

* white chart and plot backgrounds
* thin gray chart border and black plot border
* no major or minor gridlines
* solid blue first series, with distinct colors for additional series
* circular data markers and solid connecting lines
* legend at the top
* Arial chart text
* 14-point chart title, 11-point axis titles, 10-point legend, and
  9-point axis labels
* whole-number pipeline-distance labels
* whole-number CPRESS labels with a zero Y-axis minimum
* scientific-notation COPEN labels
* blank cells remain gaps

The script preserves the chart data, cell formulas, chart-title text,
axis-title text, and exact series names such as the original step names.
Existing drawing objects are retained. It does not create new DBM1, DBM2,
and similar manual text annotations because Cpressformat.png does not define
their workbook data coordinates. Those labels can be added later if a mapping
of label text to pipeline distance and Y value is supplied.

Using a manually formatted chart as the template
------------------------------------------------
The script can copy formatting from one native Excel chart that you formatted
manually. The template must be an editable Excel chart, not a pasted picture.
It may be stored in a separate template workbook or in a results workbook.

First list the charts in the template workbook:

python "C:\python_aba\takeoutSFSM\format_all_odb_excel_charts.py" --template-workbook "C:\Data\ChartFormat.xlsx" --list-template-charts

Example listing:

  1: CPRESS Along PIPE Path | lineChart | approximately 15 columns x 27 rows
  2: COPEN Along PIPE Path   | lineChart | approximately 15 columns x 27 rows

Then choose the chart number and apply it to matching target charts:

python "C:\python_aba\takeoutSFSM\format_all_odb_excel_charts.py" --input-dir "C:\Data\Results" --pattern "*.xlsx" --recursive --template-workbook "C:\Data\ChartFormat.xlsx" --template-chart 1 --target-chart-title "*CPRESS*"

The --target-chart-title option is strongly recommended when a workbook
contains several result types. In this example, the CPRESS template is applied
only to charts whose titles contain CPRESS; COPEN and other charts remain
unchanged. The wildcard is case-insensitive. The option may be repeated:

--target-chart-title "*CPRESS*" --target-chart-title "*CONTACT PRESSURE*"

If --target-chart-title is omitted, the template is applied to every chart
with the same native Excel chart type or combination of chart types. Charts
with a different type are left unchanged.

Template mode copies chart and plot backgrounds, borders, built-in chart
style, line and marker formatting, data-label formatting, legend formatting,
axis formatting and scales, fonts, gridlines, gap handling, and chart-wide
options. It preserves each target chart's data formulas, cached values,
series names, chart title text, axis title text, and axis relationships.

Series formatting is matched by series order. If a target chart has more
series than the template, the template series styles repeat. The template
chart size is also used when every chart in a workbook is selected and
compatible. Use --no-resize to keep every target chart's existing size.

The exact template workbook is automatically excluded from the target file
list when it is located under --input-dir.

Requirements and independence
-----------------------------
The script uses only the Python standard library. It is independent of all
other postprocessing scripts and does not require Abaqus, Microsoft Excel,
openpyxl, or pywin32.

Close the target workbooks in Excel before running the script.

Recommended command
-------------------
Use a normal Python installation:

python "C:\python_aba\takeoutSFSM\format_all_odb_excel_charts.py" --input-dir "C:\Data\Results"

The default input filename pattern is:

*_along_path.xlsx

Each source workbook is preserved. A formatted copy is written beside it with
the suffix _formatted, for example:

job1_cpress_copen_along_path.xlsx
job1_cpress_copen_along_path_formatted.xlsx

The run log is:

format_all_odb_excel_charts.log

Command examples
----------------

1. Format the default *_along_path.xlsx files in one folder:

python "C:\python_aba\takeoutSFSM\format_all_odb_excel_charts.py" --input-dir "C:\Data\Results"

2. Include every subfolder:

python "C:\python_aba\takeoutSFSM\format_all_odb_excel_charts.py" --input-dir "C:\Data\Results" --recursive

3. Format every .xlsx workbook instead of only *_along_path.xlsx:

python "C:\python_aba\takeoutSFSM\format_all_odb_excel_charts.py" --input-dir "C:\Data\Results" --pattern "*.xlsx"

4. Supply more than one filename pattern:

python "C:\python_aba\takeoutSFSM\format_all_odb_excel_charts.py" --input-dir "C:\Data\Results" --pattern "*cpress*.xlsx" --pattern "*copen*.xlsx"

5. Preview all source and destination paths without writing files:

python "C:\python_aba\takeoutSFSM\format_all_odb_excel_charts.py" --input-dir "C:\Data\Results" --recursive --dry-run

6. Put all formatted copies under another output folder. Relative subfolders
   are preserved when --recursive is also used:

python "C:\python_aba\takeoutSFSM\format_all_odb_excel_charts.py" --input-dir "C:\Data\Results" --recursive --output-dir "C:\Data\FormattedResults"

7. Replace previously generated _formatted copies:

python "C:\python_aba\takeoutSFSM\format_all_odb_excel_charts.py" --input-dir "C:\Data\Results" --overwrite

8. Use a different output suffix:

python "C:\python_aba\takeoutSFSM\format_all_odb_excel_charts.py" --input-dir "C:\Data\Results" --suffix "_client_format"

9. Keep the existing chart size and position:

python "C:\python_aba\takeoutSFSM\format_all_odb_excel_charts.py" --input-dir "C:\Data\Results" --no-resize

10. Set the approximate chart width and height:

python "C:\python_aba\takeoutSFSM\format_all_odb_excel_charts.py" --input-dir "C:\Data\Results" --chart-width-columns 17 --chart-height-rows 33

11. Change the chart font:

python "C:\python_aba\takeoutSFSM\format_all_odb_excel_charts.py" --input-dir "C:\Data\Results" --font "Calibri"

12. Format source workbooks in place. By default, this saves a backup named
    *_before_chart_format.xlsx first:

python "C:\python_aba\takeoutSFSM\format_all_odb_excel_charts.py" --input-dir "C:\Data\Results" --in-place

13. Format in place without a backup. Use this only when another backup is
    already available:

python "C:\python_aba\takeoutSFSM\format_all_odb_excel_charts.py" --input-dir "C:\Data\Results" --in-place --no-backup

14. Display all options:

python "C:\python_aba\takeoutSFSM\format_all_odb_excel_charts.py" --help

15. List the selectable charts in a manually formatted template workbook:

python "C:\python_aba\takeoutSFSM\format_all_odb_excel_charts.py" --template-workbook "C:\Data\ChartFormat.xlsx" --list-template-charts

16. Apply template chart 1 only to CPRESS charts:

python "C:\python_aba\takeoutSFSM\format_all_odb_excel_charts.py" --input-dir "C:\Data\Results" --pattern "*.xlsx" --recursive --template-workbook "C:\Data\ChartFormat.xlsx" --template-chart 1 --target-chart-title "*CPRESS*"

17. Apply template chart 2 to every compatible chart while retaining each
    target chart's existing size:

python "C:\python_aba\takeoutSFSM\format_all_odb_excel_charts.py" --input-dir "C:\Data\Results" --template-workbook "C:\Data\ChartFormat.xlsx" --template-chart 2 --no-resize

Abaqus Python command
---------------------
The script can also run with Abaqus Python because it uses standard-library
modules only:

abaqus python "C:\python_aba\takeoutSFSM\format_all_odb_excel_charts.py" --input-dir "C:\Data\Results"

Do not use "abaqus viewer noGUI" for this formatter. No Viewer session is
needed.

Important behavior
------------------
* Default mode never overwrites a source workbook.
* Existing formatted outputs cause an error unless --overwrite is supplied.
* --in-place cannot be combined with --output-dir.
* Excel temporary files beginning with ~$ are ignored.
* Formatted-copy files using the selected suffix and backup files named
  *_before_chart_format.xlsx are skipped during discovery.
* A workbook with no native Excel chart is reported as failed and is not
  published as an output.
* A template chart is applied only to matching chart types. Incompatible
  charts remain unchanged and are counted in the console and log.
* --target-chart-title limits template formatting to matching chart titles.
* The --font option controls the built-in formatting mode. In template mode,
  the font comes from the selected template chart.
* If a workbook contains unselected or incompatible charts, automatic resizing
  is suppressed so those charts remain unchanged. Use a separate run or an
  explicit title selection for each chart family when needed.
* The output ZIP structure and every chart XML part are validated before the
  completed workbook replaces its temporary file.

Stopping a run
--------------
Press Ctrl+C once in the Command Prompt or PowerShell window. The current
temporary workbook is removed, completed workbooks remain valid, and the
program exits with code 130. The log records that the run was interrupted.
You can rerun with --overwrite to replace formatted copies already completed
by an earlier run.

Exit codes
----------
0 = every selected workbook was formatted successfully
1 = one or more selected workbooks failed
2 = invalid input or no matching workbook was found
130 = the user terminated the run with Ctrl+C
