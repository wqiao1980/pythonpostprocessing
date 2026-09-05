from __future__ import print_function

"""Extract local-nodal U2 along the complete pipeline path.

This script applies Abaqus' NODAL result transformation, which uses local
coordinate systems defined with *TRANSFORM. Nodes without a nodal transform
remain in the global coordinate system, matching Abaqus/Viewer behavior.

Run headlessly with Abaqus/Viewer. Examples:

    abaqus viewer noGUI=extract_u2_local_selected_steps.py -- --list-steps
    abaqus viewer noGUI=extract_u2_local_selected_steps.py -- --steps Step-2 Step-5
    abaqus viewer noGUI=extract_u2_local_selected_steps.py -- --step-range 3 7

Keep extract_le11_the11_selected_steps.py in the same folder. This script
reuses its tested ODB selection, pipeline-distance, and dependency-free Excel
writing utilities. No Abaqus/CAE window or Microsoft Excel session is needed.
"""

import argparse
import csv
import datetime
import math
import os
import sys
import tempfile
import traceback
import zipfile

from abaqus import session
from abaqusConstants import (
    COMMA_SEPARATED_VALUES,
    COMPONENT,
    NODAL,
    OFF,
)


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

import extract_le11_the11_selected_steps as common


DEFAULT_INSTANCE = "PART-1-1"
DEFAULT_FRAME_INDEX = -1
FIELD_NAME = "U"
COMPONENT_NAME = "U2"


def command_line_arguments():
    """Return only arguments intended for this script after the Abaqus --."""
    arguments = sys.argv[1:]
    if "--" in arguments:
        arguments = arguments[arguments.index("--") + 1 :]
    return arguments


def parse_arguments():
    parser = argparse.ArgumentParser(
        description=(
            "Write local-nodal U2 reports and editable Excel charts along a "
            "complete pipeline path for user-selected ODB steps."
        )
    )
    parser.add_argument(
        "--input-dir",
        default=SCRIPT_DIR,
        help="Directory searched for ODB files (default: script directory).",
    )
    parser.add_argument(
        "--odb",
        action="append",
        default=[],
        help=(
            "Process this ODB only. May be repeated. Relative paths are "
            "resolved against --input-dir."
        ),
    )
    parser.add_argument(
        "--output-dir",
        default=None,
        help="Report and Excel workbook directory (default: --input-dir).",
    )
    parser.add_argument(
        "--output-name",
        default=None,
        help="Report filename for a single --odb run only.",
    )
    parser.add_argument(
        "--instance",
        default=DEFAULT_INSTANCE,
        help="Pipeline instance to process (default: PART-1-1).",
    )
    parser.add_argument(
        "--start-node",
        type=int,
        default=None,
        help="Node label used as path-distance zero. Overrides --start-node-set.",
    )
    parser.add_argument(
        "--start-node-set",
        default="START",
        help=(
            "Instance node set used to locate path-distance zero "
            "(default: START). If absent, an endpoint is selected automatically."
        ),
    )
    parser.add_argument(
        "--frame-index",
        type=int,
        default=DEFAULT_FRAME_INDEX,
        help=(
            "Zero-based frame index. Use -1 for the last frame containing U2 "
            "(default: -1)."
        ),
    )
    parser.add_argument(
        "--list-steps",
        action="store_true",
        help="List each ODB's ordered steps and exit without writing results.",
    )

    selection = parser.add_mutually_exclusive_group()
    selection.add_argument(
        "--steps",
        nargs="+",
        default=None,
        metavar="STEP",
        help="One or more exact step names to include.",
    )
    selection.add_argument(
        "--step-range",
        nargs=2,
        default=None,
        metavar=("START", "END"),
        help=(
            "Inclusive range. Use 1-based step numbers (for example 3 7) or "
            "the exact first and last step names."
        ),
    )

    args = parser.parse_args(command_line_arguments())
    if args.frame_index < -1:
        parser.error("--frame-index must be -1 or zero or greater")
    if args.output_name and len(args.odb) != 1:
        parser.error("--output-name requires exactly one --odb")
    return args


def report_name(odb_path):
    stem = os.path.splitext(os.path.basename(odb_path))[0]
    return stem + "_U2_LOCAL.rpt"


def excel_name(odb_path):
    stem = os.path.splitext(os.path.basename(odb_path))[0]
    return stem + "_U2_LOCAL_along_path.xlsx"


def build_jobs(odb_paths, output_dir, output_name):
    jobs = []
    report_paths_seen = set()
    excel_paths_seen = set()
    for odb_path in odb_paths:
        output_report_name = output_name or report_name(odb_path)
        report_path = os.path.abspath(os.path.join(output_dir, output_report_name))
        excel_path = os.path.abspath(os.path.join(output_dir, excel_name(odb_path)))
        report_key = os.path.normcase(report_path)
        excel_key = os.path.normcase(excel_path)
        if report_key in report_paths_seen:
            raise ValueError(
                "More than one ODB maps to report path: {0}".format(report_path)
            )
        if excel_key in excel_paths_seen:
            raise ValueError(
                "More than one ODB maps to Excel path: {0}".format(excel_path)
            )
        report_paths_seen.add(report_key)
        excel_paths_seen.add(excel_key)
        jobs.append((odb_path, report_path, excel_path))
    return jobs


def open_session_odb(odb_path):
    return session.openOdb(name=odb_path, readOnly=True)


def current_viewport_for_odb(odb):
    session.defaultOdbDisplay.basicOptions.setValues(
        transformationType=NODAL
    )
    viewport_name = session.currentViewportName
    viewport = session.viewports[viewport_name]
    viewport.setValues(displayedObject=odb)
    viewport.odbDisplay.basicOptions.setValues(transformationType=NODAL)
    return viewport


def field_has_u2(frame):
    if FIELD_NAME not in frame.fieldOutputs:
        return False
    components = [
        str(label).upper()
        for label in frame.fieldOutputs[FIELD_NAME].componentLabels
    ]
    return COMPONENT_NAME in components


def select_frames(odb, step_names, frame_index):
    available_steps = list(odb.steps.keys())
    step_positions = dict(
        (step_name, index) for index, step_name in enumerate(available_steps)
    )
    selected = []
    skipped = []
    for step_name in step_names:
        frames = odb.steps[step_name].frames
        if not frames:
            skipped.append((step_name, "the step contains no frames"))
            continue

        if frame_index == -1:
            candidate_indices = range(len(frames) - 1, -1, -1)
        elif frame_index >= len(frames):
            skipped.append(
                (
                    step_name,
                    "frame index {0} is unavailable; the step has {1} frame(s)".format(
                        frame_index, len(frames)
                    ),
                )
            )
            continue
        else:
            candidate_indices = (frame_index,)

        selection = None
        for candidate_index in candidate_indices:
            frame = frames[candidate_index]
            if field_has_u2(frame):
                selection = (
                    step_name,
                    step_positions[step_name],
                    candidate_index,
                    frame,
                )
                break

        if selection is None:
            if frame_index == -1:
                reason = "no frame contains displacement component U2"
            else:
                reason = "frame {0} does not contain displacement component U2".format(
                    frame_index
                )
            skipped.append((step_name, reason))
        else:
            selected.append(selection)
    return selected, skipped


def normalized_header(value):
    return "".join(
        character for character in str(value).upper() if character.isalnum()
    )


def header_indices(record):
    normalized = [normalized_header(value) for value in record]
    node_index = None
    value_index = None
    instance_index = None
    for index, value in enumerate(normalized):
        if value in ("NODELABEL", "NODE"):
            node_index = index
        elif value in ("INSTANCE", "INSTANCENAME", "PARTINSTANCE"):
            instance_index = index
        elif value == "U2" or value.endswith("UU2"):
            value_index = index
    if node_index is None or value_index is None:
        return None
    return instance_index, node_index, value_index


def open_csv_report(report_path):
    if sys.version_info[0] < 3:
        return open(report_path, "rb")
    return open(report_path, "r", newline="")


def parse_local_u2_report(report_path, instance_name, allowed_nodes):
    values_by_node = {}
    indices = None
    current_instance = None
    report_file = open_csv_report(report_path)
    try:
        for record in csv.reader(report_file):
            possible_header = header_indices(record)
            if possible_header is not None:
                indices = possible_header
                current_instance = None
                continue
            if indices is None:
                continue

            instance_index, node_index, value_index = indices
            required_index = max(
                node_index,
                value_index,
                instance_index if instance_index is not None else 0,
            )
            if len(record) <= required_index:
                continue

            if instance_index is not None:
                instance_text = record[instance_index].strip()
                if instance_text:
                    current_instance = instance_text
                if (
                    current_instance is not None
                    and current_instance.upper() != instance_name.upper()
                ):
                    continue

            try:
                node_label = int(float(record[node_index].strip()))
                u2_value = float(record[value_index].strip())
            except (TypeError, ValueError):
                continue
            if node_label in allowed_nodes:
                values_by_node[node_label] = u2_value
    finally:
        report_file.close()
    return values_by_node


def temporary_report_path():
    descriptor, report_path = tempfile.mkstemp(
        prefix="abaqus_u2_local_", suffix=".csv"
    )
    os.close(descriptor)
    os.remove(report_path)
    return report_path


def extract_local_u2(
    viewport,
    odb,
    instance_name,
    allowed_nodes,
    step_position,
    frame_index,
):
    """Use Viewer NODAL transformation and parse its CSV field report."""
    viewport.odbDisplay.setFrame(step=step_position, frame=frame_index)
    viewport.odbDisplay.setPrimaryVariable(
        variableLabel=FIELD_NAME,
        outputPosition=NODAL,
        refinement=(COMPONENT, COMPONENT_NAME),
    )
    viewport.odbDisplay.basicOptions.setValues(transformationType=NODAL)
    session.fieldReportOptions.setValues(
        reportFormat=COMMA_SEPARATED_VALUES,
        printTotal=OFF,
        printMinMax=OFF,
    )

    report_path = temporary_report_path()
    try:
        session.writeFieldReport(
            fileName=report_path,
            append=OFF,
            sortItem="Node Label",
            odb=odb,
            step=step_position,
            frame=frame_index,
            outputPosition=NODAL,
            displayGroup=viewport.odbDisplay.displayGroup,
            variable=((FIELD_NAME, NODAL, ((COMPONENT, COMPONENT_NAME),)),),
        )
        values_by_node = parse_local_u2_report(
            report_path, instance_name, allowed_nodes
        )
    finally:
        if os.path.isfile(report_path):
            os.remove(report_path)

    if not values_by_node:
        raise ValueError(
            "Abaqus wrote no parsable local U2 values for instance '{0}'. "
            "Confirm that U2 field output exists and the current display group "
            "contains the pipeline instance.".format(instance_name)
        )
    return values_by_node


def ordered_path_nodes(path_information):
    return sorted(
        path_information,
        key=lambda node_label: (
            path_information[node_label][0],
            path_information[node_label][1],
            node_label,
        ),
    )


def write_frame_table(
    report_file, path_information, coordinates, values_by_node
):
    header = "{0:>8}{1:>16}{2:>12}{3:>16}{4:>16}{5:>16}{6:>18}".format(
        "Path ID",
        "Path Distance",
        "Node Label",
        "X",
        "Y",
        "Z",
        "U2 LOCAL",
    )
    report_file.write(header + "\n")
    report_file.write("-" * len(header) + "\n")
    matching_count = 0
    for node_label in ordered_path_nodes(path_information):
        if node_label not in values_by_node:
            continue
        path_id, path_distance = path_information[node_label]
        x_coordinate, y_coordinate, z_coordinate = coordinates[node_label]
        report_file.write(
            "{0:>8d}{1:>16}{2:>12d}{3:>16}{4:>16}{5:>16}{6:>18}\n".format(
                path_id,
                common.engineering_format(path_distance),
                node_label,
                common.engineering_format(x_coordinate),
                common.engineering_format(y_coordinate),
                common.engineering_format(z_coordinate),
                common.engineering_format(values_by_node[node_label]),
            )
        )
        matching_count += 1
    report_file.write("\n")
    return matching_count


def local_u2_curves(path_information, extracted_frames):
    curves_by_path = {}
    for step_name, unused_frame_index, unused_frame, values_by_node in extracted_frames:
        points_by_path = {}
        for node_label, u2_value in values_by_node.items():
            if node_label not in path_information:
                continue
            path_id, path_distance = path_information[node_label]
            points_by_path.setdefault(path_id, []).append(
                (path_distance, u2_value, node_label)
            )
        for path_id, points in points_by_path.items():
            points.sort(key=lambda item: (item[0], item[2]))
            curves_by_path.setdefault(path_id, []).append((step_name, points))
    return curves_by_path


def xlsx_path_table(curves):
    locations = set()
    curve_maps = []
    for step_name, points in curves:
        values = {}
        for path_distance, u2_value, node_label in points:
            location = (path_distance, node_label)
            locations.add(location)
            values[location] = u2_value
        curve_maps.append((step_name, values))
    return sorted(locations, key=lambda item: (item[0], item[1])), curve_maps


def xlsx_worksheet_xml(odb_name, path_id, locations, curve_maps):
    last_column = 2 + len(curve_maps)
    last_row = 4 + len(locations)
    rows = [
        '<row r="1" ht="22" customHeight="1">{0}</row>'.format(
            common.xlsx_inline_cell(
                1,
                1,
                "Local Nodal U2 Along Pipeline Path - {0} - Path {1}".format(
                    odb_name, path_id
                ),
                1,
            )
        ),
        '<row r="2">{0}</row>'.format(
            common.xlsx_inline_cell(
                2,
                1,
                "Abaqus NODAL transformation; nodes without *TRANSFORM use global U2",
                2,
            )
        ),
    ]
    header_cells = [
        common.xlsx_inline_cell(4, 1, "Path Distance", 3),
        common.xlsx_inline_cell(4, 2, "Node Label", 3),
    ]
    for column_number, (step_name, unused_values) in enumerate(curve_maps, 3):
        header_cells.append(
            common.xlsx_inline_cell(4, column_number, step_name, 3)
        )
    rows.append(
        '<row r="4" ht="30" customHeight="1">{0}</row>'.format(
            "".join(header_cells)
        )
    )

    for row_number, location in enumerate(locations, 5):
        path_distance, node_label = location
        cells = [
            common.xlsx_number_cell(row_number, 1, path_distance, 4),
            common.xlsx_number_cell(row_number, 2, node_label, 0),
        ]
        for column_number, (unused_step, values) in enumerate(curve_maps, 3):
            cells.append(
                common.xlsx_number_cell(
                    row_number, column_number, values.get(location), 5
                )
            )
        rows.append('<row r="{0}">{1}</row>'.format(row_number, "".join(cells)))

    columns = [
        '<col min="1" max="1" width="16" customWidth="1"/>',
        '<col min="2" max="2" width="13" customWidth="1"/>',
    ]
    if last_column >= 3:
        columns.append(
            '<col min="3" max="{0}" width="18" customWidth="1"/>'.format(
                last_column
            )
        )
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:{0}{1}"/>
  <sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>{2}</cols>
  <sheetData>{3}</sheetData>
  <autoFilter ref="A4:{0}{1}"/>
  <drawing r:id="rId1"/>
</worksheet>""".format(
        common.excel_column_name(last_column),
        last_row,
        "".join(columns),
        "".join(rows),
    )


def xlsx_drawing_xml(chart_index, start_column):
    end_column = start_column + 11
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>{0}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>{1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>23</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Local U2 Chart {2}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>""".format(start_column, end_column, chart_index)


def xlsx_chart_xml(
    odb_name,
    path_id,
    sheet_name,
    locations,
    curve_maps,
    chart_index,
):
    data_start_row = 5
    data_end_row = 4 + len(locations)
    series_xml = []
    for series_index, (step_name, values) in enumerate(curve_maps):
        series_xml.append(
            common.xlsx_chart_series_xml(
                series_index,
                sheet_name,
                4,
                data_start_row,
                data_end_row,
                3 + series_index,
                step_name,
                locations,
                values,
            )
        )
    x_axis_id = 58650112 + chart_index * 2
    y_axis_id = x_axis_id + 1
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/><c:style val="10"/>
  <c:chart>
    <c:title>{0}<c:layout/><c:overlay val="0"/></c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea><c:layout/>
      <c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>{1}<c:axId val="{2}"/><c:axId val="{3}"/></c:scatterChart>
      {4}
      {5}
    </c:plotArea>
    <c:legend><c:legendPos val="r"/><c:layout/><c:overlay val="0"/></c:legend>
    <c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/>
  </c:chart>
  <c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings>
</c:chartSpace>""".format(
        common.xlsx_chart_text(
            "Local Nodal U2 Along Pipeline Path - {0} - Path {1}".format(
                odb_name, path_id
            ),
            1200,
        ),
        "".join(series_xml),
        x_axis_id,
        y_axis_id,
        common.xlsx_value_axis_xml(
            x_axis_id, y_axis_id, "b", "Path Distance", "0.0000"
        ),
        common.xlsx_value_axis_xml(
            y_axis_id,
            x_axis_id,
            "l",
            "Local Nodal U2 (model length units)",
            "0.000000E+00",
        ),
    )


def write_u2_excel(excel_path, odb_name, path_information, extracted_frames):
    curves_by_path = local_u2_curves(path_information, extracted_frames)
    if not curves_by_path:
        raise ValueError("No local U2 path data are available for plotting.")

    used_names = set()
    sheet_data = []
    for path_id in sorted(curves_by_path):
        curves = curves_by_path[path_id]
        locations, curve_maps = xlsx_path_table(curves)
        sheet_data.append(
            (
                path_id,
                common.xlsx_sheet_name(path_id, used_names),
                locations,
                curve_maps,
            )
        )
    sheet_names = [item[1] for item in sheet_data]

    workbook_zip = zipfile.ZipFile(
        excel_path, "w", compression=zipfile.ZIP_DEFLATED, allowZip64=True
    )
    try:
        workbook_zip.writestr(
            "[Content_Types].xml",
            common.xlsx_bytes(common.xlsx_content_types_xml(len(sheet_data))),
        )
        workbook_zip.writestr(
            "_rels/.rels",
            common.xlsx_bytes(common.xlsx_root_relationships_xml()),
        )
        workbook_zip.writestr(
            "docProps/core.xml",
            common.xlsx_bytes(common.xlsx_core_properties_xml()),
        )
        workbook_zip.writestr(
            "docProps/app.xml",
            common.xlsx_bytes(common.xlsx_app_properties_xml(sheet_names)),
        )
        workbook_zip.writestr(
            "xl/workbook.xml",
            common.xlsx_bytes(common.xlsx_workbook_xml(sheet_names)),
        )
        workbook_zip.writestr(
            "xl/_rels/workbook.xml.rels",
            common.xlsx_bytes(
                common.xlsx_workbook_relationships_xml(len(sheet_data))
            ),
        )
        workbook_zip.writestr(
            "xl/styles.xml", common.xlsx_bytes(common.xlsx_styles_xml())
        )

        for sheet_index, item in enumerate(sheet_data, 1):
            path_id, sheet_name, locations, curve_maps = item
            start_column = max(5, 3 + len(curve_maps))
            workbook_zip.writestr(
                "xl/worksheets/sheet{0}.xml".format(sheet_index),
                common.xlsx_bytes(
                    xlsx_worksheet_xml(
                        odb_name, path_id, locations, curve_maps
                    )
                ),
            )
            workbook_zip.writestr(
                "xl/worksheets/_rels/sheet{0}.xml.rels".format(sheet_index),
                common.xlsx_bytes(
                    common.xlsx_worksheet_relationships_xml(sheet_index)
                ),
            )
            workbook_zip.writestr(
                "xl/drawings/drawing{0}.xml".format(sheet_index),
                common.xlsx_bytes(
                    xlsx_drawing_xml(sheet_index, start_column)
                ),
            )
            workbook_zip.writestr(
                "xl/drawings/_rels/drawing{0}.xml.rels".format(sheet_index),
                common.xlsx_bytes(
                    common.xlsx_drawing_relationships_xml(sheet_index)
                ),
            )
            workbook_zip.writestr(
                "xl/charts/chart{0}.xml".format(sheet_index),
                common.xlsx_bytes(
                    xlsx_chart_xml(
                        odb_name,
                        path_id,
                        sheet_name,
                        locations,
                        curve_maps,
                        sheet_index,
                    )
                ),
            )
    finally:
        workbook_zip.close()


def write_report(
    odb_path,
    report_path,
    excel_path,
    instance_name,
    requested_steps,
    step_range,
    frame_index,
    start_node,
    start_node_set,
):
    odb = None
    try:
        print("Opening: {0}".format(odb_path))
        odb = open_session_odb(odb_path)
        instance_key = common.repository_key(
            odb.rootAssembly.instances, instance_name
        )
        if instance_key is None:
            raise ValueError(
                "Instance '{0}' was not found. Available instances: {1}".format(
                    instance_name,
                    ", ".join(odb.rootAssembly.instances.keys()),
                )
            )
        instance = odb.rootAssembly.instances[instance_key]
        path_information, coordinates, path_origins, path_notes = (
            common.build_pipeline_paths(
                instance, start_node, start_node_set
            )
        )

        available_steps = list(odb.steps.keys())
        selected_steps, selection_notes = common.select_steps(
            available_steps, requested_steps, step_range
        )
        if not selected_steps:
            details = "; ".join(selection_notes) if selection_notes else "none selected"
            raise ValueError("No steps were selected: {0}".format(details))

        selected_frames, skipped_steps = select_frames(
            odb, selected_steps, frame_index
        )
        if not selected_frames:
            details = "; ".join(
                "{0}: {1}".format(step_name, reason)
                for step_name, reason in skipped_steps
            )
            raise ValueError(
                "None of the selected steps contains usable U2 field output: {0}".format(
                    details
                )
            )

        viewport = current_viewport_for_odb(odb)
        allowed_nodes = set(path_information.keys())
        extracted_frames = []
        for step_name, step_position, selected_frame_index, frame in selected_frames:
            values_by_node = extract_local_u2(
                viewport,
                odb,
                instance_key,
                allowed_nodes,
                step_position,
                selected_frame_index,
            )
            extracted_frames.append(
                (step_name, selected_frame_index, frame, values_by_node)
            )

        with open(report_path, "w") as report_file:
            report_file.write("*" * 90 + "\n")
            report_file.write(
                "Local Nodal U2 Pipeline Report, written {0}\n\n".format(
                    datetime.datetime.now().strftime("%a %b %d %H:%M:%S %Y")
                )
            )
            report_file.write("ODB: {0}\n".format(odb_path.replace("\\", "/")))
            report_file.write("Pipeline instance: {0}\n".format(instance_key))
            report_file.write(
                "Transformation: Abaqus NODAL (*TRANSFORM local directions)\n"
            )
            report_file.write(
                "Fallback: nodes without *TRANSFORM are reported in global directions\n"
            )
            report_file.write(
                "Path distance: cumulative distance along undeformed line-element mesh\n"
            )
            report_file.write("Path origins:")
            for path_id, origin_node in path_origins:
                report_file.write(
                    " Path {0}=node {1};".format(path_id, origin_node)
                )
            report_file.write("\n\n")

            for step_name, selected_frame_index, frame, values_by_node in extracted_frames:
                report_file.write("Source 1\n")
                report_file.write("---------\n\n")
                report_file.write("   Step: {0}\n".format(step_name))
                report_file.write(
                    "   Frame index: {0}\n".format(selected_frame_index)
                )
                report_file.write("   Frame: {0}\n".format(frame.description))
                report_file.write("   Output position: nodal\n")
                matching_count = len(
                    set(values_by_node.keys()).intersection(allowed_nodes)
                )
                report_file.write(
                    "   Matching pipeline nodes: {0}\n".format(matching_count)
                )
                omitted_count = len(allowed_nodes) - matching_count
                if omitted_count:
                    report_file.write(
                        "   Pipeline nodes without U2 omitted: {0}\n".format(
                            omitted_count
                        )
                    )
                report_file.write("\n")
                write_frame_table(
                    report_file,
                    path_information,
                    coordinates,
                    values_by_node,
                )

        odb_stem = os.path.splitext(os.path.basename(odb_path))[0]
        write_u2_excel(
            excel_path, odb_stem, path_information, extracted_frames
        )

        print("Selected steps:")
        for step_name, unused_step_position, selected_frame_index, unused_frame in selected_frames:
            print("  {0} (frame index {1})".format(step_name, selected_frame_index))
        for note in selection_notes:
            print("Note: {0}.".format(note))
        for skipped_step, reason in skipped_steps:
            print("Skipped step '{0}': {1}.".format(skipped_step, reason))
        for note in path_notes:
            print("Path note: {0}.".format(note))
        print("Wrote:      {0}".format(report_path))
        print("Excel plot: {0}".format(excel_path))
    finally:
        if odb is not None:
            odb.close()


def print_steps(odb_path):
    odb = None
    try:
        odb = open_session_odb(odb_path)
        print(odb_path)
        step_names = list(odb.steps.keys())
        if not step_names:
            print("  (no analysis steps)")
        for index, step_name in enumerate(step_names, 1):
            print("  {0:>3}: {1}".format(index, step_name))
    finally:
        if odb is not None:
            odb.close()


def main():
    args = parse_arguments()
    input_dir = os.path.abspath(args.input_dir)
    odb_paths = common.find_odb_files(input_dir, args.odb)

    if args.list_steps:
        for odb_path in odb_paths:
            print_steps(odb_path)
        return 0

    output_dir = os.path.abspath(args.output_dir or input_dir)
    if not os.path.isdir(output_dir):
        os.makedirs(output_dir)
    jobs = build_jobs(odb_paths, output_dir, args.output_name)

    failures = []
    for odb_path, report_path, excel_path in jobs:
        try:
            write_report(
                odb_path=odb_path,
                report_path=report_path,
                excel_path=excel_path,
                instance_name=args.instance,
                requested_steps=args.steps,
                step_range=args.step_range,
                frame_index=args.frame_index,
                start_node=args.start_node,
                start_node_set=args.start_node_set,
            )
        except Exception as exc:
            failures.append((odb_path, str(exc)))
            print("FAILED:  {0}".format(odb_path))
            print("         {0}".format(exc))
            traceback.print_exc()

    print(
        "Completed: {0} succeeded, {1} failed.".format(
            len(jobs) - len(failures), len(failures)
        )
    )
    if failures:
        print("Failed ODB files:")
        for odb_path, message in failures:
            print("  {0}: {1}".format(odb_path, message))
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
