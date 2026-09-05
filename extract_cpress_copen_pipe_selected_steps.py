from __future__ import print_function

"""Extract CPRESS and COPEN along PIPE* elements of a pipeline.

Run with the Abaqus Python interpreter. Examples:

    abaqus python extract_cpress_copen_pipe_selected_steps.py --list-steps
    abaqus python extract_cpress_copen_pipe_selected_steps.py --steps Step-2 Step-5
    abaqus python extract_cpress_copen_pipe_selected_steps.py --step-range 3 7

Keep extract_le11_the11_selected_steps_pipe_only.py in the same folder. This
script reuses its tested ODB selection, PIPE-only path-distance, and
dependency-free Excel-writing utilities.

At a path node with multiple contact values, the script reports the maximum
CPRESS and minimum COPEN. Missing contact output is left blank, not replaced
with zero. The Excel workbook contains native, user-editable scatter charts.
"""

import argparse
import datetime
import math
import os
import sys
import traceback
import zipfile

from odbAccess import openOdb


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

import extract_le11_the11_selected_steps_pipe_only as common


DEFAULT_INSTANCE = "PART-1-1"
DEFAULT_FRAME_INDEX = -1
FIELD_NAMES = ("CPRESS", "COPEN")
SCRIPT_VERSION = "2026-09-05-r3"


def parse_arguments():
    parser = argparse.ArgumentParser(
        description=(
            "Write PIPE-only CPRESS/COPEN reports and an editable Excel "
            "workbook for selected ODB steps."
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
        "--element-set",
        action="append",
        default=[],
        metavar="SET",
        help=(
            "Restrict output to this element set. May be repeated. "
            "Instance- and assembly-level element sets are supported."
        ),
    )
    parser.add_argument(
        "--element-range",
        action="append",
        nargs=2,
        type=int,
        default=[],
        metavar=("FIRST", "LAST"),
        help=(
            "Restrict output to this inclusive element-label range. May be "
            "repeated. Sets and ranges are combined as a union."
        ),
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
            "Zero-based frame index. Use -1 for the last frame containing "
            "both CPRESS and COPEN (default: -1)."
        ),
    )
    parser.add_argument(
        "--list-steps",
        action="store_true",
        help="List each ODB's ordered steps and exit without writing results.",
    )
    parser.add_argument(
        "--list-element-sets",
        action="store_true",
        help=(
            "List element sets and their PIPE-element counts, then exit "
            "without writing results."
        ),
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
            "Inclusive range. Use 1-based step numbers (for example 3 7) "
            "or the exact first and last step names."
        ),
    )

    args = parser.parse_args()
    if args.frame_index < -1:
        parser.error("--frame-index must be -1 or zero or greater")
    if args.output_name and len(args.odb) != 1:
        parser.error("--output-name requires exactly one --odb")
    for first_label, last_label in args.element_range:
        if first_label > last_label:
            parser.error(
                "--element-range FIRST must be less than or equal to LAST"
            )
    return args


def report_name(odb_path):
    stem = os.path.splitext(os.path.basename(odb_path))[0]
    return stem + "_PIPE_ONLY_CPRESS_COPEN.rpt"


def excel_name(odb_path):
    stem = os.path.splitext(os.path.basename(odb_path))[0]
    return stem + "_PIPE_ONLY_CPRESS_COPEN_along_path.xlsx"


def build_jobs(odb_paths, output_dir, output_name):
    jobs = []
    report_paths_seen = set()
    excel_paths_seen = set()
    for odb_path in odb_paths:
        output_report_name = output_name or report_name(odb_path)
        report_path = os.path.abspath(
            os.path.join(output_dir, output_report_name)
        )
        excel_path = os.path.abspath(
            os.path.join(output_dir, excel_name(odb_path))
        )
        report_key = os.path.normcase(report_path)
        excel_key = os.path.normcase(excel_path)
        if report_key in report_paths_seen:
            raise ValueError(
                "More than one ODB maps to report path: {0}".format(
                    report_path
                )
            )
        if excel_key in excel_paths_seen:
            raise ValueError(
                "More than one ODB maps to Excel path: {0}".format(excel_path)
            )
        report_paths_seen.add(report_key)
        excel_paths_seen.add(excel_key)
        jobs.append((odb_path, report_path, excel_path))
    return jobs


def flattened_element_labels(container):
    """Yield labels, using Abaqus' bulk array accessor when available."""
    if container is None:
        return
    if hasattr(container, "label") and hasattr(container, "connectivity"):
        yield int(container.label)
        return

    get_members = getattr(container, "getMemberFromAll", None)
    if get_members is not None:
        try:
            labels = get_members("label")
            for label in labels:
                yield int(label)
            return
        except Exception:
            pass

    try:
        members = iter(container)
    except TypeError:
        return
    for member in members:
        if hasattr(member, "label") and hasattr(member, "connectivity"):
            yield int(member.label)
        else:
            for label in flattened_element_labels(member):
                yield label


def element_set_labels_for_instance(
    element_set, instance, instance_labels=None
):
    """Return labels belonging to the requested instance from an OdbSet."""
    if instance_labels is None:
        instance_labels = set(
            int(label)
            for label in flattened_element_labels(instance.elements)
        )
    try:
        instance_names = tuple(element_set.instanceNames)
    except Exception:
        instance_names = ()

    if instance_names:
        matching_indices = [
            index
            for index, name in enumerate(instance_names)
            if str(name).upper() == instance.name.upper()
        ]
        if not matching_indices:
            return set()

        elements_member = element_set.elements
        try:
            outer_count = len(elements_member)
            first_member = elements_member[0] if outer_count else None
        except (TypeError, AttributeError, IndexError):
            outer_count = 0
            first_member = None

        # Assembly-level sets normally store one element array per instance.
        if (
            outer_count == len(instance_names)
            and outer_count
            and not hasattr(first_member, "label")
        ):
            selected_containers = [
                elements_member[index] for index in matching_indices
            ]
        else:
            # A one-instance assembly set can be exposed directly as an
            # OdbMeshElementArray rather than a tuple of arrays.
            selected_containers = [elements_member]
    else:
        selected_containers = [element_set.elements]

    labels = set()
    for selected_container in selected_containers:
        for label in flattened_element_labels(selected_container):
            if label in instance_labels:
                labels.add(label)
    return labels


def resolve_element_set(
    odb, instance, requested_name, instance_labels=None
):
    """Find a case-insensitive instance- or assembly-level element set."""
    instance_key = common.repository_key(
        instance.elementSets, requested_name
    )
    if instance_key is not None:
        element_set = instance.elementSets[instance_key]
        return (
            element_set_labels_for_instance(
                element_set, instance, instance_labels
            ),
            "instance set '{0}'".format(instance_key),
        )

    assembly_key = common.repository_key(
        odb.rootAssembly.elementSets, requested_name
    )
    if assembly_key is not None:
        element_set = odb.rootAssembly.elementSets[assembly_key]
        return (
            element_set_labels_for_instance(
                element_set, instance, instance_labels
            ),
            "assembly set '{0}'".format(assembly_key),
        )

    available = sorted(
        set(
            list(instance.elementSets.keys())
            + list(odb.rootAssembly.elementSets.keys())
        )
    )
    raise ValueError(
        "Element set '{0}' was not found. Available element sets: "
        "{1}".format(
            requested_name, ", ".join(available) if available else "(none)"
        )
    )


def element_selection_description(requested_sets, requested_ranges):
    parts = []
    if requested_sets:
        parts.append(
            "sets {0}".format(
                ", ".join("'{0}'".format(name) for name in requested_sets)
            )
        )
    if requested_ranges:
        parts.append(
            "ranges {0}".format(
                ", ".join(
                    "{0}-{1}".format(first_label, last_label)
                    for first_label, last_label in requested_ranges
                )
            )
        )
    return "; ".join(parts) if parts else "all PIPE* elements"


def resolve_pipe_element_selection(
    odb, instance, requested_sets, requested_ranges
):
    """Return selected PIPE labels/nodes; set and range requests form a union."""
    instance_labels = set()
    pipe_elements = {}
    for element in instance.elements:
        label = int(element.label)
        instance_labels.add(label)
        if str(element.type).upper().startswith("PIPE"):
            pipe_elements[label] = element
    if not pipe_elements:
        raise ValueError(
            "Instance '{0}' contains no PIPE* elements.".format(instance.name)
        )

    selection_requested = bool(requested_sets or requested_ranges)
    selected_labels = set()
    notes = []

    for requested_name in requested_sets:
        set_labels, resolved_name = resolve_element_set(
            odb, instance, requested_name, instance_labels
        )
        matching_pipe_labels = set_labels.intersection(pipe_elements)
        selected_labels.update(matching_pipe_labels)
        notes.append(
            "{0}: selected {1} PIPE element(s) from {2} element(s)".format(
                resolved_name,
                len(matching_pipe_labels),
                len(set_labels),
            )
        )

    for first_label, last_label in requested_ranges:
        matching_pipe_labels = set(
            label
            for label in pipe_elements
            if first_label <= label <= last_label
        )
        selected_labels.update(matching_pipe_labels)
        notes.append(
            "element range {0}-{1}: selected {2} PIPE element(s)".format(
                first_label, last_label, len(matching_pipe_labels)
            )
        )

    if not selection_requested:
        selected_labels = set(pipe_elements.keys())
        notes.append(
            "no element filter requested; all {0} PIPE element(s) selected".format(
                len(selected_labels)
            )
        )
    elif not selected_labels:
        raise ValueError(
            "The requested element sets/ranges contain no PIPE* elements "
            "in instance '{0}'.".format(instance.name)
        )

    selected_nodes = set()
    for label in selected_labels:
        selected_nodes.update(
            int(node_label)
            for node_label in pipe_elements[label].connectivity
        )

    return (
        selected_labels,
        selected_nodes,
        element_selection_description(requested_sets, requested_ranges),
        notes,
        len(pipe_elements),
    )


def field_name_matches(candidate, base_name):
    text = str(candidate).strip().upper()
    base = base_name.upper()
    if text == base:
        return True
    if not text.startswith(base):
        return False
    if len(text) == len(base):
        return True
    return not text[len(base)].isalnum()


def contact_field_outputs(frame, base_name):
    outputs = []
    for key in frame.fieldOutputs.keys():
        field_output = frame.fieldOutputs[key]
        candidates = [key]
        try:
            candidates.append(field_output.name)
        except Exception:
            pass
        if any(field_name_matches(candidate, base_name) for candidate in candidates):
            outputs.append((str(key), field_output))
    return outputs


def frame_has_required_fields(frame):
    return all(contact_field_outputs(frame, name) for name in FIELD_NAMES)


def select_frames(odb, selected_steps, frame_index):
    selected = []
    skipped = []
    for step_name in selected_steps:
        frames = odb.steps[step_name].frames
        if not frames:
            skipped.append((step_name, "the step has no frames"))
            continue

        if frame_index == -1:
            candidate_indices = range(len(frames) - 1, -1, -1)
        elif frame_index >= len(frames):
            skipped.append(
                (
                    step_name,
                    "frame index {0} is unavailable; the step has {1} "
                    "frame(s)".format(frame_index, len(frames)),
                )
            )
            continue
        else:
            candidate_indices = (frame_index,)

        selection = None
        for candidate_index in candidate_indices:
            frame = frames[candidate_index]
            if frame_has_required_fields(frame):
                selection = (step_name, candidate_index, frame)
                break

        if selection is None:
            if frame_index == -1:
                reason = "no frame contains both CPRESS and COPEN"
            else:
                available = ", ".join(
                    sorted(str(key) for key in frames[frame_index].fieldOutputs.keys())
                )
                reason = (
                    "frame {0} does not contain both CPRESS and COPEN; "
                    "available fields: {1}".format(frame_index, available)
                )
            skipped.append((step_name, reason))
        else:
            selected.append(selection)
    return selected, skipped


def value_instance_matches(value, instance_name):
    try:
        value_instance = value.instance
    except Exception:
        value_instance = None
    if value_instance is None:
        return True
    try:
        return value_instance.name.upper() == instance_name.upper()
    except Exception:
        return True


def optional_element_label(value):
    try:
        return int(value.elementLabel)
    except Exception:
        return None


def scalar_value(value):
    data = common.field_value_data(value)
    if not data:
        raise ValueError("A contact field value contains no scalar data.")
    scalar = float(data[0])
    if math.isnan(scalar) or math.isinf(scalar):
        return None
    return scalar


def aggregate_contact_values(
    frame,
    base_name,
    instance_name,
    allowed_nodes,
    pipe_element_labels,
):
    values_by_node = {}
    counts_by_node = {}
    source_names = []
    reduce_with_maximum = base_name.upper() == "CPRESS"

    for source_name, field_output in contact_field_outputs(frame, base_name):
        source_names.append(source_name)
        for value in field_output.values:
            if not value_instance_matches(value, instance_name):
                continue
            try:
                node_label = int(value.nodeLabel)
            except Exception:
                continue
            if node_label not in allowed_nodes:
                continue

            # ELEMENT_NODAL contact data can identify its source element.
            # NODAL contact data do not carry an element label; those values
            # are retained when their node belongs to the PIPE path.
            element_label = optional_element_label(value)
            if (
                element_label not in (None, 0)
                and element_label not in pipe_element_labels
            ):
                continue

            contact_value = scalar_value(value)
            if contact_value is None:
                continue
            counts_by_node[node_label] = counts_by_node.get(node_label, 0) + 1
            if node_label not in values_by_node:
                values_by_node[node_label] = contact_value
            elif reduce_with_maximum:
                values_by_node[node_label] = max(
                    values_by_node[node_label], contact_value
                )
            else:
                values_by_node[node_label] = min(
                    values_by_node[node_label], contact_value
                )

    return values_by_node, counts_by_node, sorted(set(source_names))


def ordered_path_nodes(path_information):
    return sorted(
        path_information,
        key=lambda node_label: (
            path_information[node_label][0],
            path_information[node_label][1],
            node_label,
        ),
    )


def optional_engineering(value):
    if value is None:
        return ""
    return common.engineering_format(value)


def write_frame_table(
    report_file,
    path_information,
    coordinates,
    cpress_values,
    copen_values,
    cpress_counts,
    copen_counts,
):
    header = (
        "{0:>8}{1:>16}{2:>12}{3:>16}{4:>16}{5:>16}"
        "{6:>18}{7:>18}{8:>12}{9:>12}"
    ).format(
        "Path ID",
        "Path Distance",
        "Node Label",
        "X",
        "Y",
        "Z",
        "CPRESS MAX",
        "COPEN MIN",
        "P Samples",
        "O Samples",
    )
    report_file.write(header + "\n")
    report_file.write("-" * len(header) + "\n")

    for node_label in ordered_path_nodes(path_information):
        path_id, path_distance = path_information[node_label]
        x_coordinate, y_coordinate, z_coordinate = coordinates[node_label]
        report_file.write(
            (
                "{0:>8d}{1:>16}{2:>12d}{3:>16}{4:>16}{5:>16}"
                "{6:>18}{7:>18}{8:>12d}{9:>12d}\n"
            ).format(
                path_id,
                common.engineering_format(path_distance),
                node_label,
                common.engineering_format(x_coordinate),
                common.engineering_format(y_coordinate),
                common.engineering_format(z_coordinate),
                optional_engineering(cpress_values.get(node_label)),
                optional_engineering(copen_values.get(node_label)),
                cpress_counts.get(node_label, 0),
                copen_counts.get(node_label, 0),
            )
        )
    report_file.write("\n")


def write_frame_summary(
    report_file, path_information, cpress_values, copen_values
):
    if cpress_values:
        node_label = max(cpress_values, key=lambda label: cpress_values[label])
        path_id, path_distance = path_information[node_label]
        report_file.write(
            "Maximum CPRESS: {0} at path {1}, distance {2}, node {3}\n".format(
                common.engineering_format(cpress_values[node_label]),
                path_id,
                common.engineering_format(path_distance),
                node_label,
            )
        )
    else:
        report_file.write("Maximum CPRESS: no PIPE-path value available\n")

    if copen_values:
        node_label = min(copen_values, key=lambda label: copen_values[label])
        path_id, path_distance = path_information[node_label]
        report_file.write(
            "Minimum COPEN: {0} at path {1}, distance {2}, node {3}\n\n".format(
                common.engineering_format(copen_values[node_label]),
                path_id,
                common.engineering_format(path_distance),
                node_label,
            )
        )
    else:
        report_file.write("Minimum COPEN: no PIPE-path value available\n\n")


def contact_curves(path_information, extracted_frames, variable_name):
    curves_by_path = {}
    for (
        step_name,
        unused_frame_index,
        unused_frame,
        cpress_values,
        copen_values,
        unused_cpress_counts,
        unused_copen_counts,
        unused_cpress_sources,
        unused_copen_sources,
    ) in extracted_frames:
        values_by_node = (
            cpress_values if variable_name == "CPRESS" else copen_values
        )
        points_by_path = {}
        for node_label in ordered_path_nodes(path_information):
            path_id, path_distance = path_information[node_label]
            points_by_path.setdefault(path_id, []).append(
                (path_distance, values_by_node.get(node_label), node_label)
            )
        for path_id, points in points_by_path.items():
            curves_by_path.setdefault(path_id, []).append((step_name, points))
    return curves_by_path


def xlsx_path_table(curves):
    locations = set()
    curve_maps = []
    for step_name, points in curves:
        values = {}
        for path_distance, contact_value, node_label in points:
            location = (path_distance, node_label)
            locations.add(location)
            values[location] = contact_value
        curve_maps.append((step_name, values))
    return (
        sorted(locations, key=lambda item: (item[0], item[1])),
        curve_maps,
    )


def contact_sheet_name(variable_name, path_id, used_names):
    base_name = "{0} Path {1}".format(variable_name, path_id)
    clean_name = "".join(
        "_" if character in "[]:*?/\\" else character
        for character in base_name
    )[:31]
    candidate = clean_name
    suffix = 2
    while candidate.upper() in used_names:
        suffix_text = " {0}".format(suffix)
        candidate = clean_name[: 31 - len(suffix_text)] + suffix_text
        suffix += 1
    used_names.add(candidate.upper())
    return candidate


def aggregation_text(variable_name):
    if variable_name == "CPRESS":
        return "Maximum CPRESS at each PIPE-path node; missing output is blank"
    return "Minimum COPEN at each PIPE-path node; missing output is blank"


def xlsx_worksheet_xml(
    odb_name,
    path_id,
    variable_name,
    locations,
    curve_maps,
):
    last_column = 2 + len(curve_maps)
    last_row = 4 + len(locations)
    rows = [
        '<row r="1" ht="22" customHeight="1">{0}</row>'.format(
            common.xlsx_inline_cell(
                1,
                1,
                "{0} Along PIPE Path - {1} - Path {2}".format(
                    variable_name, odb_name, path_id
                ),
                1,
            )
        ),
        '<row r="2">{0}</row>'.format(
            common.xlsx_inline_cell(
                2, 1, aggregation_text(variable_name), 2
            )
        ),
    ]
    header_cells = [
        common.xlsx_inline_cell(4, 1, "Path Distance", 3),
        common.xlsx_inline_cell(4, 2, "Node Label", 3),
    ]
    for column_number, (step_name, unused_values) in enumerate(
        curve_maps, 3
    ):
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
        for column_number, (unused_step, values) in enumerate(
            curve_maps, 3
        ):
            cells.append(
                common.xlsx_number_cell(
                    row_number, column_number, values.get(location), 5
                )
            )
        rows.append(
            '<row r="{0}">{1}</row>'.format(
                row_number, "".join(cells)
            )
        )

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


def xlsx_drawing_xml(chart_index, start_column, variable_name):
    end_column = start_column + 11
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>{0}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>{1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>23</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="{2} Chart {3}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>""".format(
        start_column, end_column, variable_name, chart_index
    )


def xlsx_chart_xml(
    odb_name,
    path_id,
    variable_name,
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

    x_axis_id = 68650112 + chart_index * 2
    y_axis_id = x_axis_id + 1
    if variable_name == "CPRESS":
        y_title = "Maximum CPRESS (model stress units)"
    else:
        y_title = "Minimum COPEN (model length units)"
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
            "{0} Along PIPE Path - {1} - Path {2}".format(
                variable_name, odb_name, path_id
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
            y_title,
            "0.000000E+00",
        ),
    )


def write_contact_excel(
    excel_path, odb_name, path_information, extracted_frames
):
    used_names = set()
    sheet_data = []
    for variable_name in FIELD_NAMES:
        curves_by_path = contact_curves(
            path_information, extracted_frames, variable_name
        )
        for path_id in sorted(curves_by_path):
            locations, curve_maps = xlsx_path_table(
                curves_by_path[path_id]
            )
            sheet_data.append(
                (
                    path_id,
                    variable_name,
                    contact_sheet_name(
                        variable_name, path_id, used_names
                    ),
                    locations,
                    curve_maps,
                )
            )

    if not sheet_data:
        raise ValueError("No PIPE-path contact data are available for Excel.")
    sheet_names = [item[2] for item in sheet_data]
    workbook_zip = zipfile.ZipFile(
        excel_path,
        "w",
        compression=zipfile.ZIP_DEFLATED,
        allowZip64=True,
    )
    try:
        workbook_zip.writestr(
            "[Content_Types].xml",
            common.xlsx_bytes(
                common.xlsx_content_types_xml(len(sheet_data))
            ),
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
            common.xlsx_bytes(
                common.xlsx_app_properties_xml(sheet_names)
            ),
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
            "xl/styles.xml",
            common.xlsx_bytes(common.xlsx_styles_xml()),
        )

        for sheet_index, item in enumerate(sheet_data, 1):
            (
                path_id,
                variable_name,
                sheet_name,
                locations,
                curve_maps,
            ) = item
            start_column = max(5, 3 + len(curve_maps))
            workbook_zip.writestr(
                "xl/worksheets/sheet{0}.xml".format(sheet_index),
                common.xlsx_bytes(
                    xlsx_worksheet_xml(
                        odb_name,
                        path_id,
                        variable_name,
                        locations,
                        curve_maps,
                    )
                ),
            )
            workbook_zip.writestr(
                "xl/worksheets/_rels/sheet{0}.xml.rels".format(
                    sheet_index
                ),
                common.xlsx_bytes(
                    common.xlsx_worksheet_relationships_xml(sheet_index)
                ),
            )
            workbook_zip.writestr(
                "xl/drawings/drawing{0}.xml".format(sheet_index),
                common.xlsx_bytes(
                    xlsx_drawing_xml(
                        sheet_index, start_column, variable_name
                    )
                ),
            )
            workbook_zip.writestr(
                "xl/drawings/_rels/drawing{0}.xml.rels".format(
                    sheet_index
                ),
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
                        variable_name,
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
    requested_element_sets,
    requested_element_ranges,
    requested_steps,
    step_range,
    frame_index,
    start_node,
    start_node_set,
):
    odb = None
    try:
        print("Opening: {0}".format(odb_path))
        odb = openOdb(path=odb_path, readOnly=True)
        instance_key = common.repository_key(
            odb.rootAssembly.instances, instance_name
        )
        if instance_key is None:
            raise ValueError(
                "Instance '{0}' was not found. Available instances: "
                "{1}".format(
                    instance_name,
                    ", ".join(odb.rootAssembly.instances.keys()),
                )
            )
        instance = odb.rootAssembly.instances[instance_key]
        (
            full_path_information,
            coordinates,
            path_origins,
            path_notes,
        ) = common.build_pipeline_paths(
            instance, start_node, start_node_set
        )
        (
            pipe_element_labels,
            selected_nodes,
            element_selection,
            element_selection_notes,
            total_pipe_elements,
        ) = resolve_pipe_element_selection(
            odb,
            instance,
            requested_element_sets,
            requested_element_ranges,
        )
        path_information = dict(
            (node_label, full_path_information[node_label])
            for node_label in selected_nodes
            if node_label in full_path_information
        )
        if not path_information:
            raise ValueError(
                "The selected PIPE elements contain no nodes on the "
                "supported PIPE path."
            )
        allowed_nodes = set(path_information.keys())

        available_steps = list(odb.steps.keys())
        selected_steps, selection_notes = common.select_steps(
            available_steps, requested_steps, step_range
        )
        if not selected_steps:
            details = (
                "; ".join(selection_notes)
                if selection_notes
                else "none selected"
            )
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
                "None of the selected steps contains both CPRESS and "
                "COPEN: {0}".format(details)
            )

        extracted_frames = []
        for step_name, selected_frame_index, frame in selected_frames:
            (
                cpress_values,
                cpress_counts,
                cpress_sources,
            ) = aggregate_contact_values(
                frame,
                "CPRESS",
                instance_key,
                allowed_nodes,
                pipe_element_labels,
            )
            (
                copen_values,
                copen_counts,
                copen_sources,
            ) = aggregate_contact_values(
                frame,
                "COPEN",
                instance_key,
                allowed_nodes,
                pipe_element_labels,
            )
            extracted_frames.append(
                (
                    step_name,
                    selected_frame_index,
                    frame,
                    cpress_values,
                    copen_values,
                    cpress_counts,
                    copen_counts,
                    cpress_sources,
                    copen_sources,
                )
            )

        with open(report_path, "w") as report_file:
            report_file.write("*" * 100 + "\n")
            report_file.write(
                "PIPE-Only CPRESS and COPEN Pipeline Report, written "
                "{0}\n\n".format(
                    datetime.datetime.now().strftime(
                        "%a %b %d %H:%M:%S %Y"
                    )
                )
            )
            report_file.write(
                "ODB: {0}\n".format(odb_path.replace("\\", "/"))
            )
            report_file.write(
                "Pipeline instance: {0}\n".format(instance_key)
            )
            report_file.write("Element filter: PIPE* element types only\n")
            report_file.write(
                "Element selection: {0}\n".format(element_selection)
            )
            report_file.write(
                "Selected PIPE elements: {0} of {1}\n".format(
                    len(pipe_element_labels), total_pipe_elements
                )
            )
            report_file.write(
                "Selected PIPE-path nodes: {0}\n".format(
                    len(path_information)
                )
            )
            report_file.write(
                "CPRESS aggregation: maximum available value at each node\n"
            )
            report_file.write(
                "COPEN aggregation: minimum available value at each node\n"
            )
            report_file.write(
                "Missing contact output: blank (not assumed to be zero)\n"
            )
            report_file.write(
                "Path distance: cumulative distance along the complete "
                "undeformed PIPE mesh\n"
            )
            report_file.write("Full-path origins:")
            for path_id, origin_node in path_origins:
                report_file.write(
                    " Path {0}=node {1};".format(path_id, origin_node)
                )
            report_file.write("\n\n")

            for extracted in extracted_frames:
                (
                    step_name,
                    selected_frame_index,
                    frame,
                    cpress_values,
                    copen_values,
                    cpress_counts,
                    copen_counts,
                    cpress_sources,
                    copen_sources,
                ) = extracted
                report_file.write("Source 1\n")
                report_file.write("---------\n\n")
                report_file.write("   Step: {0}\n".format(step_name))
                report_file.write(
                    "   Frame index: {0}\n".format(selected_frame_index)
                )
                report_file.write(
                    "   Frame: {0}\n".format(frame.description)
                )
                report_file.write(
                    "   CPRESS fields: {0}\n".format(
                        ", ".join(cpress_sources)
                    )
                )
                report_file.write(
                    "   COPEN fields: {0}\n".format(
                        ", ".join(copen_sources)
                    )
                )
                report_file.write(
                    "   PIPE nodes with CPRESS: {0} of {1}\n".format(
                        len(cpress_values), len(allowed_nodes)
                    )
                )
                report_file.write(
                    "   PIPE nodes with COPEN: {0} of {1}\n\n".format(
                        len(copen_values), len(allowed_nodes)
                    )
                )
                write_frame_table(
                    report_file,
                    path_information,
                    coordinates,
                    cpress_values,
                    copen_values,
                    cpress_counts,
                    copen_counts,
                )
                write_frame_summary(
                    report_file,
                    path_information,
                    cpress_values,
                    copen_values,
                )

        odb_stem = os.path.splitext(os.path.basename(odb_path))[0]
        write_contact_excel(
            excel_path, odb_stem, path_information, extracted_frames
        )

        print("Selected steps:")
        for step_name, selected_frame_index, unused_frame in selected_frames:
            print(
                "  {0} (frame index {1})".format(
                    step_name, selected_frame_index
                )
            )
        for note in selection_notes:
            print("Note: {0}.".format(note))
        for skipped_step, reason in skipped_steps:
            print("Skipped step '{0}': {1}.".format(skipped_step, reason))
        for note in path_notes:
            print("Path note: {0}.".format(note))
        for note in element_selection_notes:
            print("Element selection: {0}.".format(note))
        print("Wrote:      {0}".format(report_path))
        print("Excel plot: {0}".format(excel_path))
    finally:
        if odb is not None:
            odb.close()


def print_steps(odb_path):
    odb = None
    try:
        odb = openOdb(path=odb_path, readOnly=True)
        print(odb_path)
        step_names = list(odb.steps.keys())
        if not step_names:
            print("  (no analysis steps)")
        for index, step_name in enumerate(step_names, 1):
            print("  {0:>3}: {1}".format(index, step_name))
    finally:
        if odb is not None:
            odb.close()


def print_element_sets(odb_path, instance_name):
    odb = None
    try:
        odb = openOdb(path=odb_path, readOnly=True)
        instance_key = common.repository_key(
            odb.rootAssembly.instances, instance_name
        )
        print(odb_path)
        if instance_key is None:
            print(
                "  Instance '{0}' not found. Available instances: {1}".format(
                    instance_name,
                    ", ".join(odb.rootAssembly.instances.keys()),
                )
            )
            return

        instance = odb.rootAssembly.instances[instance_key]
        instance_labels = set()
        pipe_labels = set()
        for element in instance.elements:
            label = int(element.label)
            instance_labels.add(label)
            if str(element.type).upper().startswith("PIPE"):
                pipe_labels.add(label)
        instance_set_names = sorted(instance.elementSets.keys())
        assembly_set_names = sorted(odb.rootAssembly.elementSets.keys())

        print("  Instance: {0}".format(instance_key))
        print("  Total PIPE elements: {0}".format(len(pipe_labels)))
        if not instance_set_names and not assembly_set_names:
            print("  (no element sets)")
        sys.stdout.flush()

        for set_name in instance_set_names:
            labels = element_set_labels_for_instance(
                instance.elementSets[set_name],
                instance,
                instance_labels,
            )
            print(
                "  {0:>8}  {1}  elements={2}, PIPE={3}".format(
                    "instance",
                    set_name,
                    len(labels),
                    sum(1 for label in labels if label in pipe_labels),
                )
            )
        for set_name in assembly_set_names:
            labels = element_set_labels_for_instance(
                odb.rootAssembly.elementSets[set_name],
                instance,
                instance_labels,
            )
            print(
                "  {0:>8}  {1}  elements={2}, PIPE={3}".format(
                    "assembly",
                    set_name,
                    len(labels),
                    sum(1 for label in labels if label in pipe_labels),
                )
            )
    finally:
        if odb is not None:
            odb.close()


def write_execution_log(
    log_path, input_dir, output_dir, jobs, failures
):
    with open(log_path, "w") as log_file:
        log_file.write(
            "extract_cpress_copen_pipe_selected_steps.py version "
            "{0}\n".format(SCRIPT_VERSION)
        )
        log_file.write("Time: {0}\n".format(datetime.datetime.now()))
        log_file.write("sys.argv: {0}\n".format(repr(sys.argv)))
        log_file.write("Input directory: {0}\n".format(input_dir))
        log_file.write("Output directory: {0}\n".format(output_dir))
        log_file.write("ODB jobs: {0}\n\n".format(len(jobs)))
        if not failures:
            log_file.write("Completed without errors.\n")
            return
        log_file.write("Failures: {0}\n".format(len(failures)))
        for odb_path, message, traceback_text in failures:
            log_file.write("\n" + "=" * 80 + "\n")
            log_file.write("ODB: {0}\n".format(odb_path))
            log_file.write("Error: {0}\n\n".format(message))
            log_file.write(traceback_text)
            if not traceback_text.endswith("\n"):
                log_file.write("\n")


def main():
    print(
        "extract_cpress_copen_pipe_selected_steps.py version {0}".format(
            SCRIPT_VERSION
        )
    )
    args = parse_arguments()
    input_dir = os.path.abspath(args.input_dir)
    odb_paths = common.find_odb_files(input_dir, args.odb)

    if args.list_steps or args.list_element_sets:
        for odb_path in odb_paths:
            if args.list_steps:
                print_steps(odb_path)
            if args.list_element_sets:
                print_element_sets(odb_path, args.instance)
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
                requested_element_sets=args.element_set,
                requested_element_ranges=args.element_range,
                requested_steps=args.steps,
                step_range=args.step_range,
                frame_index=args.frame_index,
                start_node=args.start_node,
                start_node_set=args.start_node_set,
            )
        except Exception as exc:
            traceback_text = traceback.format_exc()
            failures.append((odb_path, str(exc), traceback_text))
            print("FAILED:  {0}".format(odb_path))
            print("         {0}".format(exc))
            print(traceback_text)

    log_path = os.path.join(
        output_dir, "extract_cpress_copen_pipe_selected_steps.log"
    )
    write_execution_log(
        log_path, input_dir, output_dir, jobs, failures
    )
    print(
        "Completed: {0} succeeded, {1} failed.".format(
            len(jobs) - len(failures), len(failures)
        )
    )
    if failures:
        print("Diagnostic log: {0}".format(log_path))
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
