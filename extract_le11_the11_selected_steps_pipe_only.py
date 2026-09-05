from __future__ import print_function

"""Extract PIPE-only LE11, THE11, and derived mechanical strain.

Run this file with the Abaqus Python interpreter. Examples:

    abaqus python extract_le11_the11_selected_steps_pipe_only.py --list-steps
    abaqus python extract_le11_the11_selected_steps_pipe_only.py --steps Step-2 Step-5
    abaqus python extract_le11_the11_selected_steps_pipe_only.py --step-range 3 7
    abaqus python extract_le11_the11_selected_steps_pipe_only.py --step-range Preload Operation

The default instance is PART-1-1. Only element types whose names begin with
PIPE are used. Results are ordered by cumulative distance along the PIPE
mesh, starting from node set START when available.
Integration-point values are extrapolated to element nodes and contributions
at matching nodes and section points are averaged. Every section point is
preserved by default and displayed as its own LE11/THE11/MECH11 column group.
Mechanical strain is calculated as MECH11 = LE11 - THE11.
Each path row also reports the maximum MECH11 across all section points.
An editable Excel workbook plots MAX_MECH11 versus path distance for all
selected steps. Each pipeline path is written to its own worksheet with a
native Excel scatter chart whose formatting can be changed by the user.
The workbook is created with the Python standard library; Excel does not need
to be installed or open while Abaqus performs the extraction.
"""

import argparse
import datetime
import heapq
import math
import os
import sys
import traceback
import zipfile

from abaqusConstants import ELEMENT_NODAL, ON
from odbAccess import openOdb


DEFAULT_INSTANCE = "PART-1-1"
DEFAULT_FRAME_INDEX = -1
REQUIRED_FIELDS = ("LE", "THE")


def parse_arguments():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    parser = argparse.ArgumentParser(
        description=(
            "Write nodal LE11 and THE11 reports plus editable MAX_MECH11 "
            "Excel charts for a complete pipeline instance and "
            "user-selected ODB steps."
        )
    )
    parser.add_argument(
        "--input-dir",
        default=script_dir,
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
        help=(
            "Node label used as path-distance zero. Overrides --start-node-set."
        ),
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
            "Zero-based frame index. Use -1 for the last frame containing LE "
            "and THE (default: -1)."
        ),
    )
    parser.add_argument(
        "--average-section-points",
        action="store_true",
        help=(
            "Average through all section points to produce one row per node. "
            "By default each section point is kept as a separate row."
        ),
    )
    parser.add_argument(
        "--list-steps",
        action="store_true",
        help="List each ODB's ordered steps and exit without writing reports.",
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

    args = parser.parse_args()
    if args.frame_index < -1:
        parser.error("--frame-index must be -1 or zero or greater")
    if args.output_name and len(args.odb) != 1:
        parser.error("--output-name requires exactly one --odb")
    return args


def find_odb_files(input_dir, requested):
    input_dir = os.path.abspath(input_dir)
    if not os.path.isdir(input_dir):
        raise ValueError("Input directory does not exist: {0}".format(input_dir))

    if requested:
        odb_paths = []
        for item in requested:
            path = item if os.path.isabs(item) else os.path.join(input_dir, item)
            path = os.path.abspath(path)
            if not os.path.isfile(path):
                raise ValueError("ODB file does not exist: {0}".format(path))
            if not path.lower().endswith(".odb"):
                raise ValueError("Not an ODB file: {0}".format(path))
            odb_paths.append(path)
    else:
        odb_paths = [
            os.path.join(input_dir, name)
            for name in os.listdir(input_dir)
            if name.lower().endswith(".odb")
            and os.path.isfile(os.path.join(input_dir, name))
        ]

    odb_paths.sort(key=lambda value: value.lower())
    if not odb_paths:
        raise ValueError("No ODB files found in: {0}".format(input_dir))
    return odb_paths


def report_name(odb_path):
    stem = os.path.splitext(os.path.basename(odb_path))[0]
    return stem + "_PIPE_ONLY_LE11_THE11_MECH11.rpt"


def build_jobs(odb_paths, output_dir, output_name):
    jobs = []
    report_paths_seen = set()
    for odb_path in odb_paths:
        name = output_name or report_name(odb_path)
        report_path = os.path.abspath(os.path.join(output_dir, name))
        normalized = os.path.normcase(report_path)
        if normalized in report_paths_seen:
            raise ValueError(
                "More than one ODB maps to report path: {0}".format(report_path)
            )
        report_paths_seen.add(normalized)
        jobs.append((odb_path, report_path))
    return jobs


def repository_key(repository, requested_name):
    if requested_name in repository:
        return requested_name
    requested_upper = requested_name.upper()
    for key in repository.keys():
        if key.upper() == requested_upper:
            return key
    return None


def node_coordinates(instance):
    coordinates = {}
    for node in instance.nodes:
        values = tuple(float(value) for value in node.coordinates)
        if len(values) == 1:
            values = (values[0], 0.0, 0.0)
        elif len(values) == 2:
            values = (values[0], values[1], 0.0)
        coordinates[int(node.label)] = values[:3]
    return coordinates


def distance_between(first, second):
    return math.sqrt(sum((first[index] - second[index]) ** 2 for index in range(3)))


def add_graph_edge(graph, first_label, second_label, length):
    if first_label == second_label:
        return
    graph.setdefault(first_label, {})
    graph.setdefault(second_label, {})
    previous = graph[first_label].get(second_label)
    if previous is None or length < previous:
        graph[first_label][second_label] = length
        graph[second_label][first_label] = length


def pipeline_graph(instance, coordinates):
    """Build path edges exclusively for two- and three-node PIPE* elements."""
    graph = {}
    skipped_types = set()
    for element in instance.elements:
        element_type = str(element.type).upper()
        if not element_type.startswith("PIPE"):
            skipped_types.add(element_type)
            continue
        connectivity = tuple(int(label) for label in element.connectivity)
        if len(connectivity) == 2:
            physical_order = connectivity
        elif len(connectivity) == 3:
            # Abaqus quadratic PIPE elements list the two ends first and the
            # midside node third. Physical order is end 1, midside, end 2.
            physical_order = (connectivity[0], connectivity[2], connectivity[1])
        else:
            skipped_types.add(element_type)
            continue

        for index in range(len(physical_order) - 1):
            first_label = physical_order[index]
            second_label = physical_order[index + 1]
            if first_label not in coordinates or second_label not in coordinates:
                continue
            length = distance_between(
                coordinates[first_label], coordinates[second_label]
            )
            add_graph_edge(graph, first_label, second_label, length)

    if not graph:
        raise ValueError(
            "No connected two- or three-node PIPE* elements were found in "
            "instance '{0}'.".format(instance.name)
        )
    return graph, skipped_types


def connected_component(graph, start_label):
    component = set()
    pending = [start_label]
    while pending:
        node_label = pending.pop()
        if node_label in component:
            continue
        component.add(node_label)
        pending.extend(
            neighbor for neighbor in graph[node_label] if neighbor not in component
        )
    return component


def choose_component_start(graph, component):
    endpoints = sorted(
        label
        for label in component
        if len([neighbor for neighbor in graph[label] if neighbor in component]) == 1
    )
    return endpoints[0] if endpoints else min(component)


def shortest_path_distances(graph, component, start_label):
    distances = dict((label, float("inf")) for label in component)
    distances[start_label] = 0.0
    queue = [(0.0, start_label)]
    while queue:
        distance, node_label = heapq.heappop(queue)
        if distance != distances[node_label]:
            continue
        for neighbor, edge_length in graph[node_label].items():
            if neighbor not in component:
                continue
            candidate = distance + edge_length
            if candidate < distances[neighbor]:
                distances[neighbor] = candidate
                heapq.heappush(queue, (candidate, neighbor))
    return distances


def requested_path_start(instance, graph, start_node, start_node_set, notes):
    if start_node is not None:
        if start_node not in graph:
            raise ValueError(
                "Start node {0} is not connected to a supported pipeline "
                "PIPE element in instance '{1}'.".format(start_node, instance.name)
            )
        return start_node

    set_key = repository_key(instance.nodeSets, start_node_set)
    if set_key is not None:
        candidates = sorted(
            int(node.label)
            for node in instance.nodeSets[set_key].nodes
            if int(node.label) in graph
        )
        if candidates:
            if len(candidates) > 1:
                notes.append(
                    "node set '{0}' contains multiple path nodes; node {1} was "
                    "used as the origin".format(set_key, candidates[0])
                )
            return candidates[0]
        notes.append(
            "node set '{0}' contains no node connected to a supported line "
            "element".format(set_key)
        )
    else:
        notes.append(
            "node set '{0}' was not found; an endpoint was selected "
            "automatically".format(start_node_set)
        )

    all_nodes = set(graph.keys())
    return choose_component_start(graph, all_nodes)


def build_pipeline_paths(instance, start_node, start_node_set):
    """Return node -> (path id, distance), coordinates, origins, and notes."""
    coordinates = node_coordinates(instance)
    graph, skipped_types = pipeline_graph(instance, coordinates)
    notes = []
    first_start = requested_path_start(
        instance, graph, start_node, start_node_set, notes
    )

    components = []
    first_component = connected_component(graph, first_start)
    components.append((first_start, first_component))
    remaining = set(graph.keys()).difference(first_component)
    while remaining:
        seed = min(remaining)
        component = connected_component(graph, seed)
        component_start = choose_component_start(graph, component)
        components.append((component_start, component))
        remaining.difference_update(component)

    path_information = {}
    origins = []
    for path_id, (component_start, component) in enumerate(components, 1):
        distances = shortest_path_distances(graph, component, component_start)
        origins.append((path_id, component_start))
        for node_label, distance in distances.items():
            path_information[node_label] = (path_id, distance)

    if len(components) > 1:
        notes.append(
            "the instance contains {0} disconnected PIPE-mesh paths; distance "
            "starts at zero for each path ID".format(len(components))
        )
    if skipped_types:
        notes.append(
            "non-PIPE element types were excluded from path construction: {0}".format(
                ", ".join(sorted(skipped_types))
            )
        )
    return path_information, coordinates, origins, notes


def find_name_case_insensitive(available_steps, requested_name):
    requested_upper = requested_name.upper()
    for step_name in available_steps:
        if step_name.upper() == requested_upper:
            return step_name
    return None


def integer_value(text):
    try:
        return int(text)
    except ValueError:
        return None


def select_named_steps(available_steps, requested_steps):
    selected = []
    notes = []
    for requested_name in requested_steps:
        actual_name = find_name_case_insensitive(available_steps, requested_name)
        if actual_name is None:
            notes.append("requested step '{0}' is not present".format(requested_name))
        elif actual_name not in selected:
            selected.append(actual_name)
    return selected, notes


def select_step_range(available_steps, endpoints):
    start_text, end_text = endpoints
    start_number = integer_value(start_text)
    end_number = integer_value(end_text)

    if (start_number is None) != (end_number is None):
        raise ValueError(
            "--step-range endpoints must both be step numbers or both be step names"
        )

    if start_number is not None:
        if start_number < 1 or end_number < 1:
            raise ValueError("numeric step ranges are 1-based and must be positive")
        if start_number > end_number:
            raise ValueError("the start of --step-range must not exceed the end")
        if start_number > len(available_steps):
            return [], [
                "range starts at {0}, but this ODB contains only {1} step(s)".format(
                    start_number, len(available_steps)
                )
            ]

        actual_end = min(end_number, len(available_steps))
        selected = available_steps[start_number - 1 : actual_end]
        notes = []
        if end_number > len(available_steps):
            notes.append(
                "range end {0} was limited to the last available step ({1})".format(
                    end_number, len(available_steps)
                )
            )
        return selected, notes

    start_name = find_name_case_insensitive(available_steps, start_text)
    end_name = find_name_case_insensitive(available_steps, end_text)
    missing = []
    if start_name is None:
        missing.append("start step '{0}' is not present".format(start_text))
    if end_name is None:
        missing.append("end step '{0}' is not present".format(end_text))
    if missing:
        return [], missing

    start_index = available_steps.index(start_name)
    end_index = available_steps.index(end_name)
    if start_index > end_index:
        raise ValueError(
            "start step '{0}' occurs after end step '{1}'".format(
                start_name, end_name
            )
        )
    return available_steps[start_index : end_index + 1], []


def select_steps(available_steps, requested_steps, step_range):
    if requested_steps is not None:
        return select_named_steps(available_steps, requested_steps)
    if step_range is not None:
        return select_step_range(available_steps, step_range)
    return list(available_steps), []


def select_frames(odb, step_names, frame_index):
    selected = []
    skipped = []
    for step_name in step_names:
        frames = odb.steps[step_name].frames
        if len(frames) == 0:
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
        missing = []
        for candidate_index in candidate_indices:
            frame = frames[candidate_index]
            available_fields = frame.fieldOutputs.keys()
            missing = [name for name in REQUIRED_FIELDS if name not in available_fields]
            if not missing:
                selection = (step_name, candidate_index, frame)
                break

        if selection is None:
            if frame_index == -1:
                reason = "no frame contains both LE and THE"
            else:
                reason = "frame {0} is missing field(s): {1}".format(
                    frame_index, ", ".join(missing)
                )
            skipped.append((step_name, reason))
        else:
            selected.append(selection)
    return selected, skipped


def component_index(field_output, requested_component):
    labels = list(field_output.componentLabels)
    requested_upper = requested_component.upper()
    for index, label in enumerate(labels):
        if label.upper() == requested_upper:
            return index
    raise ValueError(
        "Field '{0}' does not contain component '{1}'. Available components: {2}".format(
            field_output.name, requested_component, ", ".join(labels)
        )
    )


def field_value_data(value):
    try:
        data = value.data
    except Exception:
        data = value.dataDouble
    if isinstance(data, (int, float)):
        return (float(data),)
    return tuple(float(item) for item in data)


def section_point_key(value, average_section_points):
    if average_section_points:
        return (0, "AVERAGED")
    try:
        section_point = value.sectionPoint
        number = int(section_point.number)
        description = str(section_point.description).strip()
        if number == 0 and not description:
            return (0, "")
        return (number, description)
    except Exception:
        return (0, "")


def collect_component_values(
    field_output,
    requested_component,
    instance,
    average_section_points,
    pipe_element_labels,
):
    index = component_index(field_output, requested_component)
    element_nodal = field_output.getSubset(
        region=instance, position=ELEMENT_NODAL, readOnly=ON
    )
    contributions = {}

    for value in element_nodal.values:
        value_instance = getattr(value, "instance", None)
        if value_instance is not None and value_instance.name.upper() != instance.name.upper():
            continue
        element_label = getattr(value, "elementLabel", None)
        if (
            element_label is None
            or int(element_label) not in pipe_element_labels
        ):
            continue
        data = field_value_data(value)
        if index >= len(data):
            raise ValueError(
                "Field '{0}' returned only {1} component(s) at node {2}.".format(
                    field_output.name, len(data), value.nodeLabel
                )
            )
        location = section_point_key(value, average_section_points)
        key = (int(value.nodeLabel), location[0], location[1])
        contributions.setdefault(key, []).append(data[index])

    averaged = {}
    counts = {}
    for key, values in contributions.items():
        averaged[key] = sum(values) / float(len(values))
        counts[key] = len(values)
    return averaged, counts


def extract_frame_rows(
    frame,
    instance,
    average_section_points,
    path_information,
    coordinates,
    pipe_element_labels,
):
    le_values, le_counts = collect_component_values(
        frame.fieldOutputs["LE"],
        "LE11",
        instance,
        average_section_points,
        pipe_element_labels,
    )
    the_values, the_counts = collect_component_values(
        frame.fieldOutputs["THE"],
        "THE11",
        instance,
        average_section_points,
        pipe_element_labels,
    )

    all_common_locations = set(le_values.keys()).intersection(the_values.keys())
    common_locations = [
        key for key in all_common_locations if key[0] in path_information
    ]
    common_locations.sort(
        key=lambda item: (
            path_information[item[0]][0],
            path_information[item[0]][1],
            item[0],
            item[1],
            item[2],
        )
    )
    if not common_locations:
        raise ValueError(
            "LE11 and THE11 have no matching node/section-point locations in "
            "instance '{0}'.".format(instance.name)
        )

    rows = []
    for key in common_locations:
        node_label, section_number, section_description = key
        path_id, path_distance = path_information[node_label]
        x_coordinate, y_coordinate, z_coordinate = coordinates[node_label]
        rows.append(
            (
                path_id,
                path_distance,
                node_label,
                x_coordinate,
                y_coordinate,
                z_coordinate,
                section_number,
                section_description,
                le_values[key],
                the_values[key],
            )
        )
    unmatched_le = len(set(le_values.keys()).difference(the_values.keys()))
    unmatched_the = len(set(the_values.keys()).difference(le_values.keys()))
    unmatched_path = len(all_common_locations) - len(common_locations)
    max_contributions = max(
        max(le_counts.values()),
        max(the_counts.values()),
    )
    return rows, unmatched_le, unmatched_the, unmatched_path, max_contributions


def engineering_format(value):
    if math.isnan(value) or math.isinf(value):
        return str(value)
    if value == 0.0:
        return "0.00000"

    exponent = int(math.floor(math.log10(abs(value)) / 3.0) * 3)
    mantissa = value / (10.0 ** exponent)
    digits_before_decimal = int(math.floor(math.log10(abs(mantissa)))) + 1
    decimal_places = max(0, 6 - digits_before_decimal)
    mantissa_text = ("{0:.%df}" % decimal_places).format(mantissa)
    if abs(float(mantissa_text)) >= 1000.0:
        exponent += 3
        mantissa /= 1000.0
        digits_before_decimal = int(math.floor(math.log10(abs(mantissa)))) + 1
        decimal_places = max(0, 6 - digits_before_decimal)
        mantissa_text = ("{0:.%df}" % decimal_places).format(mantissa)
    return "{0}E{1:+03d}".format(mantissa_text, exponent)


def section_point_text(number, description):
    if description == "AVERAGED":
        return "AVERAGED"
    if number == 0 and not description:
        return "-"
    if description:
        return "SP {0}: {1}".format(number, description)
    return "SP {0}".format(number)


def pivot_section_point_rows(rows):
    """Convert one-row-per-section-point data into one row per path node."""
    section_points = sorted(
        set((row[6], row[7]) for row in rows),
        key=lambda item: (item[0], item[1]),
    )
    locations = {}
    for row in rows:
        location_key = row[:6]
        section_key = (row[6], row[7])
        locations.setdefault(location_key, {})[section_key] = (row[8], row[9])
    location_keys = sorted(
        locations.keys(), key=lambda item: (item[0], item[1], item[2])
    )
    return section_points, [(key, locations[key]) for key in location_keys]


def maximum_mechanical_curves(extracted_frames):
    """Return path ID -> ordered (step name, path points) curve data."""
    curves_by_path = {}
    for extracted in extracted_frames:
        step_name = extracted[0]
        rows = extracted[3]
        unused_sections, pivoted_rows = pivot_section_point_rows(rows)
        points_by_path = {}
        for location, values_by_section in pivoted_rows:
            path_id, path_distance, node_label = location[:3]
            maximum_mechanical = max(
                le11 - the11 for le11, the11 in values_by_section.values()
            )
            points_by_path.setdefault(path_id, []).append(
                (path_distance, maximum_mechanical, node_label)
            )
        for path_id, points in points_by_path.items():
            points.sort(key=lambda item: (item[0], item[2]))
            curves_by_path.setdefault(path_id, []).append((step_name, points))
    return curves_by_path


def xml_escape(value):
    return (
        str(value)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def plot_tick(value):
    if value == 0.0:
        return "0"
    return "{0:.4g}".format(value)


def build_max_mechanical_svg(odb_name, extracted_frames):
    curves_by_path = maximum_mechanical_curves(extracted_frames)
    if not curves_by_path:
        raise ValueError("No MAX_MECH11 path data are available for plotting.")

    colors = (
        "#1565C0",
        "#D32F2F",
        "#2E7D32",
        "#7B1FA2",
        "#EF6C00",
        "#00838F",
        "#5D4037",
        "#C2185B",
        "#455A64",
        "#6A1B9A",
    )
    width = 1200
    left = 105
    right = 270
    plot_width = width - left - right
    panel_descriptions = []
    total_height = 65
    for path_id in sorted(curves_by_path):
        curves = curves_by_path[path_id]
        panel_height = max(350, 115 + 22 * len(curves))
        panel_descriptions.append((path_id, curves, total_height, panel_height))
        total_height += panel_height
    total_height += 35

    svg = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<svg xmlns="http://www.w3.org/2000/svg" width="{0}" height="{1}" viewBox="0 0 {0} {1}">'.format(
            width, total_height
        ),
        '<rect x="0" y="0" width="100%" height="100%" fill="white"/>',
        '<style>text { font-family: Arial, sans-serif; fill: #222; } .tick { font-size: 12px; } .label { font-size: 14px; font-weight: bold; } .title { font-size: 18px; font-weight: bold; }</style>',
        '<text x="{0}" y="30" text-anchor="middle" class="title">MAX_MECH11 Along Pipeline Path - {1}</text>'.format(
            width / 2.0, xml_escape(odb_name)
        ),
        '<text x="{0}" y="50" text-anchor="middle" class="tick">MECH11 = LE11 - THE11; maximum across all section points at each path node</text>'.format(
            width / 2.0
        ),
    ]

    for path_id, curves, panel_top, panel_height in panel_descriptions:
        plot_top = panel_top + 35
        plot_bottom = panel_top + panel_height - 55
        plot_height = plot_bottom - plot_top
        plot_right = left + plot_width
        all_points = [point for unused_name, points in curves for point in points]
        x_min = min(point[0] for point in all_points)
        x_max = max(point[0] for point in all_points)
        y_min = min(point[1] for point in all_points)
        y_max = max(point[1] for point in all_points)
        if x_max == x_min:
            x_max = x_min + 1.0
        if y_max == y_min:
            padding = max(abs(y_min) * 0.1, 1.0e-12)
        else:
            padding = (y_max - y_min) * 0.08
        y_min -= padding
        y_max += padding

        def x_position(value):
            return left + (value - x_min) * plot_width / float(x_max - x_min)

        def y_position(value):
            return plot_bottom - (value - y_min) * plot_height / float(y_max - y_min)

        svg.append(
            '<text x="{0}" y="{1}" class="label">Path {2}</text>'.format(
                left, panel_top + 20, path_id
            )
        )
        svg.append(
            '<rect x="{0}" y="{1}" width="{2}" height="{3}" fill="#FAFAFA" stroke="#555" stroke-width="1"/>'.format(
                left, plot_top, plot_width, plot_height
            )
        )

        tick_count = 5
        for tick_index in range(tick_count + 1):
            fraction = tick_index / float(tick_count)
            x_value = x_min + fraction * (x_max - x_min)
            x_value_position = x_position(x_value)
            svg.append(
                '<line x1="{0:.2f}" y1="{1}" x2="{0:.2f}" y2="{2}" stroke="#E0E0E0"/>'.format(
                    x_value_position, plot_top, plot_bottom
                )
            )
            svg.append(
                '<text x="{0:.2f}" y="{1}" text-anchor="middle" class="tick">{2}</text>'.format(
                    x_value_position, plot_bottom + 20, xml_escape(plot_tick(x_value))
                )
            )

            y_value = y_min + fraction * (y_max - y_min)
            y_value_position = y_position(y_value)
            svg.append(
                '<line x1="{0}" y1="{1:.2f}" x2="{2}" y2="{1:.2f}" stroke="#E0E0E0"/>'.format(
                    left, y_value_position, plot_right
                )
            )
            svg.append(
                '<text x="{0}" y="{1:.2f}" text-anchor="end" dominant-baseline="middle" class="tick">{2}</text>'.format(
                    left - 8, y_value_position, xml_escape(plot_tick(y_value))
                )
            )

        if y_min <= 0.0 <= y_max:
            zero_position = y_position(0.0)
            svg.append(
                '<line x1="{0}" y1="{1:.2f}" x2="{2}" y2="{1:.2f}" stroke="#777" stroke-width="1.2"/>'.format(
                    left, zero_position, plot_right
                )
            )

        svg.append(
            '<text x="{0}" y="{1}" text-anchor="middle" class="label">Path Distance</text>'.format(
                left + plot_width / 2.0, plot_bottom + 43
            )
        )
        svg.append(
            '<text x="25" y="{0}" text-anchor="middle" class="label" transform="rotate(-90 25 {0})">MAX_MECH11</text>'.format(
                plot_top + plot_height / 2.0
            )
        )

        for curve_index, (step_name, points) in enumerate(curves):
            color = colors[curve_index % len(colors)]
            dash_group = curve_index // len(colors)
            dash_patterns = ("", "8,4", "3,3", "10,3,2,3")
            dash_pattern = dash_patterns[dash_group % len(dash_patterns)]
            coordinates = " ".join(
                "{0:.2f},{1:.2f}".format(x_position(point[0]), y_position(point[1]))
                for point in points
            )
            dash_attribute = (
                ' stroke-dasharray="{0}"'.format(dash_pattern)
                if dash_pattern
                else ""
            )
            svg.append(
                '<polyline points="{0}" fill="none" stroke="{1}" stroke-width="2"{2}/>'.format(
                    coordinates, color, dash_attribute
                )
            )
            if len(points) <= 250:
                for point in points:
                    svg.append(
                        '<circle cx="{0:.2f}" cy="{1:.2f}" r="2.2" fill="{2}"/>'.format(
                            x_position(point[0]), y_position(point[1]), color
                        )
                    )

            legend_y = plot_top + 18 + 22 * curve_index
            legend_x = plot_right + 25
            svg.append(
                '<line x1="{0}" y1="{1}" x2="{2}" y2="{1}" stroke="{3}" stroke-width="2"{4}/>'.format(
                    legend_x,
                    legend_y,
                    legend_x + 28,
                    color,
                    dash_attribute,
                )
            )
            svg.append(
                '<text x="{0}" y="{1}" dominant-baseline="middle" class="tick">{2}</text>'.format(
                    legend_x + 36, legend_y, xml_escape(step_name)
                )
            )

    svg.append("</svg>")
    return "\n".join(svg)


def write_max_mechanical_plot(plot_path, odb_name, extracted_frames):
    svg_text = build_max_mechanical_svg(odb_name, extracted_frames)
    with open(plot_path, "wb") as plot_file:
        plot_file.write(svg_text.encode("utf-8"))


EXCEL_CHART_COLORS = (
    "1565C0",
    "D32F2F",
    "2E7D32",
    "7B1FA2",
    "EF6C00",
    "00838F",
    "5D4037",
    "C2185B",
    "455A64",
    "6A1B9A",
)


try:
    TEXT_TYPE = unicode
    BINARY_TYPE = str
except NameError:
    TEXT_TYPE = str
    BINARY_TYPE = bytes


def text_value(value):
    if isinstance(value, TEXT_TYPE):
        return value
    if isinstance(value, BINARY_TYPE):
        return value.decode("utf-8", "replace")
    return TEXT_TYPE(value)


def xlsx_xml_escape(value):
    return (
        text_value(value)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def xlsx_bytes(value):
    return text_value(value).encode("utf-8")


def excel_column_name(column_number):
    """Return an Excel column name for a one-based column number."""
    letters = []
    while column_number:
        column_number, remainder = divmod(column_number - 1, 26)
        letters.append(chr(ord("A") + remainder))
    return "".join(reversed(letters))


def excel_cell_reference(row_number, column_number):
    return "{0}{1}".format(excel_column_name(column_number), row_number)


def xlsx_inline_cell(row_number, column_number, value, style_index=0):
    reference = excel_cell_reference(row_number, column_number)
    return (
        '<c r="{0}" t="inlineStr" s="{1}"><is><t>{2}</t></is></c>'.format(
            reference, style_index, xlsx_xml_escape(value)
        )
    )


def xlsx_number_cell(row_number, column_number, value, style_index=0):
    reference = excel_cell_reference(row_number, column_number)
    if value is None or math.isnan(value) or math.isinf(value):
        return ""
    return '<c r="{0}" s="{1}"><v>{2:.15g}</v></c>'.format(
        reference, style_index, float(value)
    )


def xlsx_styles_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="2">
    <numFmt numFmtId="164" formatCode="0.0000"/>
    <numFmt numFmtId="165" formatCode="0.000000E+00"/>
  </numFmts>
  <fonts count="4">
    <font><sz val="10"/><name val="Arial"/><family val="2"/></font>
    <font><b/><color rgb="FFFFFFFF"/><sz val="10"/><name val="Arial"/><family val="2"/></font>
    <font><b/><sz val="14"/><name val="Arial"/><family val="2"/></font>
    <font><i/><color rgb="FF666666"/><sz val="10"/><name val="Arial"/><family val="2"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FFFFFFFF"/></left><right style="thin"><color rgb="FFFFFFFF"/></right><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="6">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""


def xlsx_core_properties_xml():
    timestamp = datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>Abaqus postprocessing script</dc:creator>
  <cp:lastModifiedBy>Abaqus postprocessing script</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{0}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{0}</dcterms:modified>
</cp:coreProperties>""".format(timestamp)


def xlsx_app_properties_xml(sheet_names):
    titles = "".join(
        "<vt:lpstr>{0}</vt:lpstr>".format(xlsx_xml_escape(name))
        for name in sheet_names
    )
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Excel Compatible</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>{0}</vt:i4></vt:variant></vt:vector></HeadingPairs>
  <TitlesOfParts><vt:vector size="{0}" baseType="lpstr">{1}</vt:vector></TitlesOfParts>
  <Company></Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0300</AppVersion>
</Properties>""".format(len(sheet_names), titles)


def xlsx_content_types_xml(sheet_count):
    overrides = []
    for index in range(1, sheet_count + 1):
        overrides.extend(
            [
                '<Override PartName="/xl/worksheets/sheet{0}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'.format(index),
                '<Override PartName="/xl/drawings/drawing{0}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'.format(index),
                '<Override PartName="/xl/charts/chart{0}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'.format(index),
            ]
        )
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  {0}
</Types>""".format("\n  ".join(overrides))


def xlsx_root_relationships_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"""


def xlsx_workbook_xml(sheet_names):
    sheets = "".join(
        '<sheet name="{0}" sheetId="{1}" r:id="rId{1}"/>'.format(
            xlsx_xml_escape(name), index
        )
        for index, name in enumerate(sheet_names, 1)
    )
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews><workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="12000"/></bookViews>
  <sheets>{0}</sheets>
  <calcPr calcId="191029" fullCalcOnLoad="1"/>
</workbook>""".format(sheets)


def xlsx_workbook_relationships_xml(sheet_count):
    relationships = []
    for index in range(1, sheet_count + 1):
        relationships.append(
            '<Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{0}.xml"/>'.format(index)
        )
    relationships.append(
        '<Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'.format(
            sheet_count + 1
        )
    )
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  {0}
</Relationships>""".format("\n  ".join(relationships))


def xlsx_sheet_name(path_id, used_names):
    base_name = "Path {0}".format(path_id)
    clean_name = "".join(
        "_" if character in "[]:*?/\\" else character for character in base_name
    )[:31]
    candidate = clean_name
    suffix = 2
    while candidate.lower() in used_names:
        suffix_text = " ({0})".format(suffix)
        candidate = clean_name[: 31 - len(suffix_text)] + suffix_text
        suffix += 1
    used_names.add(candidate.lower())
    return candidate


def xlsx_path_table(curves):
    locations = set()
    curve_maps = []
    for step_name, points in curves:
        values = {}
        for path_distance, maximum_mechanical, node_label in points:
            key = (path_distance, node_label)
            locations.add(key)
            values[key] = maximum_mechanical
        curve_maps.append((step_name, values))
    return sorted(locations, key=lambda item: (item[0], item[1])), curve_maps


def xlsx_worksheet_xml(odb_name, path_id, locations, curve_maps):
    last_column = 2 + len(curve_maps)
    last_row = 4 + len(locations)
    rows = [
        '<row r="1" ht="22" customHeight="1">{0}</row>'.format(
            xlsx_inline_cell(
                1,
                1,
                "MAX_MECH11 Along Pipeline Path - {0} - Path {1}".format(
                    odb_name, path_id
                ),
                1,
            )
        ),
        '<row r="2">{0}</row>'.format(
            xlsx_inline_cell(
                2,
                1,
                "MECH11 = LE11 - THE11; maximum across all section points at each path node",
                2,
            )
        ),
    ]
    header_cells = [
        xlsx_inline_cell(4, 1, "Path Distance", 3),
        xlsx_inline_cell(4, 2, "Node Label", 3),
    ]
    for column_number, (step_name, unused_values) in enumerate(curve_maps, 3):
        header_cells.append(
            xlsx_inline_cell(4, column_number, "{0} MAX_MECH11".format(step_name), 3)
        )
    rows.append(
        '<row r="4" ht="30" customHeight="1">{0}</row>'.format(
            "".join(header_cells)
        )
    )

    for row_number, location in enumerate(locations, 5):
        path_distance, node_label = location
        cells = [
            xlsx_number_cell(row_number, 1, path_distance, 4),
            xlsx_number_cell(row_number, 2, node_label, 0),
        ]
        for column_number, (unused_step, values) in enumerate(curve_maps, 3):
            cells.append(
                xlsx_number_cell(row_number, column_number, values.get(location), 5)
            )
        rows.append('<row r="{0}">{1}</row>'.format(row_number, "".join(cells)))

    columns = [
        '<col min="1" max="1" width="16" customWidth="1"/>',
        '<col min="2" max="2" width="13" customWidth="1"/>',
    ]
    if last_column >= 3:
        columns.append(
            '<col min="3" max="{0}" width="22" customWidth="1"/>'.format(
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
        excel_column_name(last_column), last_row, "".join(columns), "".join(rows)
    )


def xlsx_worksheet_relationships_xml(drawing_index):
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing{0}.xml"/>
</Relationships>""".format(drawing_index)


def xlsx_drawing_xml(chart_index, start_column):
    end_column = start_column + 11
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>{0}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>{1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>23</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="MAX_MECH11 Chart {2}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>""".format(start_column, end_column, chart_index)


def xlsx_drawing_relationships_xml(chart_index):
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart{0}.xml"/>
</Relationships>""".format(chart_index)


def xlsx_chart_text(text, font_size):
    return """<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="{0}" b="0"/><a:t>{1}</a:t></a:r></a:p></c:rich></c:tx>""".format(
        font_size, xlsx_xml_escape(text)
    )


def xlsx_numeric_cache(values):
    points = []
    for index, value in enumerate(values):
        if value is not None and not math.isnan(value) and not math.isinf(value):
            points.append('<c:pt idx="{0}"><c:v>{1:.15g}</c:v></c:pt>'.format(index, value))
    return '<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="{0}"/>{1}</c:numCache>'.format(
        len(values), "".join(points)
    )


def xlsx_chart_series_xml(
    series_index, sheet_name, header_row, data_start_row, data_end_row,
    value_column, step_name, locations, values
):
    color = EXCEL_CHART_COLORS[series_index % len(EXCEL_CHART_COLORS)]
    sheet_formula_name = sheet_name.replace("'", "''")
    x_values = [location[0] for location in locations]
    y_values = [values.get(location) for location in locations]
    x_formula = "'{0}'!$A${1}:$A${2}".format(
        sheet_formula_name, data_start_row, data_end_row
    )
    value_letter = excel_column_name(value_column)
    y_formula = "'{0}'!${1}${2}:${1}${3}".format(
        sheet_formula_name, value_letter, data_start_row, data_end_row
    )
    name_formula = "'{0}'!${1}${2}".format(
        sheet_formula_name, value_letter, header_row
    )
    return """<c:ser>
  <c:idx val="{0}"/><c:order val="{0}"/>
  <c:tx><c:strRef><c:f>{1}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{2}</c:v></c:pt></c:strCache></c:strRef></c:tx>
  <c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="{3}"/></a:solidFill></a:ln></c:spPr>
  <c:marker><c:symbol val="circle"/><c:size val="4"/><c:spPr><a:solidFill><a:srgbClr val="{3}"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="{3}"/></a:solidFill></a:ln></c:spPr></c:marker>
  <c:xVal><c:numRef><c:f>{4}</c:f>{5}</c:numRef></c:xVal>
  <c:yVal><c:numRef><c:f>{6}</c:f>{7}</c:numRef></c:yVal>
  <c:smooth val="0"/>
</c:ser>""".format(
        series_index,
        xlsx_xml_escape(name_formula),
        xlsx_xml_escape(step_name),
        color,
        xlsx_xml_escape(x_formula),
        xlsx_numeric_cache(x_values),
        xlsx_xml_escape(y_formula),
        xlsx_numeric_cache(y_values),
    )


def xlsx_value_axis_xml(axis_id, cross_axis_id, position, title, number_format):
    return """<c:valAx>
  <c:axId val="{0}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="{1}"/>
  <c:title>{2}<c:layout/><c:overlay val="0"/></c:title>
  <c:numFmt formatCode="{3}" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>
  <c:spPr><a:ln><a:solidFill><a:srgbClr val="666666"/></a:solidFill></a:ln></c:spPr>
  <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>
  <c:crossAx val="{4}"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/>
</c:valAx>""".format(
        axis_id,
        position,
        xlsx_chart_text(title, 1000),
        xlsx_xml_escape(number_format),
        cross_axis_id,
    )


def xlsx_chart_xml(odb_name, path_id, sheet_name, locations, curve_maps):
    data_start_row = 5
    data_end_row = 4 + len(locations)
    series_xml = []
    for series_index, (step_name, values) in enumerate(curve_maps):
        series_xml.append(
            xlsx_chart_series_xml(
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
    x_axis_id = 48650112 + int(path_id) * 2 if isinstance(path_id, int) else 48650112
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
        xlsx_chart_text(
            "MAX_MECH11 Along Pipeline Path - {0} - Path {1}".format(
                odb_name, path_id
            ),
            1200,
        ),
        "".join(series_xml),
        x_axis_id,
        y_axis_id,
        xlsx_value_axis_xml(
            x_axis_id, y_axis_id, "b", "Path Distance", "0.0000"
        ),
        xlsx_value_axis_xml(
            y_axis_id, x_axis_id, "l", "MAX_MECH11", "0.000E+00"
        ),
    )


def write_max_mechanical_excel(excel_path, odb_name, extracted_frames):
    """Write editable data tables and native Excel charts without add-ons."""
    curves_by_path = maximum_mechanical_curves(extracted_frames)
    if not curves_by_path:
        raise ValueError("No MAX_MECH11 path data are available for plotting.")

    used_names = set()
    sheet_data = []
    for path_id in sorted(curves_by_path):
        curves = curves_by_path[path_id]
        locations, curve_maps = xlsx_path_table(curves)
        sheet_data.append(
            (
                path_id,
                xlsx_sheet_name(path_id, used_names),
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
            "[Content_Types].xml", xlsx_bytes(xlsx_content_types_xml(len(sheet_data)))
        )
        workbook_zip.writestr(
            "_rels/.rels", xlsx_bytes(xlsx_root_relationships_xml())
        )
        workbook_zip.writestr(
            "docProps/core.xml", xlsx_bytes(xlsx_core_properties_xml())
        )
        workbook_zip.writestr(
            "docProps/app.xml", xlsx_bytes(xlsx_app_properties_xml(sheet_names))
        )
        workbook_zip.writestr(
            "xl/workbook.xml", xlsx_bytes(xlsx_workbook_xml(sheet_names))
        )
        workbook_zip.writestr(
            "xl/_rels/workbook.xml.rels",
            xlsx_bytes(xlsx_workbook_relationships_xml(len(sheet_data))),
        )
        workbook_zip.writestr("xl/styles.xml", xlsx_bytes(xlsx_styles_xml()))

        for sheet_index, item in enumerate(sheet_data, 1):
            path_id, sheet_name, locations, curve_maps = item
            start_column = max(5, 3 + len(curve_maps))
            workbook_zip.writestr(
                "xl/worksheets/sheet{0}.xml".format(sheet_index),
                xlsx_bytes(
                    xlsx_worksheet_xml(
                        odb_name, path_id, locations, curve_maps
                    )
                ),
            )
            workbook_zip.writestr(
                "xl/worksheets/_rels/sheet{0}.xml.rels".format(sheet_index),
                xlsx_bytes(xlsx_worksheet_relationships_xml(sheet_index)),
            )
            workbook_zip.writestr(
                "xl/drawings/drawing{0}.xml".format(sheet_index),
                xlsx_bytes(xlsx_drawing_xml(sheet_index, start_column)),
            )
            workbook_zip.writestr(
                "xl/drawings/_rels/drawing{0}.xml.rels".format(sheet_index),
                xlsx_bytes(xlsx_drawing_relationships_xml(sheet_index)),
            )
            workbook_zip.writestr(
                "xl/charts/chart{0}.xml".format(sheet_index),
                xlsx_bytes(
                    xlsx_chart_xml(
                        odb_name,
                        path_id,
                        sheet_name,
                        locations,
                        curve_maps,
                    )
                ),
            )
    finally:
        workbook_zip.close()


def write_frame_table(report_file, rows, section_points):
    unused_frame_sections, pivoted_rows = pivot_section_point_rows(rows)
    base_top = "{0:>8}{1:>16}{2:>12}{3:>16}{4:>16}{5:>16}".format(
        "Path ID", "Path Distance", "Node Label", "X", "Y", "Z"
    )
    base_bottom = " " * len(base_top)
    section_top = ""
    section_bottom = ""
    for section_number, section_description in section_points:
        title = section_point_text(section_number, section_description)
        section_top += "{0:^48}".format(title[:48])
        section_bottom += "{0:>16}{1:>16}{2:>16}".format(
            "LE.LE11", "THE.THE11", "MECH11"
        )

    maximum_column_width = 20
    header_top = base_top + section_top + "{0:^{1}}".format(
        "All Section Points", maximum_column_width
    )
    header_bottom = base_bottom + section_bottom + "{0:>{1}}".format(
        "MAX_MECH11", maximum_column_width
    )
    report_file.write(header_top + "\n")
    report_file.write(header_bottom + "\n")
    report_file.write("-" * len(header_top) + "\n")

    for location, values_by_section in pivoted_rows:
        (
            path_id,
            path_distance,
            node_label,
            x_coordinate,
            y_coordinate,
            z_coordinate,
        ) = location
        line = "{0:>8d}{1:>16}{2:>12d}{3:>16}{4:>16}{5:>16}".format(
            path_id,
            engineering_format(path_distance),
            node_label,
            engineering_format(x_coordinate),
            engineering_format(y_coordinate),
            engineering_format(z_coordinate),
        )
        for section_key in section_points:
            if section_key in values_by_section:
                le11, the11 = values_by_section[section_key]
                mechanical_strain = le11 - the11
                line += "{0:>16}{1:>16}{2:>16}".format(
                    engineering_format(le11),
                    engineering_format(the11),
                    engineering_format(mechanical_strain),
                )
            else:
                line += "{0:>16}{1:>16}{2:>16}".format("", "", "")
        maximum_mechanical_at_node = max(
            le11 - the11 for le11, the11 in values_by_section.values()
        )
        line += "{0:>{1}}".format(
            engineering_format(maximum_mechanical_at_node), maximum_column_width
        )
        report_file.write(line + "\n")

    minimum_le = min(rows, key=lambda row: row[8])
    maximum_le = max(rows, key=lambda row: row[8])
    minimum_the = min(rows, key=lambda row: row[9])
    maximum_the = max(rows, key=lambda row: row[9])
    minimum_mechanical = min(rows, key=lambda row: row[8] - row[9])
    maximum_mechanical = max(rows, key=lambda row: row[8] - row[9])
    report_file.write("\n")
    report_file.write(
        "Minimum LE11: {0} at path {1}, distance {2}, node {3}, {4}\n".format(
            engineering_format(minimum_le[8]),
            minimum_le[0],
            engineering_format(minimum_le[1]),
            minimum_le[2],
            section_point_text(minimum_le[6], minimum_le[7]),
        )
    )
    report_file.write(
        "Maximum LE11: {0} at path {1}, distance {2}, node {3}, {4}\n".format(
            engineering_format(maximum_le[8]),
            maximum_le[0],
            engineering_format(maximum_le[1]),
            maximum_le[2],
            section_point_text(maximum_le[6], maximum_le[7]),
        )
    )
    report_file.write(
        "Minimum THE11: {0} at path {1}, distance {2}, node {3}, {4}\n".format(
            engineering_format(minimum_the[9]),
            minimum_the[0],
            engineering_format(minimum_the[1]),
            minimum_the[2],
            section_point_text(minimum_the[6], minimum_the[7]),
        )
    )
    report_file.write(
        "Maximum THE11: {0} at path {1}, distance {2}, node {3}, {4}\n".format(
            engineering_format(maximum_the[9]),
            maximum_the[0],
            engineering_format(maximum_the[1]),
            maximum_the[2],
            section_point_text(maximum_the[6], maximum_the[7]),
        )
    )
    report_file.write(
        "Minimum MECH11 (LE11-THE11): {0} at path {1}, distance {2}, "
        "node {3}, {4}\n".format(
            engineering_format(minimum_mechanical[8] - minimum_mechanical[9]),
            minimum_mechanical[0],
            engineering_format(minimum_mechanical[1]),
            minimum_mechanical[2],
            section_point_text(minimum_mechanical[6], minimum_mechanical[7]),
        )
    )
    report_file.write(
        "Maximum MECH11 (LE11-THE11): {0} at path {1}, distance {2}, "
        "node {3}, {4}\n\n\n".format(
            engineering_format(maximum_mechanical[8] - maximum_mechanical[9]),
            maximum_mechanical[0],
            engineering_format(maximum_mechanical[1]),
            maximum_mechanical[2],
            section_point_text(maximum_mechanical[6], maximum_mechanical[7]),
        )
    )


def write_report(
    odb_path,
    report_path,
    instance_name,
    requested_steps,
    step_range,
    frame_index,
    average_section_points,
    start_node,
    start_node_set,
):
    odb = None
    try:
        print("Opening: {0}".format(odb_path))
        odb = openOdb(path=odb_path, readOnly=True)
        available_steps = list(odb.steps.keys())
        selected_steps, selection_notes = select_steps(
            available_steps, requested_steps, step_range
        )
        if not selected_steps:
            raise ValueError(
                "No requested steps are available. Available steps: {0}".format(
                    ", ".join(available_steps)
                )
            )

        instance_key = repository_key(odb.rootAssembly.instances, instance_name)
        if instance_key is None:
            raise ValueError(
                "Pipeline instance '{0}' is missing. Available instances: {1}".format(
                    instance_name, ", ".join(odb.rootAssembly.instances.keys())
                )
            )
        instance = odb.rootAssembly.instances[instance_key]
        path_information, coordinates, path_origins, path_notes = build_pipeline_paths(
            instance, start_node, start_node_set
        )
        pipe_element_labels = set(
            int(element.label)
            for element in instance.elements
            if str(element.type).upper().startswith("PIPE")
        )

        selected_frames, skipped_steps = select_frames(
            odb, selected_steps, frame_index
        )
        if not selected_frames:
            raise ValueError("No selected step contains both LE and THE.")

        extracted_frames = []
        maximum_contributions = 0
        for step_name, selected_frame_index, frame in selected_frames:
            (
                rows,
                unmatched_le,
                unmatched_the,
                unmatched_path,
                max_contributions,
            ) = extract_frame_rows(
                frame,
                instance,
                average_section_points,
                path_information,
                coordinates,
                pipe_element_labels,
            )
            maximum_contributions = max(maximum_contributions, max_contributions)
            extracted_frames.append(
                (
                    step_name,
                    selected_frame_index,
                    frame,
                    rows,
                    unmatched_le,
                    unmatched_the,
                    unmatched_path,
                )
            )

        section_points = sorted(
            set(
                (row[6], row[7])
                for extracted in extracted_frames
                for row in extracted[3]
            ),
            key=lambda item: (item[0], item[1]),
        )

        with open(report_path, "w") as report_file:
            report_file.write("*" * 90 + "\n")
            report_file.write(
                "LE11, THE11, and MECH11 Pipeline Report, written {0}\n\n".format(
                    datetime.datetime.now().strftime("%a %b %d %H:%M:%S %Y")
                )
            )
            report_file.write("ODB: {0}\n".format(odb_path.replace("\\", "/")))
            report_file.write("Pipeline instance: {0}\n".format(instance_key))
            report_file.write("Element filter: PIPE* element types only\n")
            report_file.write("Mechanical strain: MECH11 = LE11 - THE11\n")
            report_file.write(
                "Path distance: cumulative distance along undeformed PIPE mesh\n"
            )
            report_file.write("Path origins:")
            for path_id, origin_node in path_origins:
                report_file.write(" Path {0}=node {1};".format(path_id, origin_node))
            report_file.write("\n")
            report_file.write(
                "Section points: {0}\n\n".format(
                    "averaged together" if average_section_points else "preserved"
                )
            )
            report_file.write(
                "Section-point columns: {0}\n\n".format(
                    ", ".join(
                        section_point_text(number, description)
                        for number, description in section_points
                    )
                )
            )
            for (
                step_name,
                selected_frame_index,
                frame,
                rows,
                unmatched_le,
                unmatched_the,
                unmatched_path,
            ) in extracted_frames:
                report_file.write("Source 1\n")
                report_file.write("---------\n\n")
                report_file.write("   Step: {0}\n".format(step_name))
                report_file.write("   Frame index: {0}\n".format(selected_frame_index))
                report_file.write("   Frame: {0}\n".format(frame.description))
                report_file.write("   Output position: averaged element-nodal\n")
                report_file.write("   Matching locations: {0}\n".format(len(rows)))
                if unmatched_le or unmatched_the:
                    report_file.write(
                        "   Unmatched locations omitted: LE11={0}, THE11={1}\n".format(
                            unmatched_le, unmatched_the
                        )
                    )
                if unmatched_path:
                    report_file.write(
                        "   Matching strain locations omitted because they are not "
                        "on a supported pipeline path: {0}\n".format(unmatched_path)
                    )
                report_file.write("\n")
                write_frame_table(report_file, rows, section_points)

        odb_stem = os.path.splitext(os.path.basename(odb_path))[0]
        excel_plot_path = os.path.join(
            os.path.dirname(report_path),
            odb_stem + "_PIPE_ONLY_MAX_MECH11_along_path.xlsx",
        )
        write_max_mechanical_excel(
            excel_plot_path, odb_stem, extracted_frames
        )

        print("Selected steps:")
        for step_name, selected_frame_index, unused_frame in selected_frames:
            print("  {0} (frame index {1})".format(step_name, selected_frame_index))
        for note in selection_notes:
            print("Note: {0}.".format(note))
        for skipped_step, reason in skipped_steps:
            print("Skipped step '{0}': {1}.".format(skipped_step, reason))
        for note in path_notes:
            print("Path note: {0}.".format(note))
        if maximum_contributions > 1:
            print(
                "Note: up to {0} element-nodal contributions were averaged at "
                "a matching node/section point.".format(maximum_contributions)
            )
        print("Wrote:   {0}".format(report_path))
        print("Excel plot: {0}".format(excel_plot_path))
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


def main():
    args = parse_arguments()
    input_dir = os.path.abspath(args.input_dir)
    odb_paths = find_odb_files(input_dir, args.odb)

    if args.list_steps:
        for odb_path in odb_paths:
            print_steps(odb_path)
        return 0

    output_dir = os.path.abspath(args.output_dir or input_dir)
    if not os.path.isdir(output_dir):
        os.makedirs(output_dir)
    jobs = build_jobs(odb_paths, output_dir, args.output_name)

    failures = []
    for odb_path, report_path in jobs:
        try:
            write_report(
                odb_path=odb_path,
                report_path=report_path,
                instance_name=args.instance,
                requested_steps=args.steps,
                step_range=args.step_range,
                frame_index=args.frame_index,
                average_section_points=args.average_section_points,
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
