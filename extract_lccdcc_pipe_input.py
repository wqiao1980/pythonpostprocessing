from __future__ import print_function

"""Create LCCDCC input tables from Abaqus pipe-element ODB results.

The script is self-contained and runs with ``abaqus python`` without opening
Abaqus/CAE or Viewer. One tab-delimited text file is written for every selected
ODB step. Values are signed absolute envelopes across section-point and
selected pipe-element contributions at each node on the pipeline path.
"""

import argparse
import datetime
import heapq
import math
import os
import re
import sys
import traceback

from abaqusConstants import ELEMENT_NODAL, ON
from odbAccess import openOdb


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPT_VERSION = "2026-09-06-r2"
DEFAULT_INSTANCE = "PART-1-1"
DEFAULT_FRAME_INDEX = -1
DEFAULT_PRECISION = 6
OUTPUT_HEADERS = (
    "KP",
    "KP",
    "WD",
    "ESF1",
    "Axial Strain",
    "THE",
    "SK2",
    "SK1",
    "BMY",
    "BM-Z",
)
OUTPUT_UNITS = (
    "(m)",
    "(m)",
    "(m)",
    "(N)",
    "1",
    "",
    "1/m",
    "1/m",
    "N-m",
    "N-m",
)
RESULT_SPECS = (
    ("ESF1", (("ESF1", None), ("ESF", "ESF1"))),
    ("SE1", (("SE1", None), ("SE", "SE1"))),
    ("THE11", (("THE11", None), ("THE", "THE11"))),
    ("SK2", (("SK2", None), ("SK", "SK2"))),
    ("SK1", (("SK1", None), ("SK", "SK1"))),
    ("SM2", (("SM2", None), ("SM", "SM2"))),
    ("SM1", (("SM1", None), ("SM", "SM1"))),
)


def parse_arguments():
    parser = argparse.ArgumentParser(
        description=(
            "Write LCCDCC pipe-input text files containing KP, WD, ESF1, "
            "SE1, THE11, SK2, SK1, SM2, and SM1."
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
        help="Output directory (default: --input-dir).",
    )
    parser.add_argument(
        "--instance",
        default=DEFAULT_INSTANCE,
        help="Pipeline instance (default: PART-1-1).",
    )
    parser.add_argument(
        "--pipe-element-set",
        action="append",
        default=[],
        metavar="SET",
        help=(
            "Restrict output to this pipe element set. May be repeated; "
            "all pipe selections form a union."
        ),
    )
    parser.add_argument(
        "--pipe-element",
        action="append",
        nargs="+",
        type=int,
        default=[],
        metavar="LABEL",
        help=(
            "Restrict output to one or more exact pipe element labels. "
            "May be repeated."
        ),
    )
    parser.add_argument(
        "--pipe-element-range",
        action="append",
        nargs=2,
        type=int,
        default=[],
        metavar=("FIRST", "LAST"),
        help=(
            "Restrict output to an inclusive pipe element-label range. "
            "May be repeated."
        ),
    )
    parser.add_argument(
        "--start-node",
        type=int,
        default=None,
        help="Exact node label used as KP zero.",
    )
    parser.add_argument(
        "--end-node",
        type=int,
        default=None,
        help="Exact node label used as the end of the path.",
    )
    parser.add_argument(
        "--start-node-set",
        default=None,
        help=(
            "Exact instance node-set name for the start. If omitted, a "
            "unique set name containing START is detected."
        ),
    )
    parser.add_argument(
        "--end-node-set",
        default=None,
        help=(
            "Exact instance node-set name for the end. If omitted, a unique "
            "set name containing END is detected."
        ),
    )
    parser.add_argument(
        "--aslaid-step",
        default=None,
        metavar="STEP",
        help=(
            "Required for extraction: exact as-laid step name or 1-based "
            "step position. WD is COORD3 from its final frame."
        ),
    )
    parser.add_argument(
        "--frame-index",
        type=int,
        default=DEFAULT_FRAME_INDEX,
        help=(
            "Zero-based frame index used in every selected step. Use -1 "
            "for the final frame (default: -1)."
        ),
    )
    parser.add_argument(
        "--precision",
        type=int,
        default=DEFAULT_PRECISION,
        help=(
            "Digits after the decimal in scientific-format columns "
            "(default: 6)."
        ),
    )
    parser.add_argument(
        "--list-steps",
        action="store_true",
        help="List ordered steps and frame counts, then exit.",
    )
    parser.add_argument(
        "--list-pipe-element-sets",
        action="store_true",
        help=(
            "List only element sets containing pipe elements in --instance, "
            "then exit."
        ),
    )
    parser.add_argument(
        "--list-endpoint-sets",
        action="store_true",
        help=(
            "List instance node sets whose names contain START or END, then "
            "exit."
        ),
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Print short progress messages. Extraction is quiet by default.",
    )
    step_selection = parser.add_mutually_exclusive_group()
    step_selection.add_argument(
        "--steps",
        nargs="+",
        default=None,
        metavar="STEP",
        help="One or more exact step names (default: all ODB steps).",
    )
    step_selection.add_argument(
        "--step-range",
        nargs=2,
        default=None,
        metavar=("START", "END"),
        help=(
            "Inclusive range using 1-based step positions or exact first "
            "and last step names."
        ),
    )
    args = parser.parse_args()
    if args.frame_index < -1:
        parser.error("--frame-index must be -1 or zero or greater")
    if args.precision < 1 or args.precision > 12:
        parser.error("--precision must be between 1 and 12")
    for first_label, last_label in args.pipe_element_range:
        if first_label > last_label:
            parser.error(
                "--pipe-element-range FIRST must be less than or equal to LAST"
            )
    if not (
        args.list_steps
        or args.list_pipe_element_sets
        or args.list_endpoint_sets
    ) and args.aslaid_step is None:
        parser.error(
            "--aslaid-step is required for extraction; supply its exact "
            "name or 1-based step position"
        )
    return args


def repository_key(repository, requested_name):
    if requested_name in repository:
        return requested_name
    requested_upper = str(requested_name).upper()
    for key in repository.keys():
        if str(key).upper() == requested_upper:
            return key
    return None


def find_odb_files(input_dir, requested):
    input_dir = os.path.abspath(input_dir)
    if not os.path.isdir(input_dir):
        raise ValueError("Input directory does not exist: {0}".format(input_dir))
    if requested:
        paths = []
        for requested_path in requested:
            path = (
                requested_path
                if os.path.isabs(requested_path)
                else os.path.join(input_dir, requested_path)
            )
            path = os.path.abspath(path)
            if not os.path.isfile(path):
                raise ValueError("ODB file does not exist: {0}".format(path))
            if not path.lower().endswith(".odb"):
                raise ValueError("Not an ODB file: {0}".format(path))
            paths.append(path)
    else:
        paths = [
            os.path.join(input_dir, name)
            for name in os.listdir(input_dir)
            if name.lower().endswith(".odb")
            and os.path.isfile(os.path.join(input_dir, name))
        ]
    paths.sort(key=lambda value: value.lower())
    if not paths:
        raise ValueError("No ODB files found in: {0}".format(input_dir))
    return paths


def resolve_instance(odb, requested_name):
    key = repository_key(odb.rootAssembly.instances, requested_name)
    if key is None:
        raise ValueError(
            "Instance '{0}' was not found. Available instances: {1}".format(
                requested_name,
                ", ".join(odb.rootAssembly.instances.keys()),
            )
        )
    return key, odb.rootAssembly.instances[key]


def is_pipe_element(element):
    return str(element.type).upper().startswith("PIPE")


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


def physical_node_order(element):
    connectivity = tuple(int(label) for label in element.connectivity)
    if len(connectivity) == 2:
        return connectivity
    if len(connectivity) == 3:
        # Abaqus quadratic line elements list end nodes first and midside last.
        return (connectivity[0], connectivity[2], connectivity[1])
    return None


def distance_between(first, second):
    return math.sqrt(
        sum((first[index] - second[index]) ** 2 for index in range(3))
    )


def add_graph_edge(
    graph, edge_elements, first_label, second_label, length, element_label
):
    if first_label == second_label:
        return
    graph.setdefault(first_label, {})
    graph.setdefault(second_label, {})
    previous = graph[first_label].get(second_label)
    if previous is None or length < previous:
        graph[first_label][second_label] = length
        graph[second_label][first_label] = length
    edge_key = tuple(sorted((first_label, second_label)))
    edge_elements.setdefault(edge_key, set()).add(element_label)


def build_pipe_graph(instance, coordinates):
    graph = {}
    edge_elements = {}
    all_elements = {}
    pipe_elements = {}
    unsupported = []
    for element in instance.elements:
        label = int(element.label)
        all_elements[label] = element
        if not is_pipe_element(element):
            continue
        physical_order = physical_node_order(element)
        if physical_order is None:
            unsupported.append(label)
            continue
        pipe_elements[label] = element
        for index in range(len(physical_order) - 1):
            first_label = physical_order[index]
            second_label = physical_order[index + 1]
            if first_label not in coordinates or second_label not in coordinates:
                continue
            add_graph_edge(
                graph,
                edge_elements,
                first_label,
                second_label,
                distance_between(
                    coordinates[first_label], coordinates[second_label]
                ),
                label,
            )
    if not graph:
        raise ValueError(
            "No connected two- or three-node pipe elements were found in "
            "instance '{0}'.".format(instance.name)
        )
    return graph, edge_elements, all_elements, pipe_elements, unsupported


def nodes_from_instance_set(instance, set_name, graph):
    key = repository_key(instance.nodeSets, set_name)
    if key is None:
        raise ValueError(
            "Node set '{0}' was not found in instance '{1}'. Available node "
            "sets: {2}".format(
                set_name, instance.name, ", ".join(instance.nodeSets.keys())
            )
        )
    candidates = sorted(
        set(
            int(node.label)
            for node in instance.nodeSets[key].nodes
            if int(node.label) in graph
        )
    )
    return key, candidates


def single_endpoint(set_key, candidates, graph, role):
    if not candidates:
        raise ValueError(
            "{0} node set '{1}' contains no pipe-path node.".format(
                role.capitalize(), set_key
            )
        )
    if len(candidates) == 1:
        return candidates[0]
    endpoints = [label for label in candidates if len(graph[label]) == 1]
    if len(endpoints) == 1:
        return endpoints[0]
    raise ValueError(
        "{0} node set '{1}' contains several pipe-path nodes: {2}. Use "
        "--{0}-node to choose one.".format(
            role, set_key, ", ".join(str(label) for label in candidates)
        )
    )


def automatic_endpoint_set(instance, graph, token, role):
    matches = []
    for set_key in instance.nodeSets.keys():
        if token in str(set_key).upper():
            unused_key, candidates = nodes_from_instance_set(
                instance, set_key, graph
            )
            if candidates:
                matches.append((set_key, candidates))
    if not matches:
        return None, None
    if len(matches) > 1:
        raise ValueError(
            "More than one node set containing '{0}' has pipe-path nodes: "
            "{1}. Use --{2}-node-set with the exact name.".format(
                token,
                ", ".join(str(item[0]) for item in matches),
                role,
            )
        )
    set_key, candidates = matches[0]
    return (
        single_endpoint(set_key, candidates, graph, role),
        "auto-detected set '{0}'".format(set_key),
    )


def endpoint_from_options(
    instance, graph, node_label, set_name, token, role
):
    if node_label is not None:
        if node_label not in graph:
            raise ValueError(
                "{0} node {1} is not connected to a supported pipe "
                "element.".format(role.capitalize(), node_label)
            )
        return node_label, "explicit node {0}".format(node_label)
    if set_name:
        key, candidates = nodes_from_instance_set(instance, set_name, graph)
        return (
            single_endpoint(key, candidates, graph, role),
            "explicit set '{0}'".format(key),
        )
    return automatic_endpoint_set(instance, graph, token, role)


def shortest_distances(graph, start_label):
    distances = dict((label, float("inf")) for label in graph)
    distances[start_label] = 0.0
    queue = [(0.0, start_label)]
    while queue:
        distance, node_label = heapq.heappop(queue)
        if distance != distances[node_label]:
            continue
        for neighbor, edge_length in graph[node_label].items():
            candidate = distance + edge_length
            if candidate < distances[neighbor]:
                distances[neighbor] = candidate
                heapq.heappush(queue, (candidate, neighbor))
    return distances


def inferred_opposite_endpoint(graph, known_label, role):
    distances = shortest_distances(graph, known_label)
    candidates = [
        label
        for label in graph
        if label != known_label
        and len(graph[label]) == 1
        and not math.isinf(distances[label])
    ]
    if not candidates:
        raise ValueError(
            "No reachable pipe endpoint can be inferred for the {0}. "
            "Specify it explicitly.".format(role)
        )
    candidates.sort(key=lambda label: (-distances[label], label))
    return candidates[0]


def resolve_start_end(
    instance,
    graph,
    start_node,
    end_node,
    start_node_set,
    end_node_set,
):
    start_label, start_source = endpoint_from_options(
        instance, graph, start_node, start_node_set, "START", "start"
    )
    end_label, end_source = endpoint_from_options(
        instance, graph, end_node, end_node_set, "END", "end"
    )
    if start_label is None and end_label is None:
        endpoints = sorted(label for label in graph if len(graph[label]) == 1)
        if len(endpoints) != 2:
            raise ValueError(
                "No unique START/END sets were found and the pipe graph has "
                "{0} endpoints. Specify --start-node and --end-node.".format(
                    len(endpoints)
                )
            )
        distances = shortest_distances(graph, endpoints[0])
        if math.isinf(distances[endpoints[1]]):
            raise ValueError("The two pipe endpoints are not connected.")
        start_label, end_label = endpoints
        start_source = "automatically selected graph endpoint"
        end_source = "automatically selected graph endpoint"
    elif start_label is None:
        start_label = inferred_opposite_endpoint(graph, end_label, "start")
        start_source = "farthest reachable graph endpoint"
    elif end_label is None:
        end_label = inferred_opposite_endpoint(graph, start_label, "end")
        end_source = "farthest reachable graph endpoint"
    if start_label == end_label:
        raise ValueError("The start and end nodes must be different.")
    return start_label, end_label, start_source, end_source


def start_to_end_route(graph, start_label, end_label):
    distances = dict((label, float("inf")) for label in graph)
    previous = {}
    distances[start_label] = 0.0
    queue = [(0.0, start_label)]
    while queue:
        distance, node_label = heapq.heappop(queue)
        if distance != distances[node_label]:
            continue
        if node_label == end_label:
            break
        for neighbor, edge_length in graph[node_label].items():
            candidate = distance + edge_length
            if candidate < distances[neighbor]:
                distances[neighbor] = candidate
                previous[neighbor] = node_label
                heapq.heappush(queue, (candidate, neighbor))
    if math.isinf(distances[end_label]):
        raise ValueError(
            "No connected pipe path exists from node {0} to node {1}.".format(
                start_label, end_label
            )
        )
    route = [end_label]
    while route[-1] != start_label:
        route.append(previous[route[-1]])
    route.reverse()
    route_distances = {start_label: 0.0}
    for index in range(1, len(route)):
        route_distances[route[index]] = (
            route_distances[route[index - 1]]
            + graph[route[index - 1]][route[index]]
        )
    return route, route_distances


def route_element_labels(route, edge_elements):
    labels = set()
    for index in range(len(route) - 1):
        labels.update(
            edge_elements.get(tuple(sorted((route[index], route[index + 1]))), set())
        )
    return labels


def flattened_element_labels(container):
    if container is None:
        return
    if hasattr(container, "label") and hasattr(container, "connectivity"):
        yield int(container.label)
        return
    get_members = getattr(container, "getMemberFromAll", None)
    if get_members is not None:
        try:
            for label in get_members("label"):
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


def element_set_labels_for_instance(element_set, instance, instance_labels):
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
        if (
            outer_count == len(instance_names)
            and outer_count
            and not hasattr(first_member, "label")
        ):
            containers = [elements_member[index] for index in matching_indices]
        else:
            containers = [elements_member]
    else:
        containers = [element_set.elements]
    labels = set()
    for container in containers:
        for label in flattened_element_labels(container):
            if label in instance_labels:
                labels.add(label)
    return labels


def resolve_element_set(odb, instance, requested_name, instance_labels):
    key = repository_key(instance.elementSets, requested_name)
    if key is not None:
        return (
            element_set_labels_for_instance(
                instance.elementSets[key], instance, instance_labels
            ),
            "instance set '{0}'".format(key),
        )
    key = repository_key(odb.rootAssembly.elementSets, requested_name)
    if key is not None:
        return (
            element_set_labels_for_instance(
                odb.rootAssembly.elementSets[key], instance, instance_labels
            ),
            "assembly set '{0}'".format(key),
        )
    available = sorted(
        set(
            list(instance.elementSets.keys())
            + list(odb.rootAssembly.elementSets.keys())
        )
    )
    raise ValueError(
        "Pipe element set '{0}' was not found. Available element sets: "
        "{1}".format(
            requested_name, ", ".join(available) if available else "(none)"
        )
    )


def flattened_requested_labels(groups):
    labels = []
    for group in groups:
        labels.extend(group)
    return labels


def resolve_selected_pipe_elements(
    odb,
    instance,
    all_elements,
    pipe_elements,
    requested_sets,
    requested_elements,
    requested_ranges,
):
    exact_labels = flattened_requested_labels(requested_elements)
    if not (requested_sets or exact_labels or requested_ranges):
        return None, ["no pipe selection supplied; full route used"]
    instance_labels = set(all_elements.keys())
    pipe_labels = set(pipe_elements.keys())
    selected = set()
    notes = []
    for requested_name in requested_sets:
        labels, resolved_name = resolve_element_set(
            odb, instance, requested_name, instance_labels
        )
        matching = labels.intersection(pipe_labels)
        selected.update(matching)
        notes.append(
            "{0}: {1} pipe element(s); {2} non-pipe element(s) omitted".format(
                resolved_name, len(matching), len(labels) - len(matching)
            )
        )
    missing = sorted(set(exact_labels).difference(instance_labels))
    if missing:
        raise ValueError(
            "Requested pipe element labels are missing: {0}".format(
                ", ".join(str(label) for label in missing)
            )
        )
    wrong_type = sorted(
        set(exact_labels).intersection(instance_labels).difference(pipe_labels)
    )
    if wrong_type:
        raise ValueError(
            "These --pipe-element labels are not pipe elements: {0}".format(
                ", ".join(str(label) for label in wrong_type)
            )
        )
    selected.update(set(exact_labels).intersection(pipe_labels))
    if exact_labels:
        notes.append(
            "exact labels: {0} pipe element(s)".format(
                len(set(exact_labels).intersection(pipe_labels))
            )
        )
    for first_label, last_label in requested_ranges:
        existing = set(
            label
            for label in instance_labels
            if first_label <= label <= last_label
        )
        matching = existing.intersection(pipe_labels)
        selected.update(matching)
        notes.append(
            "range {0}-{1}: {2} pipe element(s); {3} non-pipe element(s) "
            "omitted".format(
                first_label,
                last_label,
                len(matching),
                len(existing) - len(matching),
            )
        )
    if not selected:
        raise ValueError("The requested selections contain no pipe elements.")
    return selected, notes


def find_step_name(available_steps, requested_name):
    requested_upper = str(requested_name).upper()
    for step_name in available_steps:
        if str(step_name).upper() == requested_upper:
            return step_name
    return None


def integer_value(text):
    try:
        return int(text)
    except ValueError:
        return None


def select_steps(available_steps, requested_steps, step_range):
    if requested_steps is not None:
        selected = []
        notes = []
        for requested in requested_steps:
            actual = find_step_name(available_steps, requested)
            if actual is None:
                notes.append("requested step '{0}' was not found".format(requested))
            elif actual not in selected:
                selected.append(actual)
        return selected, notes
    if step_range is None:
        return list(available_steps), []
    start_text, end_text = step_range
    start_number = integer_value(start_text)
    end_number = integer_value(end_text)
    if (start_number is None) != (end_number is None):
        raise ValueError(
            "--step-range endpoints must both be positions or both be names"
        )
    if start_number is not None:
        if start_number < 1 or end_number < 1:
            raise ValueError("Step positions are 1-based and must be positive.")
        if start_number > end_number:
            raise ValueError("The start of --step-range exceeds the end.")
        if start_number > len(available_steps):
            return [], [
                "range starts at {0}, but the ODB has only {1} steps".format(
                    start_number, len(available_steps)
                )
            ]
        actual_end = min(end_number, len(available_steps))
        notes = []
        if end_number > len(available_steps):
            notes.append(
                "range end {0} limited to step {1}".format(
                    end_number, len(available_steps)
                )
            )
        return available_steps[start_number - 1 : actual_end], notes
    start_name = find_step_name(available_steps, start_text)
    end_name = find_step_name(available_steps, end_text)
    missing = []
    if start_name is None:
        missing.append("start step '{0}' was not found".format(start_text))
    if end_name is None:
        missing.append("end step '{0}' was not found".format(end_text))
    if missing:
        return [], missing
    start_index = available_steps.index(start_name)
    end_index = available_steps.index(end_name)
    if start_index > end_index:
        raise ValueError("The start step occurs after the end step.")
    return available_steps[start_index : end_index + 1], []


def resolve_aslaid_step(available_steps, requested_step):
    step_number = integer_value(str(requested_step))
    if step_number is not None:
        if step_number < 1 or step_number > len(available_steps):
            raise ValueError(
                "--aslaid-step position {0} is outside the available range "
                "1-{1}.".format(step_number, len(available_steps))
            )
        return available_steps[step_number - 1]
    step_name = find_step_name(available_steps, requested_step)
    if step_name is None:
        raise ValueError(
            "As-laid step '{0}' was not found. Available steps: {1}".format(
                requested_step, ", ".join(available_steps)
            )
        )
    return step_name


def selected_frame(step, frame_index):
    if not step.frames:
        return None, None
    if frame_index == -1:
        index = len(step.frames) - 1
    else:
        if frame_index >= len(step.frames):
            return None, None
        index = frame_index
    return index, step.frames[index]


def field_value_data(value):
    try:
        data = value.data
    except Exception:
        data = value.dataDouble
    if isinstance(data, (int, float)):
        return (float(data),)
    return tuple(float(item) for item in data)


def component_index(field_output, requested_component):
    try:
        labels = list(field_output.componentLabels)
    except Exception:
        labels = []
    requested_upper = requested_component.upper()
    for index, label in enumerate(labels):
        if str(label).upper() == requested_upper:
            return index
    return None


def resolve_result_field(frame, alternatives):
    for field_name, component_name in alternatives:
        key = repository_key(frame.fieldOutputs, field_name)
        if key is None:
            continue
        field_output = frame.fieldOutputs[key]
        if component_name is None:
            index = component_index(field_output, field_name)
            return field_output, 0 if index is None else index, str(key)
        index = component_index(field_output, component_name)
        if index is not None:
            return (
                field_output,
                index,
                "{0}.{1}".format(key, component_name),
            )
        try:
            labels = list(field_output.componentLabels)
        except Exception:
            labels = []
        if not labels:
            return (
                field_output,
                0,
                "{0}.{1}".format(key, component_name),
            )
    return None, None, None


def signed_absolute_envelope(values):
    if not values:
        return None
    return max(values, key=lambda value: (abs(value), value))


def envelope_result_at_nodes(
    frame,
    alternatives,
    instance,
    output_nodes,
    output_element_labels,
):
    field_output, component, source = resolve_result_field(frame, alternatives)
    if field_output is None:
        return {}, None
    try:
        subset = field_output.getSubset(region=instance)
    except Exception:
        subset = field_output
    try:
        subset = subset.getSubset(position=ELEMENT_NODAL, readOnly=ON)
    except Exception:
        pass
    target_nodes = set(output_nodes)
    target_elements = set(output_element_labels)
    contributions = dict((label, []) for label in output_nodes)
    for value in subset.values:
        try:
            node_label = int(value.nodeLabel)
            element_label = int(value.elementLabel)
        except Exception:
            continue
        if node_label not in target_nodes or element_label not in target_elements:
            continue
        value_instance = getattr(value, "instance", None)
        if (
            value_instance is not None
            and str(value_instance.name).upper() != str(instance.name).upper()
        ):
            continue
        try:
            data = field_value_data(value)
        except Exception:
            continue
        if component >= len(data):
            continue
        scalar = data[component]
        if math.isnan(scalar) or math.isinf(scalar):
            continue
        contributions[node_label].append(scalar)
    enveloped = {}
    for node_label, values in contributions.items():
        envelope = signed_absolute_envelope(values)
        if envelope is not None:
            enveloped[node_label] = envelope
    return enveloped, source


def nodal_vectors(field_output, instance, requested_nodes):
    try:
        subset = field_output.getSubset(region=instance)
    except Exception:
        subset = field_output
    target_nodes = set(requested_nodes)
    contributions = dict((label, []) for label in requested_nodes)
    for value in subset.values:
        try:
            node_label = int(value.nodeLabel)
        except Exception:
            continue
        if node_label not in target_nodes:
            continue
        value_instance = getattr(value, "instance", None)
        if (
            value_instance is not None
            and str(value_instance.name).upper() != str(instance.name).upper()
        ):
            continue
        try:
            vector = field_value_data(value)
        except Exception:
            continue
        if len(vector) >= 3:
            contributions[node_label].append(vector)
    averaged = {}
    for node_label, vectors in contributions.items():
        if vectors:
            averaged[node_label] = tuple(
                sum(vector[index] for vector in vectors) / float(len(vectors))
                for index in range(3)
            )
    return averaged


def coord3_water_depths(frame, instance, output_nodes, aslaid_step_name):
    coord_key = repository_key(frame.fieldOutputs, "COORD")
    if coord_key is None:
        raise ValueError(
            "Final frame of as-laid step '{0}' does not contain COORD field "
            "output. WD requires COORD3.".format(aslaid_step_name)
        )
    coordinate_vectors = nodal_vectors(
        frame.fieldOutputs[coord_key], instance, output_nodes
    )
    missing = [
        node_label
        for node_label in output_nodes
        if node_label not in coordinate_vectors
    ]
    if missing:
        raise ValueError(
            "Final frame of as-laid step '{0}' has no COORD3 value for {1} "
            "output node(s), beginning with: {2}".format(
                aslaid_step_name,
                len(missing),
                ", ".join(str(label) for label in missing[:20]),
            )
        )
    return dict(
        (node_label, coordinate_vectors[node_label][2])
        for node_label in output_nodes
    )


def safe_step_filename(step_name):
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", str(step_name)).strip("._-")
    return value or "STEP"


def lccdcc_output_path(output_dir, odb_path, step_position, step_name):
    odb_stem = os.path.splitext(os.path.basename(odb_path))[0]
    filename = "{0}_Step{1:03d}_{2}_LCCDCC_INPUT.txt".format(
        odb_stem, step_position, safe_step_filename(step_name)
    )
    return os.path.abspath(os.path.join(output_dir, filename))


def scientific_text(value, precision):
    if value is None:
        return ""
    format_string = "{0:.%dE}" % precision
    return format_string.format(float(value))


def write_lccdcc_table(
    output_path,
    output_nodes,
    route_distances,
    water_depths,
    results,
    precision,
):
    with open(output_path, "w") as output_file:
        output_file.write("\t".join(OUTPUT_HEADERS) + "\n")
        output_file.write("\t".join(OUTPUT_UNITS) + "\n")
        for node_label in output_nodes:
            kp = route_distances[node_label]
            row = [
                "{0:.2f}".format(kp),
                scientific_text(kp, precision),
                scientific_text(water_depths.get(node_label), precision),
                scientific_text(results["ESF1"].get(node_label), precision),
                scientific_text(results["SE1"].get(node_label), precision),
                scientific_text(results["THE11"].get(node_label), precision),
                scientific_text(results["SK2"].get(node_label), precision),
                scientific_text(results["SK1"].get(node_label), precision),
                scientific_text(results["SM2"].get(node_label), precision),
                scientific_text(results["SM1"].get(node_label), precision),
            ]
            output_file.write("\t".join(row) + "\n")


def process_odb(odb_path, output_dir, args):
    odb = None
    messages = []
    output_paths = []
    try:
        if args.verbose:
            print("Opening: {0}".format(odb_path))
        odb = openOdb(path=odb_path, readOnly=True)
        instance_key, instance = resolve_instance(odb, args.instance)
        coordinates = node_coordinates(instance)
        (
            graph,
            edge_elements,
            all_elements,
            pipe_elements,
            unsupported,
        ) = build_pipe_graph(instance, coordinates)
        (
            start_label,
            end_label,
            start_source,
            end_source,
        ) = resolve_start_end(
            instance,
            graph,
            args.start_node,
            args.end_node,
            args.start_node_set,
            args.end_node_set,
        )
        route, route_distances = start_to_end_route(
            graph, start_label, end_label
        )
        route_elements = route_element_labels(route, edge_elements)
        selected, selection_notes = resolve_selected_pipe_elements(
            odb,
            instance,
            all_elements,
            pipe_elements,
            args.pipe_element_set,
            args.pipe_element,
            args.pipe_element_range,
        )
        if selected is None:
            output_element_labels = set(route_elements)
        else:
            output_element_labels = selected.intersection(route_elements)
            outside_count = len(selected) - len(output_element_labels)
            if outside_count:
                selection_notes.append(
                    "{0} selected pipe element(s) outside the start-to-end "
                    "route were omitted".format(outside_count)
                )
        if not output_element_labels:
            raise ValueError(
                "No selected pipe elements lie on the start-to-end path."
            )
        output_node_set = set()
        for element_label in output_element_labels:
            output_node_set.update(
                int(label)
                for label in pipe_elements[element_label].connectivity
                if int(label) in route_distances
            )
        output_nodes = [
            node_label for node_label in route if node_label in output_node_set
        ]
        available_steps = list(odb.steps.keys())
        aslaid_step_name = resolve_aslaid_step(
            available_steps, args.aslaid_step
        )
        aslaid_frames = odb.steps[aslaid_step_name].frames
        if not aslaid_frames:
            raise ValueError(
                "As-laid step '{0}' contains no frames.".format(
                    aslaid_step_name
                )
            )
        aslaid_frame_index = len(aslaid_frames) - 1
        aslaid_frame = aslaid_frames[aslaid_frame_index]
        water_depths = coord3_water_depths(
            aslaid_frame, instance, output_nodes, aslaid_step_name
        )
        selected_steps, step_notes = select_steps(
            available_steps, args.steps, args.step_range
        )
        if not selected_steps:
            raise ValueError(
                "No selected steps are available. Available steps: {0}".format(
                    ", ".join(available_steps)
                )
            )
        messages.append("ODB: {0}".format(odb_path))
        messages.append("Instance: {0}".format(instance_key))
        messages.append(
            "Path: node {0} ({1}) to node {2} ({3}); KP length={4:.12g}".format(
                start_label,
                start_source,
                end_label,
                end_source,
                route_distances[end_label],
            )
        )
        messages.append(
            "Output selection: {0} pipe element(s), {1} path node(s)".format(
                len(output_element_labels), len(output_nodes)
            )
        )
        messages.append(
            "WD: COORD3 from final frame {0} of as-laid step '{1}'; "
            "description={2}".format(
                aslaid_frame_index,
                aslaid_step_name,
                getattr(aslaid_frame, "description", ""),
            )
        )
        messages.append(
            "WD values: {0}/{1}".format(
                len(water_depths), len(output_nodes)
            )
        )
        if unsupported:
            messages.append(
                "Unsupported pipe connectivity omitted: {0} element(s)".format(
                    len(unsupported)
                )
            )
        messages.extend("Selection note: " + note for note in selection_notes)
        messages.extend("Step note: " + note for note in step_notes)
        for step_name in selected_steps:
            frame_index, frame = selected_frame(
                odb.steps[step_name], args.frame_index
            )
            if frame is None:
                messages.append(
                    "Skipped step '{0}': frame index {1} is unavailable".format(
                        step_name, args.frame_index
                    )
                )
                continue
            results = {}
            sources = {}
            for result_name, alternatives in RESULT_SPECS:
                values, source = envelope_result_at_nodes(
                    frame,
                    alternatives,
                    instance,
                    output_nodes,
                    output_element_labels,
                )
                results[result_name] = values
                sources[result_name] = source
            step_position = available_steps.index(step_name) + 1
            output_path = lccdcc_output_path(
                output_dir, odb_path, step_position, step_name
            )
            write_lccdcc_table(
                output_path,
                output_nodes,
                route_distances,
                water_depths,
                results,
                args.precision,
            )
            output_paths.append(output_path)
            messages.append(
                "Step: {0}; position={1}; frame={2}; description={3}".format(
                    step_name,
                    step_position,
                    frame_index,
                    getattr(frame, "description", ""),
                )
            )
            for result_name, unused_alternatives in RESULT_SPECS:
                messages.append(
                    "  {0}: source={1}; values={2}/{3}".format(
                        result_name,
                        sources[result_name] or "not found",
                        len(results[result_name]),
                        len(output_nodes),
                    )
                )
            messages.append("  Wrote: {0}".format(output_path))
            if args.verbose:
                print("Wrote: {0}".format(output_path))
        if not output_paths:
            raise ValueError("No LCCDCC input files were written.")
        return output_paths, messages
    finally:
        if odb is not None:
            odb.close()


def print_steps(odb_path):
    odb = None
    try:
        odb = openOdb(path=odb_path, readOnly=True)
        print(odb_path)
        for index, step_name in enumerate(odb.steps.keys(), 1):
            print(
                "  {0:>4}  {1}  frames={2}".format(
                    index, step_name, len(odb.steps[step_name].frames)
                )
            )
    finally:
        if odb is not None:
            odb.close()


def print_pipe_element_sets(odb_path, instance_name):
    odb = None
    try:
        odb = openOdb(path=odb_path, readOnly=True)
        unused_key, instance = resolve_instance(odb, instance_name)
        element_by_label = dict(
            (int(element.label), element) for element in instance.elements
        )
        instance_labels = set(element_by_label.keys())
        pipe_labels = set(
            label
            for label, element in element_by_label.items()
            if is_pipe_element(element)
        )
        print(odb_path)
        listed = 0
        for scope, repository in (
            ("instance", instance.elementSets),
            ("assembly", odb.rootAssembly.elementSets),
        ):
            for set_name in sorted(repository.keys()):
                labels = element_set_labels_for_instance(
                    repository[set_name], instance, instance_labels
                ).intersection(pipe_labels)
                if labels:
                    print(
                        "  {0:<9} {1}  pipe elements={2}".format(
                            scope, set_name, len(labels)
                        )
                    )
                    listed += 1
        if not listed:
            print("  (no element sets contain pipe elements)")
    finally:
        if odb is not None:
            odb.close()


def print_endpoint_sets(odb_path, instance_name):
    odb = None
    try:
        odb = openOdb(path=odb_path, readOnly=True)
        unused_key, instance = resolve_instance(odb, instance_name)
        print(odb_path)
        listed = 0
        for set_name in sorted(instance.nodeSets.keys()):
            upper_name = str(set_name).upper()
            if "START" not in upper_name and "END" not in upper_name:
                continue
            labels = sorted(
                set(int(node.label) for node in instance.nodeSets[set_name].nodes)
            )
            print(
                "  {0}  nodes={1}  labels={2}".format(
                    set_name,
                    len(labels),
                    ",".join(str(label) for label in labels[:20]),
                )
            )
            listed += 1
        if not listed:
            print("  (no node-set names contain START or END)")
    finally:
        if odb is not None:
            odb.close()


def write_log(log_path, input_dir, output_dir, successes, failures):
    with open(log_path, "w") as log_file:
        log_file.write(
            "extract_lccdcc_pipe_input.py version {0}\n".format(SCRIPT_VERSION)
        )
        log_file.write("Time: {0}\n".format(datetime.datetime.now()))
        log_file.write("sys.argv: {0}\n".format(repr(sys.argv)))
        log_file.write("Input directory: {0}\n".format(input_dir))
        log_file.write("Output directory: {0}\n".format(output_dir))
        log_file.write("Envelope: signed maximum absolute value\n")
        log_file.write("Successful ODBs: {0}\n".format(len(successes)))
        log_file.write("Failed ODBs: {0}\n".format(len(failures)))
        for odb_path, output_paths, messages in successes:
            log_file.write("\n" + "=" * 100 + "\n")
            for message in messages:
                log_file.write(message + "\n")
        for odb_path, message, traceback_text in failures:
            log_file.write("\n" + "=" * 100 + "\n")
            log_file.write("FAILED ODB: {0}\n".format(odb_path))
            log_file.write("Error: {0}\n\n".format(message))
            log_file.write(traceback_text)


def main():
    args = parse_arguments()
    if args.verbose:
        print(
            "extract_lccdcc_pipe_input.py version {0}".format(SCRIPT_VERSION)
        )
    input_dir = os.path.abspath(args.input_dir)
    odb_paths = find_odb_files(input_dir, args.odb)
    if (
        args.list_steps
        or args.list_pipe_element_sets
        or args.list_endpoint_sets
    ):
        for odb_path in odb_paths:
            if args.list_steps:
                print_steps(odb_path)
            if args.list_pipe_element_sets:
                print_pipe_element_sets(odb_path, args.instance)
            if args.list_endpoint_sets:
                print_endpoint_sets(odb_path, args.instance)
        return 0
    output_dir = os.path.abspath(args.output_dir or input_dir)
    if not os.path.isdir(output_dir):
        os.makedirs(output_dir)
    successes = []
    failures = []
    for odb_path in odb_paths:
        try:
            output_paths, messages = process_odb(odb_path, output_dir, args)
            successes.append((odb_path, output_paths, messages))
        except Exception as exc:
            failures.append((odb_path, str(exc), traceback.format_exc()))
            if args.verbose:
                print("FAILED: {0}".format(odb_path))
    log_path = os.path.join(output_dir, "extract_lccdcc_pipe_input.log")
    write_log(log_path, input_dir, output_dir, successes, failures)
    if args.verbose:
        print(
            "Completed: {0} ODB(s) succeeded, {1} failed.".format(
                len(successes), len(failures)
            )
        )
        print("Log: {0}".format(log_path))
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
