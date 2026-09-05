from __future__ import print_function

"""Extract ESF1 and original nodal Z along a PIPE31H start-to-end path.

The script is self-contained and runs with ``abaqus python``. By default it
processes every step and every PIPE31H element on the resolved start-to-end
route. Element sets, exact element labels, and inclusive label ranges may be
used to restrict output without resetting the full-route path distance.
"""

import argparse
import datetime
import heapq
import math
import os
import sys
import traceback

from abaqusConstants import ELEMENT_NODAL, ON
from odbAccess import openOdb


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPT_VERSION = "2026-09-05-r1"
DEFAULT_INSTANCE = "PART-1-1"
DEFAULT_ELEMENT_TYPE = "PIPE31H"
DEFAULT_FRAME_INDEX = -1


def parse_arguments():
    parser = argparse.ArgumentParser(
        description=(
            "Write ESF1 and water depth (original nodal Z) along the "
            "PIPE31H path from a start node to an end node."
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
        help="Report directory (default: --input-dir).",
    )
    parser.add_argument(
        "--output-name",
        default=None,
        help="Custom report filename for a single --odb run only.",
    )
    parser.add_argument(
        "--instance",
        default=DEFAULT_INSTANCE,
        help="Pipeline instance (default: PART-1-1).",
    )
    parser.add_argument(
        "--element-set",
        action="append",
        default=[],
        metavar="SET",
        help=(
            "Restrict output to this instance- or assembly-level element "
            "set. May be repeated; selections form a union."
        ),
    )
    parser.add_argument(
        "--element",
        action="append",
        nargs="+",
        type=int,
        default=[],
        metavar="LABEL",
        help=(
            "Restrict output to one or more exact element labels. May be "
            "repeated; selections form a union."
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
            "Restrict output to an inclusive element-label range. May be "
            "repeated; selections form a union."
        ),
    )
    parser.add_argument(
        "--start-node",
        type=int,
        default=None,
        help="Exact node label used as path-distance zero.",
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
            "unique set name containing START is detected automatically."
        ),
    )
    parser.add_argument(
        "--end-node-set",
        default=None,
        help=(
            "Exact instance node-set name for the end. If omitted, a unique "
            "set name containing END is detected automatically."
        ),
    )
    parser.add_argument(
        "--frame-index",
        type=int,
        default=DEFAULT_FRAME_INDEX,
        help=(
            "Zero-based frame index. Use -1 for the last frame containing "
            "ESF1 (default: -1)."
        ),
    )
    parser.add_argument(
        "--list-steps",
        action="store_true",
        help="List each ODB's ordered steps and exit without writing reports.",
    )
    parser.add_argument(
        "--list-element-sets",
        action="store_true",
        help=(
            "List only element sets containing PIPE31H elements in "
            "--instance, then exit without writing reports."
        ),
    )
    parser.add_argument(
        "--list-endpoint-sets",
        action="store_true",
        help=(
            "List instance node sets whose names contain START or END, then "
            "exit without writing reports."
        ),
    )

    step_selection = parser.add_mutually_exclusive_group()
    step_selection.add_argument(
        "--steps",
        nargs="+",
        default=None,
        metavar="STEP",
        help="One or more exact step names to include.",
    )
    step_selection.add_argument(
        "--step-range",
        nargs=2,
        default=None,
        metavar=("START", "END"),
        help=(
            "Inclusive range using 1-based step positions or exact first and "
            "last step names."
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


def repository_key(repository, requested_name):
    if requested_name in repository:
        return requested_name
    requested_upper = requested_name.upper()
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


def default_report_name(odb_path):
    stem = os.path.splitext(os.path.basename(odb_path))[0]
    return stem + "_PIPE31H_ESF1_WATER_DEPTH_PATH.rpt"


def build_jobs(odb_paths, output_dir, output_name):
    jobs = []
    paths_seen = set()
    for odb_path in odb_paths:
        report_path = os.path.abspath(
            os.path.join(
                output_dir, output_name or default_report_name(odb_path)
            )
        )
        key = os.path.normcase(report_path)
        if key in paths_seen:
            raise ValueError(
                "More than one ODB maps to report path: {0}".format(
                    report_path
                )
            )
        paths_seen.add(key)
        jobs.append((odb_path, report_path))
    return jobs


def resolve_instance(odb, requested_name):
    instance_key = repository_key(
        odb.rootAssembly.instances, requested_name
    )
    if instance_key is None:
        raise ValueError(
            "Instance '{0}' was not found. Available instances: {1}".format(
                requested_name,
                ", ".join(odb.rootAssembly.instances.keys()),
            )
        )
    return instance_key, odb.rootAssembly.instances[instance_key]


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
    return math.sqrt(
        sum((first[index] - second[index]) ** 2 for index in range(3))
    )


def physical_node_order(element):
    connectivity = tuple(int(label) for label in element.connectivity)
    if len(connectivity) == 2:
        return connectivity
    if len(connectivity) == 3:
        # Abaqus quadratic line elements list the two end nodes first and the
        # midside node third.
        return (connectivity[0], connectivity[2], connectivity[1])
    return None


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


def build_pipe31h_graph(instance, coordinates):
    graph = {}
    edge_elements = {}
    all_elements = {}
    pipe_elements = {}
    unsupported = []
    for element in instance.elements:
        element_label = int(element.label)
        all_elements[element_label] = element
        if str(element.type).upper() != DEFAULT_ELEMENT_TYPE:
            continue
        physical_order = physical_node_order(element)
        if physical_order is None:
            unsupported.append(element_label)
            continue
        pipe_elements[element_label] = element
        for index in range(len(physical_order) - 1):
            first_label = physical_order[index]
            second_label = physical_order[index + 1]
            if (
                first_label not in coordinates
                or second_label not in coordinates
            ):
                continue
            add_graph_edge(
                graph,
                edge_elements,
                first_label,
                second_label,
                distance_between(
                    coordinates[first_label], coordinates[second_label]
                ),
                element_label,
            )
    if not graph:
        raise ValueError(
            "No connected two- or three-node PIPE31H elements were found in "
            "instance '{0}'.".format(instance.name)
        )
    return graph, edge_elements, all_elements, pipe_elements, unsupported


def nodes_from_instance_set(instance, set_name, graph):
    set_key = repository_key(instance.nodeSets, set_name)
    if set_key is None:
        raise ValueError(
            "Node set '{0}' was not found in instance '{1}'. Available "
            "sets: {2}".format(
                set_name,
                instance.name,
                ", ".join(instance.nodeSets.keys()),
            )
        )
    candidates = sorted(
        set(
            int(node.label)
            for node in instance.nodeSets[set_key].nodes
            if int(node.label) in graph
        )
    )
    return set_key, candidates


def single_endpoint_from_candidates(set_key, candidates, graph, role):
    if not candidates:
        raise ValueError(
            "{0} node set '{1}' contains no node on the PIPE31H graph.".format(
                role.capitalize(), set_key
            )
        )
    if len(candidates) == 1:
        return candidates[0]
    endpoints = [label for label in candidates if len(graph[label]) == 1]
    if len(endpoints) == 1:
        return endpoints[0]
    raise ValueError(
        "{0} node set '{1}' contains multiple PIPE31H path nodes: {2}. "
        "Use --{0}-node to select one label.".format(
            role,
            set_key,
            ", ".join(str(label) for label in candidates),
        )
    )


def auto_endpoint_set(instance, graph, token, role):
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
            "More than one node set containing '{0}' has PIPE31H path "
            "nodes: {1}. Use --{2}-node-set with the exact set name.".format(
                token,
                ", ".join(str(item[0]) for item in matches),
                role,
            )
        )
    set_key, candidates = matches[0]
    node_label = single_endpoint_from_candidates(
        set_key, candidates, graph, role
    )
    return node_label, "auto-detected set '{0}'".format(set_key)


def endpoint_from_options(
    instance, graph, node_label, set_name, token, role
):
    if node_label is not None:
        if node_label not in graph:
            raise ValueError(
                "{0} node {1} is not connected to a supported PIPE31H "
                "element.".format(role.capitalize(), node_label)
            )
        return node_label, "explicit node {0}".format(node_label)
    if set_name:
        set_key, candidates = nodes_from_instance_set(
            instance, set_name, graph
        )
        return (
            single_endpoint_from_candidates(
                set_key, candidates, graph, role
            ),
            "explicit set '{0}'".format(set_key),
        )
    return auto_endpoint_set(instance, graph, token, role)


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
            "No reachable PIPE31H endpoint can be inferred for the {0}. "
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
        instance,
        graph,
        start_node,
        start_node_set,
        "START",
        "start",
    )
    end_label, end_source = endpoint_from_options(
        instance,
        graph,
        end_node,
        end_node_set,
        "END",
        "end",
    )

    if start_label is None and end_label is None:
        endpoints = sorted(
            label for label in graph if len(graph[label]) == 1
        )
        if len(endpoints) != 2:
            raise ValueError(
                "No unique START/END node sets were found and the PIPE31H "
                "graph has {0} endpoints. Specify --start-node and "
                "--end-node (or their node-set options).".format(
                    len(endpoints)
                )
            )
        distances = shortest_distances(graph, endpoints[0])
        if math.isinf(distances[endpoints[1]]):
            raise ValueError(
                "The two automatically found PIPE31H endpoints are not "
                "connected. Specify start and end nodes on one path."
            )
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
        raise ValueError("The start node and end node must be different.")
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
            "No connected PIPE31H path exists from node {0} to node "
            "{1}.".format(start_label, end_label)
        )
    route = [end_label]
    while route[-1] != start_label:
        route.append(previous[route[-1]])
    route.reverse()
    route_distances = dict(
        (node_label, distances[node_label]) for node_label in route
    )
    return route, route_distances


def route_element_labels(route, edge_elements):
    labels = set()
    for index in range(len(route) - 1):
        edge_key = tuple(sorted((route[index], route[index + 1])))
        labels.update(edge_elements.get(edge_key, ()))
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
            selected_containers = [
                elements_member[index] for index in matching_indices
            ]
        else:
            selected_containers = [elements_member]
    else:
        selected_containers = [element_set.elements]
    labels = set()
    for container in selected_containers:
        for label in flattened_element_labels(container):
            if label in instance_labels:
                labels.add(label)
    return labels


def resolve_element_set(odb, instance, requested_name, instance_labels):
    set_key = repository_key(instance.elementSets, requested_name)
    if set_key is not None:
        return (
            element_set_labels_for_instance(
                instance.elementSets[set_key], instance, instance_labels
            ),
            "instance set '{0}'".format(set_key),
        )
    set_key = repository_key(
        odb.rootAssembly.elementSets, requested_name
    )
    if set_key is not None:
        return (
            element_set_labels_for_instance(
                odb.rootAssembly.elementSets[set_key],
                instance,
                instance_labels,
            ),
            "assembly set '{0}'".format(set_key),
        )
    available = sorted(
        set(
            list(instance.elementSets.keys())
            + list(odb.rootAssembly.elementSets.keys())
        )
    )
    raise ValueError(
        "Element set '{0}' was not found. Available sets: {1}".format(
            requested_name, ", ".join(available) if available else "(none)"
        )
    )


def flatten_requested_elements(requested_elements):
    labels = []
    for group in requested_elements:
        labels.extend(group)
    return labels


def resolve_requested_pipe_elements(
    odb,
    instance,
    all_elements,
    pipe_elements,
    requested_sets,
    requested_elements,
    requested_ranges,
):
    exact_labels = flatten_requested_elements(requested_elements)
    selection_requested = bool(
        requested_sets or exact_labels or requested_ranges
    )
    if not selection_requested:
        return None, [
            "no element filter requested; all PIPE31H route elements selected"
        ]

    instance_labels = set(all_elements.keys())
    requested_labels = set()
    notes = []
    for requested_name in requested_sets:
        labels, resolved_name = resolve_element_set(
            odb, instance, requested_name, instance_labels
        )
        requested_labels.update(labels)
        notes.append(
            "{0}: {1} element(s) found".format(resolved_name, len(labels))
        )

    missing = []
    for label in exact_labels:
        if label in instance_labels:
            requested_labels.add(label)
        else:
            missing.append(label)
    if missing:
        raise ValueError(
            "Requested element labels were not found in instance '{0}': "
            "{1}".format(
                instance.name,
                ", ".join(str(label) for label in sorted(set(missing))),
            )
        )
    if exact_labels:
        notes.append(
            "exact labels: {0} element(s) found".format(
                len(set(exact_labels))
            )
        )

    for first_label, last_label in requested_ranges:
        labels = set(
            label
            for label in instance_labels
            if first_label <= label <= last_label
        )
        requested_labels.update(labels)
        notes.append(
            "range {0}-{1}: {2} element(s) found".format(
                first_label, last_label, len(labels)
            )
        )

    selected_pipe_labels = requested_labels.intersection(pipe_elements)
    ignored_count = len(requested_labels) - len(selected_pipe_labels)
    if ignored_count:
        notes.append(
            "{0} requested non-PIPE31H element(s) ignored".format(
                ignored_count
            )
        )
    if not selected_pipe_labels:
        raise ValueError(
            "The requested element selections contain no PIPE31H elements "
            "in instance '{0}'.".format(instance.name)
        )
    return selected_pipe_labels, notes


def find_step_name(available_steps, requested_name):
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


def select_steps(available_steps, requested_steps, step_range):
    if requested_steps is not None:
        selected = []
        notes = []
        for requested_name in requested_steps:
            actual_name = find_step_name(available_steps, requested_name)
            if actual_name is None:
                notes.append(
                    "requested step '{0}' is not present".format(
                        requested_name
                    )
                )
            elif actual_name not in selected:
                selected.append(actual_name)
        return selected, notes
    if step_range is None:
        return list(available_steps), []

    start_text, end_text = step_range
    start_number = integer_value(start_text)
    end_number = integer_value(end_text)
    if (start_number is None) != (end_number is None):
        raise ValueError(
            "--step-range endpoints must both be step numbers or both be "
            "step names"
        )
    if start_number is not None:
        if start_number < 1 or end_number < 1:
            raise ValueError(
                "numeric step ranges are 1-based and must be positive"
            )
        if start_number > end_number:
            raise ValueError(
                "the start of --step-range must not exceed the end"
            )
        if start_number > len(available_steps):
            return [], [
                "range starts at {0}, but the ODB has only {1} step(s)".format(
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


def select_frames(odb, selected_steps, frame_index):
    selected = []
    skipped = []
    for step_name in selected_steps:
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
                    "frame index {0} is unavailable; step has {1} "
                    "frame(s)".format(frame_index, len(frames)),
                )
            )
            continue
        else:
            candidate_indices = (frame_index,)
        selection = None
        for candidate_index in candidate_indices:
            frame = frames[candidate_index]
            if "ESF1" in frame.fieldOutputs:
                selection = (step_name, candidate_index, frame)
                break
        if selection is None:
            skipped.append(
                (step_name, "no requested frame contains ESF1")
            )
        else:
            selected.append(selection)
    return selected, skipped


def field_value_scalar(value):
    try:
        data = value.data
    except Exception:
        data = value.dataDouble
    if isinstance(data, (int, float)):
        return float(data)
    values = tuple(float(item) for item in data)
    if not values:
        raise ValueError("An ESF1 field value contains no data.")
    return values[0]


def average_esf1_at_nodes(
    field_output, instance_name, output_nodes, output_element_labels
):
    target_nodes = set(output_nodes)
    contributions = dict((label, []) for label in output_nodes)
    element_nodal = field_output.getSubset(
        position=ELEMENT_NODAL, readOnly=ON
    )
    for value in element_nodal.values:
        try:
            node_label = int(value.nodeLabel)
            element_label = int(value.elementLabel)
        except Exception:
            continue
        if (
            node_label not in target_nodes
            or element_label not in output_element_labels
        ):
            continue
        value_instance = getattr(value, "instance", None)
        if (
            value_instance is not None
            and value_instance.name.upper() != instance_name.upper()
        ):
            continue
        scalar = field_value_scalar(value)
        if math.isnan(scalar) or math.isinf(scalar):
            continue
        contributions[node_label].append(scalar)

    values_by_node = {}
    counts_by_node = {}
    for node_label in output_nodes:
        values = contributions[node_label]
        if values:
            values_by_node[node_label] = sum(values) / float(len(values))
            counts_by_node[node_label] = len(values)
    return values_by_node, counts_by_node


def engineering_format(value):
    if value is None:
        return ""
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
        digits_before_decimal = int(
            math.floor(math.log10(abs(mantissa)))
        ) + 1
        decimal_places = max(0, 6 - digits_before_decimal)
        mantissa_text = ("{0:.%df}" % decimal_places).format(mantissa)
    return "{0}E{1:+03d}".format(mantissa_text, exponent)


def write_table(
    report_file,
    output_nodes,
    route_distances,
    coordinates,
    extracted_frames,
):
    fixed_widths = (18, 14, 18)
    step_width = 20
    header = (
        "{0:>{w0}}{1:>{w1}}{2:>{w2}}".format(
            "Path Distance",
            "Node Label",
            "Water Depth Z",
            w0=fixed_widths[0],
            w1=fixed_widths[1],
            w2=fixed_widths[2],
        )
        + "".join(
            "{0:>{width}}".format(step_name + " ESF1", width=step_width)
            for step_name, unused_index, unused_frame, unused_values, unused_counts
            in extracted_frames
        )
    )
    report_file.write(header + "\n")
    report_file.write("-" * len(header) + "\n")
    for node_label in output_nodes:
        line = (
            "{0:>{w0}}{1:>{w1}d}{2:>{w2}}".format(
                engineering_format(route_distances[node_label]),
                node_label,
                engineering_format(coordinates[node_label][2]),
                w0=fixed_widths[0],
                w1=fixed_widths[1],
                w2=fixed_widths[2],
            )
        )
        for (
            unused_step,
            unused_index,
            unused_frame,
            values_by_node,
            unused_counts,
        ) in extracted_frames:
            line += "{0:>{width}}".format(
                engineering_format(values_by_node.get(node_label)),
                width=step_width,
            )
        report_file.write(line + "\n")
    report_file.write("\n")


def write_summaries(report_file, route_distances, extracted_frames):
    for (
        step_name,
        frame_index,
        unused_frame,
        values_by_node,
        unused_counts,
    ) in extracted_frames:
        if not values_by_node:
            report_file.write(
                "{0} (frame {1}): no ESF1 values available\n".format(
                    step_name, frame_index
                )
            )
            continue
        minimum_node = min(
            values_by_node, key=lambda label: values_by_node[label]
        )
        maximum_node = max(
            values_by_node, key=lambda label: values_by_node[label]
        )
        report_file.write(
            "{0} (frame {1}) minimum ESF1: {2} at distance {3}, node "
            "{4}\n".format(
                step_name,
                frame_index,
                engineering_format(values_by_node[minimum_node]),
                engineering_format(route_distances[minimum_node]),
                minimum_node,
            )
        )
        report_file.write(
            "{0} (frame {1}) maximum ESF1: {2} at distance {3}, node "
            "{4}\n".format(
                step_name,
                frame_index,
                engineering_format(values_by_node[maximum_node]),
                engineering_format(route_distances[maximum_node]),
                maximum_node,
            )
        )


def write_report(
    odb_path,
    report_path,
    instance_name,
    requested_sets,
    requested_elements,
    requested_ranges,
    requested_steps,
    step_range,
    frame_index,
    start_node,
    end_node,
    start_node_set,
    end_node_set,
):
    odb = None
    try:
        print("Opening: {0}".format(odb_path))
        odb = openOdb(path=odb_path, readOnly=True)
        instance_key, instance = resolve_instance(odb, instance_name)
        coordinates = node_coordinates(instance)
        (
            graph,
            edge_elements,
            all_elements,
            pipe_elements,
            unsupported_elements,
        ) = build_pipe31h_graph(instance, coordinates)
        (
            start_label,
            end_label,
            start_source,
            end_source,
        ) = resolve_start_end(
            instance,
            graph,
            start_node,
            end_node,
            start_node_set,
            end_node_set,
        )
        route, route_distances = start_to_end_route(
            graph, start_label, end_label
        )
        elements_on_route = route_element_labels(route, edge_elements)
        requested_pipe_labels, element_notes = resolve_requested_pipe_elements(
            odb,
            instance,
            all_elements,
            pipe_elements,
            requested_sets,
            requested_elements,
            requested_ranges,
        )
        if requested_pipe_labels is None:
            output_element_labels = set(elements_on_route)
        else:
            output_element_labels = requested_pipe_labels.intersection(
                elements_on_route
            )
            off_route_count = len(requested_pipe_labels) - len(
                output_element_labels
            )
            if off_route_count:
                element_notes.append(
                    "{0} selected PIPE31H element(s) are outside the "
                    "start-to-end route and were omitted".format(
                        off_route_count
                    )
                )
        if not output_element_labels:
            raise ValueError(
                "No selected PIPE31H elements lie on the path from node {0} "
                "to node {1}.".format(start_label, end_label)
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
        selected_steps, step_notes = select_steps(
            available_steps, requested_steps, step_range
        )
        if not selected_steps:
            raise ValueError(
                "No requested steps are available. Available steps: {0}".format(
                    ", ".join(available_steps)
                )
            )
        selected_frames, skipped_steps = select_frames(
            odb, selected_steps, frame_index
        )
        if not selected_frames:
            details = "; ".join(
                "{0}: {1}".format(step_name, reason)
                for step_name, reason in skipped_steps
            )
            raise ValueError("No usable ESF1 frames were found. {0}".format(details))

        extracted_frames = []
        for step_name, selected_frame_index, frame in selected_frames:
            values_by_node, counts_by_node = average_esf1_at_nodes(
                frame.fieldOutputs["ESF1"],
                instance_key,
                output_nodes,
                output_element_labels,
            )
            extracted_frames.append(
                (
                    step_name,
                    selected_frame_index,
                    frame,
                    values_by_node,
                    counts_by_node,
                )
            )

        with open(report_path, "w") as report_file:
            report_file.write("*" * 112 + "\n")
            report_file.write(
                "PIPE31H ESF1 and Water Depth Along Path, written {0}\n".format(
                    datetime.datetime.now().strftime(
                        "%a %b %d %H:%M:%S %Y"
                    )
                )
            )
            report_file.write("Script version: {0}\n".format(SCRIPT_VERSION))
            report_file.write(
                "ODB: {0}\n".format(odb_path.replace("\\", "/"))
            )
            report_file.write("Instance: {0}\n".format(instance_key))
            report_file.write(
                "Default path element type: {0}\n".format(
                    DEFAULT_ELEMENT_TYPE
                )
            )
            report_file.write(
                "Start node: {0} ({1})\n".format(
                    start_label, start_source
                )
            )
            report_file.write(
                "End node: {0} ({1})\n".format(end_label, end_source)
            )
            report_file.write(
                "Full start-to-end route: {0} node(s), {1}\n".format(
                    len(route),
                    engineering_format(route_distances[end_label]),
                )
            )
            report_file.write(
                "Output: {0} PIPE31H element(s), {1} route node(s)\n".format(
                    len(output_element_labels), len(output_nodes)
                )
            )
            report_file.write(
                "Path distance: cumulative 3D distance using original ODB "
                "nodal coordinates\n"
            )
            report_file.write(
                "Water Depth Z: original ODB nodal coordinate Z; sign and "
                "model units are unchanged\n"
            )
            report_file.write(
                "ESF1: element-nodal values averaged using output elements "
                "only\n"
            )
            if unsupported_elements:
                report_file.write(
                    "Path note: {0} PIPE31H element(s) with unsupported "
                    "connectivity were omitted\n".format(
                        len(unsupported_elements)
                    )
                )
            for note in element_notes:
                report_file.write("Element note: {0}\n".format(note))
            for note in step_notes:
                report_file.write("Step note: {0}\n".format(note))
            for step_name, selected_frame_index, frame in selected_frames:
                report_file.write(
                    "Frame: {0}, index {1}, {2}\n".format(
                        step_name, selected_frame_index, frame.description
                    )
                )
            report_file.write("\n")
            write_table(
                report_file,
                output_nodes,
                route_distances,
                coordinates,
                extracted_frames,
            )
            write_summaries(
                report_file, route_distances, extracted_frames
            )

        print(
            "Path: node {0} to node {1}, distance {2}".format(
                start_label,
                end_label,
                engineering_format(route_distances[end_label]),
            )
        )
        print(
            "Output: {0} PIPE31H element(s), {1} path node(s)".format(
                len(output_element_labels), len(output_nodes)
            )
        )
        for note in element_notes:
            print("Element note: {0}.".format(note))
        for note in step_notes:
            print("Step note: {0}.".format(note))
        for skipped_step, reason in skipped_steps:
            print("Skipped step '{0}': {1}.".format(skipped_step, reason))
        missing_total = sum(
            len(output_nodes) - len(values_by_node)
            for unused_step, unused_index, unused_frame, values_by_node, unused_counts
            in extracted_frames
        )
        if missing_total:
            print(
                "Warning: {0} step/node ESF1 cell(s) are blank because no "
                "selected element-nodal value was available.".format(
                    missing_total
                )
            )
        print("Wrote: {0}".format(report_path))
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
        instance_key, instance = resolve_instance(odb, instance_name)
        instance_labels = set()
        pipe_labels = set()
        for element in instance.elements:
            label = int(element.label)
            instance_labels.add(label)
            if str(element.type).upper() == DEFAULT_ELEMENT_TYPE:
                pipe_labels.add(label)
        print(odb_path)
        print("  Instance: {0}".format(instance_key))
        print("  Total PIPE31H elements: {0}".format(len(pipe_labels)))
        listed = 0
        for set_name in sorted(instance.elementSets.keys()):
            labels = element_set_labels_for_instance(
                instance.elementSets[set_name], instance, instance_labels
            ).intersection(pipe_labels)
            if labels:
                print(
                    "  {0:>8}  {1}  PIPE31H elements={2}".format(
                        "instance", set_name, len(labels)
                    )
                )
                listed += 1
        for set_name in sorted(odb.rootAssembly.elementSets.keys()):
            labels = element_set_labels_for_instance(
                odb.rootAssembly.elementSets[set_name],
                instance,
                instance_labels,
            ).intersection(pipe_labels)
            if labels:
                print(
                    "  {0:>8}  {1}  PIPE31H elements={2}".format(
                        "assembly", set_name, len(labels)
                    )
                )
                listed += 1
        if not listed:
            print("  (no element sets contain PIPE31H elements)")
    finally:
        if odb is not None:
            odb.close()


def print_endpoint_sets(odb_path, instance_name):
    odb = None
    try:
        odb = openOdb(path=odb_path, readOnly=True)
        instance_key, instance = resolve_instance(odb, instance_name)
        print(odb_path)
        print("  Instance: {0}".format(instance_key))
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


def write_execution_log(log_path, input_dir, output_dir, jobs, failures):
    with open(log_path, "w") as log_file:
        log_file.write(
            "extract_esf1_water_depth_pipe31h_path.py version {0}\n".format(
                SCRIPT_VERSION
            )
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
        "extract_esf1_water_depth_pipe31h_path.py version {0}".format(
            SCRIPT_VERSION
        )
    )
    args = parse_arguments()
    input_dir = os.path.abspath(args.input_dir)
    odb_paths = find_odb_files(input_dir, args.odb)

    if args.list_steps or args.list_element_sets or args.list_endpoint_sets:
        for odb_path in odb_paths:
            if args.list_steps:
                print_steps(odb_path)
            if args.list_element_sets:
                print_element_sets(odb_path, args.instance)
            if args.list_endpoint_sets:
                print_endpoint_sets(odb_path, args.instance)
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
                requested_sets=args.element_set,
                requested_elements=args.element,
                requested_ranges=args.element_range,
                requested_steps=args.steps,
                step_range=args.step_range,
                frame_index=args.frame_index,
                start_node=args.start_node,
                end_node=args.end_node,
                start_node_set=args.start_node_set,
                end_node_set=args.end_node_set,
            )
        except Exception as exc:
            traceback_text = traceback.format_exc()
            failures.append((odb_path, str(exc), traceback_text))
            print("FAILED: {0}".format(odb_path))
            print("        {0}".format(exc))
            print(traceback_text)

    log_path = os.path.join(
        output_dir, "extract_esf1_water_depth_pipe31h_path.log"
    )
    write_execution_log(log_path, input_dir, output_dir, jobs, failures)
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
