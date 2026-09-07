from __future__ import print_function

"""Calculate maximum pipe S11 changes for user-defined ODB step pairs.

The script is self-contained and runs with ``abaqus python``. S11 is read at
every section point and every frame of each requested step. Each step is
enveloped independently; frames are not paired. Intermediate values are
processed in memory unless ``--write-intermediate`` is supplied. The final
report contains maximum signed delta S11 profiles along the pipeline path,
and the final data is split into inner-, middle-, and outer-fiber workbooks.
"""

import argparse
import datetime
import heapq
import math
import os
import re
import sys
import traceback
import zipfile

from abaqusConstants import ELEMENT_NODAL, ON
from odbAccess import openOdb


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPT_VERSION = "2026-09-06-r7"
DEFAULT_INSTANCE = "PART-1-1"
DEFAULT_PRECISION = 8

try:
    TEXT_TYPE = unicode
except NameError:
    TEXT_TYPE = str


def parse_arguments():
    parser = argparse.ArgumentParser(
        description=(
            "Calculate maximum signed delta S11 along a pipe path for "
            "one or more user-defined ODB step pairs."
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
        "--step-pair",
        action="append",
        nargs=2,
        default=[],
        metavar=("FIRST", "SECOND"),
        help=(
            "Step pair for delta S11. Each reference may be an exact step "
            "name or 1-based position. May be repeated."
        ),
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
        "--write-intermediate",
        action="store_true",
        help=(
            "Write per-frame, per-element, per-node, per-section-point S11 "
            "and delta S11 .rpt files. Default: calculate only in memory."
        ),
    )
    parser.add_argument(
        "--precision",
        type=int,
        default=DEFAULT_PRECISION,
        help=(
            "Digits after the decimal in scientific-format result columns "
            "(default: 8)."
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
    args = parser.parse_args()
    for first_label, last_label in args.pipe_element_range:
        if first_label > last_label:
            parser.error(
                "--pipe-element-range FIRST must be less than or equal to LAST"
            )
    if args.precision < 1 or args.precision > 12:
        parser.error("--precision must be between 1 and 12")
    listing = (
        args.list_steps
        or args.list_pipe_element_sets
        or args.list_endpoint_sets
    )
    if not listing and not args.step_pair:
        parser.error("at least one --step-pair FIRST SECOND is required")
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


def integer_value(text):
    try:
        return int(text)
    except ValueError:
        return None


def find_step_name(available_steps, requested_name):
    requested_upper = str(requested_name).upper()
    for step_name in available_steps:
        if str(step_name).upper() == requested_upper:
            return step_name
    return None


def resolve_step_reference(available_steps, reference):
    position = integer_value(str(reference))
    if position is not None:
        if position < 1 or position > len(available_steps):
            raise ValueError(
                "Step position {0} is outside the available range 1-{1}.".format(
                    position, len(available_steps)
                )
            )
        return available_steps[position - 1]
    name = find_step_name(available_steps, reference)
    if name is None:
        raise ValueError(
            "Step '{0}' was not found. Available steps: {1}".format(
                reference, ", ".join(available_steps)
            )
        )
    return name


def resolve_step_pairs(available_steps, requested_pairs):
    pairs = []
    seen = set()
    for first_reference, second_reference in requested_pairs:
        first_name = resolve_step_reference(available_steps, first_reference)
        second_name = resolve_step_reference(available_steps, second_reference)
        if first_name == second_name:
            raise ValueError(
                "A step pair must contain two different steps; both resolved "
                "to '{0}'.".format(first_name)
            )
        key = (first_name, second_name)
        if key not in seen:
            seen.add(key)
            pairs.append(key)
    return pairs


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


def resolve_s11_field(frame):
    exact_key = repository_key(frame.fieldOutputs, "S11")
    if exact_key is not None:
        field_output = frame.fieldOutputs[exact_key]
        index = component_index(field_output, "S11")
        return field_output, 0 if index is None else index, str(exact_key)
    stress_key = repository_key(frame.fieldOutputs, "S")
    if stress_key is None:
        return None, None, None
    field_output = frame.fieldOutputs[stress_key]
    index = component_index(field_output, "S11")
    if index is None:
        try:
            labels = list(field_output.componentLabels)
        except Exception:
            labels = []
        if labels:
            return None, None, None
        index = 0
    return field_output, index, "{0}.S11".format(stress_key)


def section_point_identity(value):
    section_point = getattr(value, "sectionPoint", None)
    if section_point is None:
        return ("NONE", ""), "NO_SECTION_POINT"
    number = getattr(section_point, "number", None)
    description = str(getattr(section_point, "description", "") or "")
    if number is not None:
        try:
            number_value = int(number)
        except (TypeError, ValueError):
            number_value = str(number)
        label = "SP {0}".format(number_value)
        if description:
            label += " " + description
        return ("NUMBER", number_value), label
    if description:
        return ("DESCRIPTION", description), description
    return ("UNKNOWN", str(section_point)), str(section_point)


def extract_s11_locations(
    frame, instance, output_nodes, output_element_labels
):
    field_output, component, source = resolve_s11_field(frame)
    if field_output is None:
        return {}, {}, None
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
    contributions = {}
    section_labels = {}
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
        section_identity, section_label = section_point_identity(value)
        location = (node_label, element_label, section_identity)
        contributions.setdefault(location, []).append(scalar)
        section_labels[location] = section_label
    averaged = {}
    for location, values in contributions.items():
        averaged[location] = sum(values) / float(len(values))
    return averaged, section_labels, source


def safe_filename(value):
    text = re.sub(r"[^A-Za-z0-9._-]+", "_", str(value)).strip("._-")
    return text or "STEP"


def scientific_text(value, precision):
    if value is None:
        return ""
    format_string = "{0:.%dE}" % precision
    return format_string.format(float(value))


def xml_escape(value):
    """Return text that is safe inside an Open XML element."""
    return (
        TEXT_TYPE(value)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def excel_column_name(index):
    """Convert a one-based column number to an Excel column name."""
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def xlsx_inline_cell(reference, value, style_index):
    return (
        '<c r="{0}" s="{1}" t="inlineStr"><is><t xml:space="preserve">'
        "{2}</t></is></c>"
    ).format(reference, style_index, xml_escape(value))


def xlsx_number_cell(reference, value, style_index):
    if value is None:
        return '<c r="{0}" s="{1}"/>'.format(reference, style_index)
    number = "{0:.15g}".format(float(value))
    return '<c r="{0}" s="{1}"><v>{2}</v></c>'.format(
        reference, style_index, number
    )


def write_zip_xml(archive, member_name, xml_text):
    archive.writestr(member_name, xml_text.encode("utf-8"))


def xlsx_styles_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="0.00000000E+00"/></numFmts>
  <fonts count="4">
    <font><sz val="10"/><name val="Arial"/><family val="2"/></font>
    <font><b/><sz val="14"/><color rgb="FF1F4E78"/><name val="Arial"/><family val="2"/></font>
    <font><i/><sz val="10"/><color rgb="FF666666"/><name val="Arial"/><family val="2"/></font>
    <font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/><family val="2"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FFD9E2F3"/></left><right style="thin"><color rgb="FFD9E2F3"/></right><top style="thin"><color rgb="FFD9E2F3"/></top><bottom style="thin"><color rgb="FFD9E2F3"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="3" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""


def xlsx_worksheet_xml(
    sheet_name, title, note, section_note, headers, rows
):
    column_count = len(headers)
    last_column = excel_column_name(column_count)
    last_row = 5 + len(rows)
    row_xml = []
    row_xml.append(
        '<row r="1" ht="22" customHeight="1">{0}</row>'.format(
            xlsx_inline_cell("A1", title, 1)
        )
    )
    row_xml.append(
        '<row r="2">{0}</row>'.format(
            xlsx_inline_cell("A2", note, 2)
        )
    )
    row_xml.append(
        '<row r="3">{0}</row>'.format(
            xlsx_inline_cell("A3", section_note, 2)
        )
    )
    header_cells = []
    for column_index, header in enumerate(headers, 1):
        header_cells.append(
            xlsx_inline_cell(
                "{0}5".format(excel_column_name(column_index)), header, 3
            )
        )
    row_xml.append(
        '<row r="5" ht="45" customHeight="1">{0}</row>'.format(
            "".join(header_cells)
        )
    )
    for row_index, values in enumerate(rows, 6):
        cells = []
        for column_index, value in enumerate(values, 1):
            cells.append(
                xlsx_number_cell(
                    "{0}{1}".format(
                        excel_column_name(column_index), row_index
                    ),
                    value,
                    4,
                )
            )
        row_xml.append('<row r="{0}">{1}</row>'.format(row_index, "".join(cells)))
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:{last_column}{last_row}"/>
  <sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane xSplit="1" ySplit="5" topLeftCell="B6" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="B6" sqref="B6"/></sheetView></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols><col min="1" max="1" width="18" customWidth="1"/><col min="2" max="{column_count}" width="30" customWidth="1"/></cols>
  <sheetData>{rows_xml}</sheetData>
  <autoFilter ref="A5:{last_column}{last_row}"/>
</worksheet>""".format(
        last_column=last_column,
        last_row=last_row,
        column_count=column_count,
        rows_xml="".join(row_xml),
    )


def write_data_xlsx(path, sheet_name, title, note, section_note, headers, rows):
    """Write one editable, dependency-free Excel workbook."""
    created = datetime.datetime.utcnow().replace(microsecond=0).isoformat() + "Z"
    worksheet_xml = xlsx_worksheet_xml(
        sheet_name, title, note, section_note, headers, rows
    )
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>"""
    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"""
    workbook_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews><workbookView/></bookViews>
  <sheets><sheet name="{0}" sheetId="1" r:id="rId1"/></sheets>
  <calcPr calcId="191029" fullCalcOnLoad="1" forceFullCalc="1"/>
</workbook>""".format(xml_escape(sheet_name))
    workbook_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""
    core_props = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>{0}</dc:title><dc:creator>Abaqus postprocessing</dc:creator><cp:lastModifiedBy>Abaqus postprocessing</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">{1}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">{1}</dcterms:modified>
</cp:coreProperties>""".format(xml_escape(title), created)
    app_props = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Abaqus Python</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>{0}</vt:lpstr></vt:vector></TitlesOfParts><Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion>
</Properties>""".format(xml_escape(sheet_name))
    archive = zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED)
    try:
        write_zip_xml(archive, "[Content_Types].xml", content_types)
        write_zip_xml(archive, "_rels/.rels", root_rels)
        write_zip_xml(archive, "docProps/core.xml", core_props)
        write_zip_xml(archive, "docProps/app.xml", app_props)
        write_zip_xml(archive, "xl/workbook.xml", workbook_xml)
        write_zip_xml(archive, "xl/_rels/workbook.xml.rels", workbook_rels)
        write_zip_xml(archive, "xl/styles.xml", xlsx_styles_xml())
        write_zip_xml(archive, "xl/worksheets/sheet1.xml", worksheet_xml)
    finally:
        archive.close()


def frame_time(frame):
    if frame is None:
        return None
    try:
        return float(frame.frameValue)
    except Exception:
        return None


def intermediate_path(
    output_dir, odb_path, pair_index, first_name, second_name
):
    stem = os.path.splitext(os.path.basename(odb_path))[0]
    filename = "{0}_Pair{1:03d}_{2}_TO_{3}_S11_INTERMEDIATE.rpt".format(
        stem,
        pair_index,
        safe_filename(first_name),
        safe_filename(second_name),
    )
    return os.path.abspath(os.path.join(output_dir, filename))


def final_report_path(output_dir, odb_path):
    stem = os.path.splitext(os.path.basename(odb_path))[0]
    return os.path.abspath(
        os.path.join(output_dir, stem + "_MAX_DELTA_S11_PATH.rpt")
    )


def fiber_excel_path(output_dir, odb_path, fiber_name):
    stem = os.path.splitext(os.path.basename(odb_path))[0]
    return os.path.abspath(
        os.path.join(
            output_dir,
            stem + "_MAX_DELTA_S11_{0}_FIBER.xlsx".format(fiber_name),
        )
    )


def open_intermediate_report(
    path, odb_path, instance_name, first_name, second_name
):
    report = open(path, "w")
    report.write("Pipe S11 Step-Pair Intermediate Results\n")
    report.write("Script version: {0}\n".format(SCRIPT_VERSION))
    report.write("ODB: {0}\n".format(odb_path.replace("\\", "/")))
    report.write("Instance: {0}\n".format(instance_name))
    report.write("First step: {0}\n".format(first_name))
    report.write("Second step: {0}\n".format(second_name))
    report.write(
        "Frames are not paired. Each step is enveloped independently at "
        "each matching element/node/section-point location.\n"
    )
    report.write(
        "Delta S11 = maximum S11 over all first-step frames - maximum "
        "S11 over all second-step frames.\n\n"
    )
    return report


def sorted_locations(locations, route_distances):
    return sorted(
        locations,
        key=lambda location: (
            route_distances.get(location[0], float("inf")),
            location[0],
            location[1],
            str(location[2]),
        ),
    )


def scan_step_s11_envelope(
    step_name,
    role,
    frames,
    instance,
    output_nodes,
    output_element_labels,
    route_distances,
    intermediate,
    precision,
):
    maximum_by_location = {}
    control_by_location = {}
    section_labels = {}
    source_names = set()
    sample_count = 0
    if intermediate is not None:
        intermediate.write("RAW S11 - {0} STEP: {1}\n".format(role, step_name))
        intermediate.write(
            "Step Role\tStep Name\tFrame Index\tFrame Time\tPath Distance\t"
            "Node Label\tElement Label\tSection Point\tS11\n"
        )
    for frame_index, frame in enumerate(frames):
        values, frame_sections, source = extract_s11_locations(
            frame,
            instance,
            output_nodes,
            output_element_labels,
        )
        if source:
            source_names.add(source)
        for location in sorted_locations(values.keys(), route_distances):
            value = values[location]
            sample_count += 1
            section_label = frame_sections.get(location, "")
            section_labels[location] = section_label
            previous = maximum_by_location.get(location)
            if previous is None or value > previous:
                maximum_by_location[location] = value
                control_by_location[location] = {
                    "frame": frame_index,
                    "time": frame_time(frame),
                    "value": value,
                    "section": section_label,
                }
            if intermediate is not None:
                row = (
                    role,
                    str(step_name),
                    str(frame_index),
                    scientific_text(frame_time(frame), precision),
                    scientific_text(route_distances.get(location[0]), precision),
                    str(location[0]),
                    str(location[1]),
                    section_label,
                    scientific_text(value, precision),
                )
                intermediate.write("\t".join(row) + "\n")
    if intermediate is not None:
        intermediate.write("\n")
    return {
        "maximum": maximum_by_location,
        "control": control_by_location,
        "sections": section_labels,
        "sources": source_names,
        "samples": sample_count,
    }


def calculate_pair_profile(
    odb,
    first_name,
    second_name,
    pair_index,
    odb_path,
    output_dir,
    instance,
    output_nodes,
    output_element_labels,
    route_distances,
    write_intermediate,
    precision,
):
    first_frames = odb.steps[first_name].frames
    second_frames = odb.steps[second_name].frames
    if not first_frames or not second_frames:
        raise ValueError(
            "Step pair '{0}' -> '{1}' cannot be compared because one or "
            "both steps contain no frames.".format(first_name, second_name)
        )
    maximum_by_node = {}
    control_by_node = {}
    delta_by_location = {}
    matched_section_labels = {}
    intermediate = None
    intermediate_output = None
    try:
        if write_intermediate:
            intermediate_output = intermediate_path(
                output_dir,
                odb_path,
                pair_index,
                first_name,
                second_name,
            )
            intermediate = open_intermediate_report(
                intermediate_output,
                odb_path,
                instance.name,
                first_name,
                second_name,
            )
        first_envelope = scan_step_s11_envelope(
            first_name,
            "FIRST",
            first_frames,
            instance,
            output_nodes,
            output_element_labels,
            route_distances,
            intermediate,
            precision,
        )
        second_envelope = scan_step_s11_envelope(
            second_name,
            "SECOND",
            second_frames,
            instance,
            output_nodes,
            output_element_labels,
            route_distances,
            intermediate,
            precision,
        )
        first_values = first_envelope["maximum"]
        second_values = second_envelope["maximum"]
        first_locations = set(first_values.keys())
        second_locations = set(second_values.keys())
        common_locations = first_locations.intersection(second_locations)
        for location in common_locations:
            first_value = first_values[location]
            second_value = second_values[location]
            delta = first_value - second_value
            delta_by_location[location] = delta
            node_label = location[0]
            matched_section_labels[location[2]] = second_envelope[
                "sections"
            ].get(
                location,
                first_envelope["sections"].get(location, ""),
            )
            previous = maximum_by_node.get(node_label)
            if previous is None or delta > previous:
                maximum_by_node[node_label] = delta
                first_control = first_envelope["control"][location]
                second_control = second_envelope["control"][location]
                control_by_node[node_label] = {
                    "first_frame": first_control["frame"],
                    "first_time": first_control["time"],
                    "second_frame": second_control["frame"],
                    "second_time": second_control["time"],
                    "element": location[1],
                    "section": second_envelope["sections"].get(
                        location,
                        first_envelope["sections"].get(location, ""),
                    ),
                    "first": first_value,
                    "second": second_value,
                    "delta": delta,
                }
        if intermediate is not None:
            intermediate.write("MATCHED LOCATION ENVELOPES AND DELTA S11\n")
            intermediate.write(
                "Path Distance\tNode Label\tElement Label\tSection Point\t"
                "Maximum S11 First\tFirst Controlling Frame\t"
                "First Controlling Frame Time\tMaximum S11 Second\t"
                "Second Controlling Frame\tSecond Controlling Frame Time\t"
                "Delta S11\n"
            )
            union_locations = first_locations.union(second_locations)
            for location in sorted_locations(union_locations, route_distances):
                first_value = first_values.get(location)
                second_value = second_values.get(location)
                first_control = first_envelope["control"].get(location, {})
                second_control = second_envelope["control"].get(location, {})
                delta = (
                    first_value - second_value
                    if first_value is not None and second_value is not None
                    else None
                )
                section_label = second_envelope["sections"].get(
                    location,
                    first_envelope["sections"].get(location, ""),
                )
                row = (
                    scientific_text(route_distances.get(location[0]), precision),
                    str(location[0]),
                    str(location[1]),
                    section_label,
                    scientific_text(first_value, precision),
                    str(first_control.get("frame", "")),
                    scientific_text(first_control.get("time"), precision),
                    scientific_text(second_value, precision),
                    str(second_control.get("frame", "")),
                    scientific_text(second_control.get("time"), precision),
                    scientific_text(delta, precision),
                )
                intermediate.write("\t".join(row) + "\n")
    finally:
        if intermediate is not None:
            intermediate.close()
    source_names = first_envelope["sources"].union(second_envelope["sources"])
    return {
        "first": first_name,
        "second": second_name,
        "first_frames": len(first_frames),
        "second_frames": len(second_frames),
        "maximum": maximum_by_node,
        "control": control_by_node,
        "deltas": delta_by_location,
        "section_labels": matched_section_labels,
        "sources": sorted(source_names),
        "first_samples": first_envelope["samples"],
        "second_samples": second_envelope["samples"],
        "compared_locations": len(common_locations),
        "missing_first": len(second_locations - first_locations),
        "missing_second": len(first_locations - second_locations),
        "intermediate": intermediate_output,
    }


FIBER_NAMES = ("INNER", "MIDDLE", "OUTER")
TARGET_ANGLES = (-90, 0, 90, 180)


def section_identity_sort_key(identity):
    kind = str(identity[0]).upper() if identity else ""
    value = identity[1] if len(identity) > 1 else ""
    if kind == "NUMBER":
        try:
            return (0, 0, float(value), "")
        except (TypeError, ValueError):
            return (0, 1, 0.0, str(value))
    return (1, 0, 0.0, str(identity))


FLOAT_PATTERN = r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[EeDd][+-]?\d+)?"


def section_radius_from_label(label):
    text = str(label).strip()
    text_without_sp = re.sub(
        r"^\s*SP\s+[+-]?\d+(?:\.\d+)?\s*[:,-]?\s*",
        "",
        text,
        flags=re.IGNORECASE,
    )
    patterns = (
        r"THICK\s*PIPE\s*SECTION\s*RADIUS\s*(?:=|:)?\s*(%s)"
        % FLOAT_PATTERN,
        r"SECTION\s*RADIUS\s*(?:=|:)?\s*(%s)" % FLOAT_PATTERN,
        r"RADIAL(?:\s+(?:COORDINATE|POSITION|RADIUS))?\s*(?:=|:)?\s*(%s)"
        % FLOAT_PATTERN,
        r"RADIUS\s*(?:=|:)?\s*(%s)" % FLOAT_PATTERN,
    )
    for pattern in patterns:
        match = re.search(pattern, text_without_sp, flags=re.IGNORECASE)
        if match:
            return float(match.group(1).replace("D", "E").replace("d", "e"))
    number_tokens = re.findall(FLOAT_PATTERN, text_without_sp)
    if len(number_tokens) == 1:
        return float(number_tokens[0].replace("D", "E").replace("d", "e"))
    return None


def section_angle_from_label(label):
    text = str(label).strip()
    text_without_sp = re.sub(
        r"^\s*SP\s+[+-]?\d+(?:\.\d+)?\s*[:,-]?\s*",
        "",
        text,
        flags=re.IGNORECASE,
    )
    patterns = (
        r"THICK\s*PIPE\s*SECTION\s*ANGLE\s*(?:=|:)?\s*(%s)"
        % FLOAT_PATTERN,
        r"SECTION\s*ANGLE\s*(?:=|:)?\s*(%s)" % FLOAT_PATTERN,
        r"CIRCUMFERENTIAL(?:\s+(?:ANGLE|POSITION))?\s*(?:=|:)?\s*(%s)"
        % FLOAT_PATTERN,
        r"ANGLE\s*(?:=|:)?\s*(%s)" % FLOAT_PATTERN,
    )
    for pattern in patterns:
        match = re.search(pattern, text_without_sp, flags=re.IGNORECASE)
        if match:
            return float(match.group(1).replace("D", "E").replace("d", "e"))
    return None


def radius_values_match(first, second):
    tolerance = 1.0e-7 * max(1.0, abs(first), abs(second))
    return abs(first - second) <= tolerance


def target_angle(value):
    normalized = float(value)
    while normalized <= -180.0:
        normalized += 360.0
    while normalized > 180.0:
        normalized -= 360.0
    for requested in TARGET_ANGLES:
        if abs(normalized - float(requested)) <= 1.0e-5:
            return requested
    return None


def resolve_fiber_angle_section_groups(profiles):
    labels_by_identity = {}
    for profile in profiles:
        for location in profile["deltas"].keys():
            identity = location[2]
            label = profile["section_labels"].get(identity, str(identity))
            if identity not in labels_by_identity or not labels_by_identity[identity]:
                labels_by_identity[identity] = label
    if not labels_by_identity:
        raise ValueError(
            "No matching section-point S11 locations were found for the "
            "requested step pairs."
        )
    radius_by_identity = {}
    angle_by_identity = {}
    missing_coordinates = []
    for identity, label in labels_by_identity.items():
        radius = section_radius_from_label(label)
        angle = section_angle_from_label(label)
        if radius is None or angle is None:
            missing_coordinates.append(label)
        else:
            radius_by_identity[identity] = radius
            angle_by_identity[identity] = angle
    if missing_coordinates:
        available = ", ".join(
            labels_by_identity[identity]
            for identity in sorted(
                labels_by_identity.keys(), key=section_identity_sort_key
            )
        )
        raise ValueError(
            "The output thick-pipe section radius and angle could not both "
            "be read from every matching section-point description. Matching "
            "section points were: {0}".format(available)
        )
    selected_identities = []
    omitted_angles = set()
    normalized_angle_by_identity = {}
    for identity, angle in angle_by_identity.items():
        normalized = target_angle(angle)
        if normalized is None:
            omitted_angles.add(angle)
        else:
            normalized_angle_by_identity[identity] = normalized
            selected_identities.append(identity)
    if not selected_identities:
        raise ValueError(
            "No matching section points were found at angles -90, 0, 90, "
            "or 180 degrees."
        )
    radial_groups = []
    for identity in sorted(
        selected_identities,
        key=lambda item: (
            abs(radius_by_identity[item]),
            section_identity_sort_key(item),
        ),
    ):
        absolute_radius = abs(radius_by_identity[identity])
        if (
            radial_groups
            and radius_values_match(radial_groups[-1][0], absolute_radius)
        ):
            radial_groups[-1][1].add(identity)
        else:
            radial_groups.append((absolute_radius, set([identity])))
    if len(radial_groups) != 3:
        radius_text = ", ".join(
            "{0:.12g}".format(group[0]) for group in radial_groups
        )
        raise ValueError(
            "Expected exactly three distinct absolute thick-pipe section "
            "radius magnitudes, but found {0}: {1}".format(
                len(radial_groups), radius_text
            )
        )
    if not radius_values_match(radial_groups[-1][0], 1.0):
        raise ValueError(
            "The largest absolute thick-pipe section radius must be 1.0 for "
            "the outer fiber, but the largest value found was {0:.12g}.".format(
                radial_groups[-1][0]
            )
        )
    fiber_radius = {
        "INNER": radial_groups[0][0],
        "MIDDLE": radial_groups[1][0],
        "OUTER": radial_groups[2][0],
    }
    groups = {}
    for fiber_name in FIBER_NAMES:
        for angle in TARGET_ANGLES:
            groups[(fiber_name, angle)] = set()
    for identity in selected_identities:
        absolute_radius = abs(radius_by_identity[identity])
        fiber_name = None
        for candidate in FIBER_NAMES:
            if radius_values_match(absolute_radius, fiber_radius[candidate]):
                fiber_name = candidate
                break
        groups[(fiber_name, normalized_angle_by_identity[identity])].add(identity)
    missing_groups = []
    duplicate_groups = []
    for fiber_name in FIBER_NAMES:
        for angle in TARGET_ANGLES:
            identities = groups[(fiber_name, angle)]
            if not identities:
                missing_groups.append("{0} at {1} deg".format(fiber_name, angle))
            elif len(identities) > 1:
                duplicate_groups.append(
                    "{0} at {1} deg: {2}".format(
                        fiber_name,
                        angle,
                        ", ".join(
                            labels_by_identity[item]
                            for item in sorted(
                                identities, key=section_identity_sort_key
                            )
                        ),
                    )
                )
    if missing_groups:
        raise ValueError(
            "Required thick-pipe radius/angle section points are missing: "
            "{0}".format(", ".join(missing_groups))
        )
    if duplicate_groups:
        raise ValueError(
            "Positive and negative radius points would be combined in the "
            "same fiber/angle column, which is not allowed. Duplicate groups: "
            "{0}".format("; ".join(duplicate_groups))
        )
    method = (
        "output thick-pipe section radius and angle: smallest |radius|=inner, "
        "middle |radius|=middle, |radius|=1.0=outer; radius signs remain "
        "separate at angles -90, 0, 90, and 180 degrees"
    )
    if omitted_angles:
        method += "; other angles omitted: {0}".format(
            ", ".join(
                "{0:.12g}".format(value) for value in sorted(omitted_angles)
            )
        )
    display = {}
    for fiber_name in FIBER_NAMES:
        for angle in TARGET_ANGLES:
            identity = next(iter(groups[(fiber_name, angle)]))
            display[(fiber_name, angle)] = (
                "radius={0:.12g}, angle={1} deg: {2}".format(
                    radius_by_identity[identity],
                    angle,
                    labels_by_identity[identity],
                )
            )
    return groups, display, method


def maximum_delta_by_node_for_section_group(profile, identities):
    maximum = {}
    for location, delta in profile["deltas"].items():
        if location[2] not in identities:
            continue
        node_label = location[0]
        previous = maximum.get(node_label)
        if previous is None or delta > previous:
            maximum[node_label] = delta
    return maximum


def pair_fiber_angle_column_name(profile, fiber_name, angle):
    return "DELTA_S11 MAX {0} FIBER ANGLE {1} DEG [{2} - {3}]".format(
        fiber_name, angle, profile["first"], profile["second"]
    )


def write_fiber_excel_workbooks(
    output_dir,
    odb_path,
    route_distances,
    output_nodes,
    profiles,
    profile_fiber_values,
    group_display,
):
    """Write one final-data workbook for each thick-pipe fiber radius."""
    paths = {}
    odb_name = os.path.basename(odb_path)
    for fiber_name in FIBER_NAMES:
        headers = ["Pipeline Distance"]
        for profile in profiles:
            for angle in TARGET_ANGLES:
                headers.append(
                    pair_fiber_angle_column_name(
                        profile, fiber_name, angle
                    )
                )
        rows = []
        for node_label in output_nodes:
            row = [route_distances[node_label]]
            for fiber_values in profile_fiber_values:
                for angle in TARGET_ANGLES:
                    row.append(
                        fiber_values[(fiber_name, angle)].get(node_label)
                    )
            rows.append(row)
        title = "{0} Fiber - Maximum Delta S11 Along Pipeline Path".format(
            fiber_name.title()
        )
        note = (
            "ODB: {0}. Delta S11 = maximum S11 over all first-step "
            "frames - maximum S11 over all second-step frames; frames "
            "are not paired."
        ).format(odb_name)
        section_note = "Section points: {0}".format(
            "; ".join(
                "{0} deg ({1})".format(
                    angle, group_display[(fiber_name, angle)]
                )
                for angle in TARGET_ANGLES
            )
        )
        path = fiber_excel_path(output_dir, odb_path, fiber_name)
        write_data_xlsx(
            path,
            fiber_name.title() + " Fiber",
            title,
            note,
            section_note,
            headers,
            rows,
        )
        paths[fiber_name] = path
    return paths


def write_final_report(
    path,
    odb_path,
    instance_name,
    start_label,
    end_label,
    route_distances,
    output_nodes,
    output_element_labels,
    profiles,
    precision,
):
    groups, group_display, mapping_method = (
        resolve_fiber_angle_section_groups(profiles)
    )
    profile_fiber_values = []
    for profile in profiles:
        profile_fiber_values.append(
            dict(
                (
                    (fiber_name, angle),
                    maximum_delta_by_node_for_section_group(
                        profile, groups[(fiber_name, angle)]
                    ),
                )
                for fiber_name in FIBER_NAMES
                for angle in TARGET_ANGLES
            )
        )
    with open(path, "w") as report:
        report.write("Pipe Maximum Delta S11 Along Pipeline Path\n")
        report.write("Script version: {0}\n".format(SCRIPT_VERSION))
        report.write("ODB: {0}\n".format(odb_path.replace("\\", "/")))
        report.write("Instance: {0}\n".format(instance_name))
        report.write(
            "Path: node {0} to node {1}; length={2}\n".format(
                start_label,
                end_label,
                scientific_text(route_distances[end_label], precision),
            )
        )
        report.write(
            "Pipe elements: {0}; path nodes: {1}\n".format(
                len(output_element_labels), len(output_nodes)
            )
        )
        report.write(
            "At each matching element/node/section-point location: Delta "
            "S11 = maximum S11 over all first-step frames - maximum S11 "
            "over all second-step frames.\n"
        )
        report.write(
            "Frames are not paired. The two steps are enveloped independently. "
            "For each radius/angle column, the final value is the maximum "
            "signed delta S11 across contributing pipe elements at each path "
            "node. Positive and negative radius section points are not "
            "combined.\n"
        )
        report.write("Section-point mapping method: {0}\n".format(mapping_method))
        for fiber_name in FIBER_NAMES:
            for angle in TARGET_ANGLES:
                report.write(
                    "{0} fiber at {1} deg: {2}\n".format(
                        fiber_name.title(),
                        angle,
                        group_display[(fiber_name, angle)],
                    )
                )
        for index, profile in enumerate(profiles, 1):
            report.write(
                "Pair {0}: '{1}' -> '{2}'; frames={3}/{4}; samples={5}/{6}; "
                "matched locations={7}\n".format(
                    index,
                    profile["first"],
                    profile["second"],
                    profile["first_frames"],
                    profile["second_frames"],
                    profile["first_samples"],
                    profile["second_samples"],
                    profile["compared_locations"],
                )
            )
        report.write("\n")
        headers = ["Pipeline Distance"]
        for profile in profiles:
            for fiber_name in FIBER_NAMES:
                for angle in TARGET_ANGLES:
                    headers.append(
                        pair_fiber_angle_column_name(
                            profile, fiber_name, angle
                        )
                    )
        report.write("\t".join(headers) + "\n")
        for node_label in output_nodes:
            row = [scientific_text(route_distances[node_label], precision)]
            for fiber_values in profile_fiber_values:
                for fiber_name in FIBER_NAMES:
                    for angle in TARGET_ANGLES:
                        row.append(
                            scientific_text(
                                fiber_values[(fiber_name, angle)].get(
                                    node_label
                                ),
                                precision,
                            )
                        )
            report.write("\t".join(row) + "\n")
    excel_paths = write_fiber_excel_workbooks(
        os.path.dirname(path),
        odb_path,
        route_distances,
        output_nodes,
        profiles,
        profile_fiber_values,
        group_display,
    )
    return group_display, mapping_method, excel_paths


def process_odb(odb_path, output_dir, args):
    odb = None
    messages = []
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
        step_pairs = resolve_step_pairs(available_steps, args.step_pair)
        profiles = []
        for pair_index, pair in enumerate(step_pairs, 1):
            profile = calculate_pair_profile(
                odb,
                pair[0],
                pair[1],
                pair_index,
                odb_path,
                output_dir,
                instance,
                output_nodes,
                output_element_labels,
                route_distances,
                args.write_intermediate,
                args.precision,
            )
            profiles.append(profile)
            if args.verbose and profile["intermediate"]:
                print("Intermediate: {0}".format(profile["intermediate"]))
        report_path = final_report_path(output_dir, odb_path)
        group_display, mapping_method, excel_paths = write_final_report(
            report_path,
            odb_path,
            instance_key,
            start_label,
            end_label,
            route_distances,
            output_nodes,
            output_element_labels,
            profiles,
            args.precision,
        )
        messages.append("ODB: {0}".format(odb_path))
        messages.append("Instance: {0}".format(instance_key))
        messages.append(
            "Path: node {0} ({1}) to node {2} ({3}); length={4:.12g}".format(
                start_label,
                start_source,
                end_label,
                end_source,
                route_distances[end_label],
            )
        )
        messages.append(
            "Selection: {0} pipe element(s), {1} path node(s)".format(
                len(output_element_labels), len(output_nodes)
            )
        )
        messages.append(
            "Section-point mapping method: {0}".format(mapping_method)
        )
        for fiber_name in FIBER_NAMES:
            for angle in TARGET_ANGLES:
                messages.append(
                    "{0} fiber at {1} deg: {2}".format(
                        fiber_name.title(),
                        angle,
                        group_display[(fiber_name, angle)],
                    )
                )
        if unsupported:
            messages.append(
                "Unsupported pipe connectivity omitted: {0} element(s)".format(
                    len(unsupported)
                )
            )
        messages.extend("Selection note: " + note for note in selection_notes)
        for index, profile in enumerate(profiles, 1):
            messages.append(
                "Pair {0}: '{1}' -> '{2}'; frame counts={3}/{4}; "
                "samples={5}/{6}; matched locations={7}; S11 sources={8}".format(
                    index,
                    profile["first"],
                    profile["second"],
                    profile["first_frames"],
                    profile["second_frames"],
                    profile["first_samples"],
                    profile["second_samples"],
                    profile["compared_locations"],
                    ", ".join(profile["sources"])
                    if profile["sources"]
                    else "not found",
                )
            )
            messages.append(
                "  unmatched locations: missing first={0}, missing "
                "second={1}".format(
                    profile["missing_first"], profile["missing_second"]
                )
            )
            if profile["intermediate"]:
                messages.append(
                    "  Intermediate: {0}".format(profile["intermediate"])
                )
        messages.append("Final report: {0}".format(report_path))
        for fiber_name in FIBER_NAMES:
            messages.append(
                "{0} fiber Excel: {1}".format(
                    fiber_name.title(), excel_paths[fiber_name]
                )
            )
        if args.verbose:
            print("Final report: {0}".format(report_path))
            for fiber_name in FIBER_NAMES:
                print(
                    "{0} fiber Excel: {1}".format(
                        fiber_name.title(), excel_paths[fiber_name]
                    )
                )
        return report_path, profiles, messages
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
            "extract_pipe_s11_delta_step_pairs.py version {0}\n".format(
                SCRIPT_VERSION
            )
        )
        log_file.write("Time: {0}\n".format(datetime.datetime.now()))
        log_file.write("sys.argv: {0}\n".format(repr(sys.argv)))
        log_file.write("Input directory: {0}\n".format(input_dir))
        log_file.write("Output directory: {0}\n".format(output_dir))
        log_file.write("Frame handling: each step enveloped independently\n")
        log_file.write(
            "Delta at matching location: max S11 over first-step frames - "
            "max S11 over second-step frames\n"
        )
        log_file.write("Final envelope: maximum signed delta S11\n")
        log_file.write("Successful ODBs: {0}\n".format(len(successes)))
        log_file.write("Failed ODBs: {0}\n".format(len(failures)))
        for odb_path, report_path, messages in successes:
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
            "extract_pipe_s11_delta_step_pairs.py version {0}".format(
                SCRIPT_VERSION
            )
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
            report_path, unused_profiles, messages = process_odb(
                odb_path, output_dir, args
            )
            successes.append((odb_path, report_path, messages))
        except Exception as exc:
            failures.append((odb_path, str(exc), traceback.format_exc()))
            if args.verbose:
                print("FAILED: {0}".format(odb_path))
    log_path = os.path.join(
        output_dir, "extract_pipe_s11_delta_step_pairs.log"
    )
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
