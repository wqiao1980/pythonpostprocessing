from __future__ import print_function

"""Extract element histories for selected pipes, connectors, and springs.

The script is self-contained and runs with ``abaqus python`` without opening
Abaqus/CAE or Viewer. It writes a tabular report and a native Excel workbook
with editable history charts for ESF1, CTF1, S11, E11, and CTF1 + S11 pairs.
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
SCRIPT_VERSION = "2026-09-06-r1"
DEFAULT_INSTANCE = "PART-1-1"
VARIABLES = ("ESF1", "CTF1", "S11", "E11")
PARENT_FIELDS = {
    "ESF1": ("ESF", "ESF1"),
    "CTF1": ("CTF", "CTF1"),
    "S11": ("S", "S11"),
    "E11": ("E", "E11"),
}
CHART_COLORS = (
    "4472C4",
    "ED7D31",
    "70AD47",
    "A5A5A5",
    "FFC000",
    "5B9BD5",
    "264478",
    "9E480E",
    "43682B",
    "636363",
)


def add_group_arguments(parser, prefix, description):
    parser.add_argument(
        "--{0}-element-set".format(prefix),
        action="append",
        default=[],
        metavar="SET",
        help=(
            "Select a {0} element set. May be repeated; selections form a "
            "union.".format(description)
        ),
    )
    parser.add_argument(
        "--{0}-element".format(prefix),
        action="append",
        nargs="+",
        type=int,
        default=[],
        metavar="LABEL",
        help=(
            "Select one or more exact {0} element labels. May be repeated."
            .format(description)
        ),
    )
    parser.add_argument(
        "--{0}-element-range".format(prefix),
        action="append",
        nargs=2,
        type=int,
        default=[],
        metavar=("FIRST", "LAST"),
        help=(
            "Select an inclusive {0} element-label range. May be repeated."
            .format(description)
        ),
    )


def parse_arguments():
    parser = argparse.ArgumentParser(
        description=(
            "Write ESF1 pipe, CTF1 connector, and S11/E11 spring histories "
            "for selected element sets or labels."
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
        "--output-name",
        default=None,
        help="Custom .rpt filename for a single --odb run only.",
    )
    parser.add_argument(
        "--instance",
        default=DEFAULT_INSTANCE,
        help="Instance containing the selected elements (default: PART-1-1).",
    )
    add_group_arguments(parser, "pipe", "pipe")
    add_group_arguments(parser, "connector", "connector")
    add_group_arguments(parser, "spring", "spring")
    parser.add_argument(
        "--list-steps",
        action="store_true",
        help="List ordered steps and frame counts, then exit.",
    )
    parser.add_argument(
        "--list-element-sets",
        action="store_true",
        help=(
            "List element sets containing pipe, connector, or spring "
            "elements in --instance, then exit."
        ),
    )
    step_selection = parser.add_mutually_exclusive_group()
    step_selection.add_argument(
        "--steps",
        nargs="+",
        default=None,
        metavar="STEP",
        help="One or more exact step names to include (default: all steps).",
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
    if args.output_name and len(args.odb) != 1:
        parser.error("--output-name requires exactly one --odb")
    for prefix in ("pipe", "connector", "spring"):
        for first_label, last_label in getattr(
            args, prefix + "_element_range"
        ):
            if first_label > last_label:
                parser.error(
                    "--{0}-element-range FIRST must be less than or equal "
                    "to LAST".format(prefix)
                )
    if not (args.list_steps or args.list_element_sets):
        any_selection = False
        for prefix in ("pipe", "connector", "spring"):
            if (
                getattr(args, prefix + "_element_set")
                or getattr(args, prefix + "_element")
                or getattr(args, prefix + "_element_range")
            ):
                any_selection = True
        if not any_selection:
            parser.error(
                "select at least one pipe, connector, or spring element set, "
                "label, or range"
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


def default_report_name(odb_path):
    stem = os.path.splitext(os.path.basename(odb_path))[0]
    return stem + "_PIPE_CONNECTOR_SPRING_HISTORY.rpt"


def default_excel_name(odb_path):
    stem = os.path.splitext(os.path.basename(odb_path))[0]
    return stem + "_PIPE_CONNECTOR_SPRING_HISTORY.xlsx"


def normalized_output_name(name):
    if not name:
        return None
    root, extension = os.path.splitext(name)
    if extension.lower() == ".rpt":
        return name
    return name + ".rpt"


def build_jobs(odb_paths, output_dir, output_name):
    jobs = []
    seen = set()
    custom_report = normalized_output_name(output_name)
    for odb_path in odb_paths:
        report_name = custom_report or default_report_name(odb_path)
        report_path = os.path.abspath(os.path.join(output_dir, report_name))
        if custom_report:
            excel_name = os.path.splitext(custom_report)[0] + ".xlsx"
        else:
            excel_name = default_excel_name(odb_path)
        excel_path = os.path.abspath(os.path.join(output_dir, excel_name))
        for path in (report_path, excel_path):
            key = os.path.normcase(path)
            if key in seen:
                raise ValueError("More than one output maps to: {0}".format(path))
            seen.add(key)
        jobs.append((odb_path, report_path, excel_path))
    return jobs


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
        "Element set '{0}' was not found. Available sets: {1}".format(
            requested_name, ", ".join(available) if available else "(none)"
        )
    )


def element_kind(element):
    element_type = str(element.type).upper()
    if element_type.startswith("PIPE"):
        return "pipe"
    if element_type.startswith("CONN"):
        return "connector"
    if element_type.startswith("SPRING"):
        return "spring"
    return None


def flattened_requested_labels(groups):
    labels = []
    for group in groups:
        labels.extend(group)
    return labels


def resolve_group_selection(
    odb,
    instance,
    element_by_label,
    kind,
    requested_sets,
    requested_elements,
    requested_ranges,
):
    exact_labels = flattened_requested_labels(requested_elements)
    if not (requested_sets or exact_labels or requested_ranges):
        return [], [], []
    instance_labels = set(element_by_label.keys())
    compatible_labels = set(
        label
        for label, element in element_by_label.items()
        if element_kind(element) == kind
    )
    selected = set()
    descriptions = []
    notes = []
    for requested_name in requested_sets:
        labels, resolved_name = resolve_element_set(
            odb, instance, requested_name, instance_labels
        )
        matching = labels.intersection(compatible_labels)
        selected.update(matching)
        descriptions.append(resolved_name)
        notes.append(
            "{0}: {1} {2} element(s); {3} other element(s) omitted".format(
                resolved_name,
                len(matching),
                kind,
                len(labels) - len(matching),
            )
        )
    missing = sorted(set(exact_labels).difference(instance_labels))
    if missing:
        raise ValueError(
            "Requested {0} element labels are missing from instance "
            "'{1}': {2}".format(
                kind,
                instance.name,
                ", ".join(str(label) for label in missing),
            )
        )
    wrong_type = sorted(
        set(exact_labels).intersection(instance_labels).difference(
            compatible_labels
        )
    )
    if wrong_type:
        raise ValueError(
            "These --{0}-element labels do not have a {0} element type: "
            "{1}".format(kind, ", ".join(str(label) for label in wrong_type))
        )
    if exact_labels:
        matching = set(exact_labels).intersection(compatible_labels)
        selected.update(matching)
        descriptions.append(
            "labels {0}".format(", ".join(str(label) for label in exact_labels))
        )
        notes.append(
            "exact labels: {0} {1} element(s)".format(len(matching), kind)
        )
    for first_label, last_label in requested_ranges:
        existing = set(
            label
            for label in instance_labels
            if first_label <= label <= last_label
        )
        matching = existing.intersection(compatible_labels)
        selected.update(matching)
        descriptions.append("range {0}-{1}".format(first_label, last_label))
        notes.append(
            "range {0}-{1}: {2} {3} element(s); {4} other element(s) "
            "omitted".format(
                first_label,
                last_label,
                len(matching),
                kind,
                len(existing) - len(matching),
            )
        )
    if not selected:
        raise ValueError(
            "The requested {0} selections contain no {0} elements in "
            "instance '{1}'.".format(kind, instance.name)
        )
    return sorted(selected), descriptions, notes


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
        missing = []
        for requested in requested_steps:
            actual = find_step_name(available_steps, requested)
            if actual is None:
                missing.append("step '{0}' was not found".format(requested))
            elif actual not in selected:
                selected.append(actual)
        return selected, missing
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


def safe_float(value, default=0.0):
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def build_frame_records(odb, selected_steps):
    records = []
    fallback_offset = 0.0
    for step_name in selected_steps:
        step = odb.steps[step_name]
        try:
            step_start = float(step.totalTime)
        except Exception:
            step_start = fallback_offset
        largest_step_time = 0.0
        for frame_index, frame in enumerate(step.frames):
            step_time = safe_float(getattr(frame, "frameValue", 0.0))
            largest_step_time = max(largest_step_time, step_time)
            records.append(
                {
                    "step": step_name,
                    "frame": frame_index,
                    "step_time": step_time,
                    "total_time": step_start + step_time,
                    "description": getattr(frame, "description", ""),
                    "odb_frame": frame,
                    "values": {},
                }
            )
        fallback_offset = max(fallback_offset, step_start + largest_step_time)
    if not records:
        raise ValueError("The selected steps contain no frames.")
    return records


def field_value_data(value):
    try:
        data = value.data
    except Exception:
        data = value.dataDouble
    if isinstance(data, (int, float)):
        return (float(data),)
    return tuple(float(item) for item in data)


def component_index(field_output, component_name):
    try:
        labels = list(field_output.componentLabels)
    except Exception:
        labels = []
    component_upper = component_name.upper()
    for index, label in enumerate(labels):
        if str(label).upper() == component_upper:
            return index
    return None


def field_component_labels(field_output):
    try:
        return list(field_output.componentLabels)
    except Exception:
        return []


def resolve_frame_field(frame, variable):
    exact_key = repository_key(frame.fieldOutputs, variable)
    if exact_key is not None:
        field_output = frame.fieldOutputs[exact_key]
        index = component_index(field_output, variable)
        return field_output, 0 if index is None else index, str(exact_key)
    parent_name, component_name = PARENT_FIELDS[variable]
    parent_key = repository_key(frame.fieldOutputs, parent_name)
    if parent_key is None:
        return None, None, None
    field_output = frame.fieldOutputs[parent_key]
    index = component_index(field_output, component_name)
    if index is None:
        if field_component_labels(field_output):
            return None, None, None
        # Some one-component element outputs are stored as scalar S, E, ESF,
        # or CTF fields without a componentLabels sequence.
        index = 0
    return (
        field_output,
        index,
        "{0}.{1}".format(parent_key, component_name),
    )


def average_variable_by_element(
    frame, variable, instance, selected_labels
):
    if not selected_labels:
        return {}, None
    field_output, component, source = resolve_frame_field(frame, variable)
    if field_output is None:
        return {}, None
    target = set(selected_labels)
    contributions = dict((label, []) for label in selected_labels)
    try:
        values = field_output.getSubset(region=instance).values
    except Exception:
        values = field_output.values
    for value in values:
        try:
            element_label = int(value.elementLabel)
        except Exception:
            continue
        if element_label not in target:
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
        contributions[element_label].append(scalar)
    averaged = {}
    for element_label, values_for_element in contributions.items():
        if values_for_element:
            averaged[element_label] = sum(values_for_element) / float(
                len(values_for_element)
            )
    return averaged, source


def shared_node_pairs(element_by_label, connector_labels, spring_labels):
    pairs = []
    for connector_label in connector_labels:
        connector_nodes = set(
            int(label) for label in element_by_label[connector_label].connectivity
        )
        for spring_label in spring_labels:
            spring_nodes = set(
                int(label) for label in element_by_label[spring_label].connectivity
            )
            shared = sorted(connector_nodes.intersection(spring_nodes))
            if len(shared) == 1:
                pairs.append((shared[0], connector_label, spring_label))
    pairs.sort()
    return pairs


def extract_histories(
    records,
    instance,
    pipe_labels,
    connector_labels,
    spring_labels,
    pairs,
):
    source_names = dict((variable, set()) for variable in VARIABLES)
    available_counts = dict((variable, 0) for variable in VARIABLES)
    for record in records:
        frame = record["odb_frame"]
        requests = (
            ("ESF1", pipe_labels),
            ("CTF1", connector_labels),
            ("S11", spring_labels),
            ("E11", spring_labels),
        )
        for variable, labels in requests:
            values, source = average_variable_by_element(
                frame, variable, instance, labels
            )
            record["values"][variable] = values
            if source is not None:
                source_names[variable].add(source)
            available_counts[variable] += len(values)
        sums = {}
        ctf_values = record["values"]["CTF1"]
        s11_values = record["values"]["S11"]
        for node_label, connector_label, spring_label in pairs:
            if connector_label in ctf_values and spring_label in s11_values:
                sums[(node_label, connector_label, spring_label)] = (
                    ctf_values[connector_label] + s11_values[spring_label]
                )
        record["values"]["CTF1+S11"] = sums
    return source_names, available_counts


def series_definitions(
    records, pipe_labels, connector_labels, spring_labels, pairs
):
    groups = []
    for variable, labels, prefix in (
        ("ESF1", pipe_labels, "PIPE"),
        ("CTF1", connector_labels, "CONN"),
        ("S11", spring_labels, "SPRING"),
        ("E11", spring_labels, "SPRING"),
    ):
        series = []
        for label in labels:
            series.append(
                {
                    "name": "{0} {1} {2}".format(prefix, label, variable),
                    "values": [
                        record["values"][variable].get(label)
                        for record in records
                    ],
                }
            )
        groups.append((variable, series))
    sum_series = []
    for node_label, connector_label, spring_label in pairs:
        pair = (node_label, connector_label, spring_label)
        sum_series.append(
            {
                "name": "NODE {0}: CONN {1} CTF1 + SPRING {2} S11".format(
                    node_label, connector_label, spring_label
                ),
                "values": [
                    record["values"]["CTF1+S11"].get(pair)
                    for record in records
                ],
            }
        )
    groups.append(("CTF1+S11", sum_series))
    return groups


def finite_text(value):
    if value is None:
        return ""
    return "{0:.12g}".format(float(value))


def write_report(
    report_path,
    odb_path,
    instance_name,
    selected_steps,
    step_notes,
    selections,
    selection_notes,
    element_by_label,
    records,
    groups,
    pairs,
    source_names,
    available_counts,
):
    headers = ["Total Time", "Step", "Frame", "Step Time"]
    for unused_variable, series in groups:
        headers.extend(item["name"] for item in series)
    with open(report_path, "w") as report_file:
        report_file.write("*" * 112 + "\n")
        report_file.write(
            "Pipe, Connector, and Spring Element Histories, written {0}\n".format(
                datetime.datetime.now().strftime("%a %b %d %H:%M:%S %Y")
            )
        )
        report_file.write("Script version: {0}\n".format(SCRIPT_VERSION))
        report_file.write("ODB: {0}\n".format(odb_path.replace("\\", "/")))
        report_file.write("Instance: {0}\n".format(instance_name))
        report_file.write(
            "Steps: {0}; all frames included\n".format(", ".join(selected_steps))
        )
        for note in step_notes:
            report_file.write("Step note: {0}\n".format(note))
        for kind in ("pipe", "connector", "spring"):
            labels = selections[kind]
            report_file.write(
                "{0} elements ({1}): {2}\n".format(
                    kind.capitalize(),
                    len(labels),
                    ", ".join(
                        "{0} [{1}]".format(
                            label, str(element_by_label[label].type)
                        )
                        for label in labels
                    )
                    if labels
                    else "(not selected)",
                )
            )
        for note in selection_notes:
            report_file.write("Selection note: {0}\n".format(note))
        for variable in VARIABLES:
            report_file.write(
                "{0} source(s): {1}; numeric values={2}\n".format(
                    variable,
                    ", ".join(sorted(source_names[variable]))
                    if source_names[variable]
                    else "not found/not requested",
                    available_counts[variable],
                )
            )
        if pairs:
            for node_label, connector_label, spring_label in pairs:
                report_file.write(
                    "Sum pair: shared node {0}, connector {1}, spring {2}\n".format(
                        node_label, connector_label, spring_label
                    )
                )
        else:
            report_file.write(
                "Sum pair: none (no selected connector/spring element pair "
                "shares exactly one node)\n"
            )
        report_file.write(
            "CTF1+S11 is direct signed addition in each element's reported "
            "local-1 convention; confirm compatible signs/units.\n\n"
        )
        report_file.write("\t".join(headers) + "\n")
        for row_index, record in enumerate(records):
            row = [
                finite_text(record["total_time"]),
                str(record["step"]),
                str(record["frame"]),
                finite_text(record["step_time"]),
            ]
            for unused_variable, series in groups:
                row.extend(
                    finite_text(item["values"][row_index]) for item in series
                )
            report_file.write("\t".join(row) + "\n")


def xml_escape(value):
    return (
        str(value)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def excel_column_name(column_number):
    letters = []
    while column_number:
        column_number, remainder = divmod(column_number - 1, 26)
        letters.append(chr(ord("A") + remainder))
    return "".join(reversed(letters))


def inline_cell(row, column, value, style=0):
    reference = "{0}{1}".format(excel_column_name(column), row)
    return '<c r="{0}" t="inlineStr" s="{1}"><is><t>{2}</t></is></c>'.format(
        reference, style, xml_escape(value)
    )


def number_cell(row, column, value, style=0):
    if value is None:
        return ""
    reference = "{0}{1}".format(excel_column_name(column), row)
    return '<c r="{0}" s="{1}"><v>{2:.15g}</v></c>'.format(
        reference, style, float(value)
    )


def numeric_cache(values):
    points = []
    for index, value in enumerate(values):
        if value is not None and not math.isnan(value) and not math.isinf(value):
            points.append(
                '<c:pt idx="{0}"><c:v>{1:.15g}</c:v></c:pt>'.format(
                    index, float(value)
                )
            )
    return '<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="{0}"/>{1}</c:numCache>'.format(
        len(values), "".join(points)
    )


def xlsx_content_types_xml():
    chart_overrides = "".join(
        '<Override PartName="/xl/charts/chart{0}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'.format(
            index
        )
        for index in range(1, 6)
    )
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  {0}
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>""".format(chart_overrides)


def xlsx_root_relationships_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"""


def xlsx_core_properties_xml():
    timestamp = datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Pipe Connector Spring Histories</dc:title><dc:creator>Abaqus Python</dc:creator><cp:lastModifiedBy>Abaqus Python</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">{0}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">{0}</dcterms:modified>
</cp:coreProperties>""".format(timestamp)


def xlsx_app_properties_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Abaqus Python</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>History Data</vt:lpstr></vt:vector></TitlesOfParts><Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>"""


def xlsx_workbook_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="14000"/></bookViews><sheets><sheet name="History Data" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029"/></workbook>"""


def xlsx_workbook_relationships_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"""


def xlsx_styles_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="2"><numFmt numFmtId="164" formatCode="0.000000"/><numFmt numFmtId="165" formatCode="0.000000E+00"/></numFmts><fonts count="4"><font><sz val="10"/><name val="Arial"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="10"/><name val="Arial"/></font><font><b/><sz val="14"/><name val="Arial"/></font><font><i/><color rgb="FF666666"/><sz val="10"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="6"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"""


def xlsx_worksheet_xml(odb_name, records, groups):
    series = [item for unused_variable, items in groups for item in items]
    last_column = 4 + len(series)
    last_row = 4 + len(records)
    rows = [
        '<row r="1" ht="22" customHeight="1">{0}</row>'.format(
            inline_cell(
                1,
                1,
                "Pipe, Connector, and Spring Histories - {0}".format(odb_name),
                1,
            )
        ),
        '<row r="2">{0}</row>'.format(
            inline_cell(
                2,
                1,
                "All frames in selected steps; native editable Excel charts",
                2,
            )
        ),
    ]
    headers = [
        inline_cell(4, 1, "Total Time", 3),
        inline_cell(4, 2, "Step", 3),
        inline_cell(4, 3, "Frame", 3),
        inline_cell(4, 4, "Step Time", 3),
    ]
    for column, item in enumerate(series, 5):
        item["column"] = column
        headers.append(inline_cell(4, column, item["name"], 3))
    rows.append('<row r="4" ht="32" customHeight="1">{0}</row>'.format("".join(headers)))
    for row, record in enumerate(records, 5):
        cells = [
            number_cell(row, 1, record["total_time"], 4),
            inline_cell(row, 2, record["step"], 0),
            number_cell(row, 3, record["frame"], 0),
            number_cell(row, 4, record["step_time"], 4),
        ]
        series_index = row - 5
        for column, item in enumerate(series, 5):
            cells.append(number_cell(row, column, item["values"][series_index], 5))
        rows.append('<row r="{0}">{1}</row>'.format(row, "".join(cells)))
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:{0}{1}"/><sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="1" width="15" customWidth="1"/><col min="2" max="2" width="22" customWidth="1"/><col min="3" max="4" width="12" customWidth="1"/><col min="5" max="{2}" width="25" customWidth="1"/></cols><sheetData>{3}</sheetData><autoFilter ref="A4:{0}{1}"/><drawing r:id="rId1"/></worksheet>""".format(
        excel_column_name(last_column), last_row, last_column, "".join(rows)
    )


def xlsx_worksheet_relationships_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"""


def xlsx_drawing_xml(start_column):
    anchors = []
    for index in range(5):
        top_row = index * 25
        anchors.append(
            """<xdr:twoCellAnchor><xdr:from><xdr:col>{0}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{2}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>{1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{3}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="{4}" name="History Chart {5}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId{5}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>""".format(
                start_column,
                start_column + 12,
                top_row,
                top_row + 23,
                index + 2,
                index + 1,
            )
        )
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">{0}</xdr:wsDr>""".format("".join(anchors))


def xlsx_drawing_relationships_xml():
    relationships = "".join(
        '<Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart{0}.xml"/>'.format(
            index
        )
        for index in range(1, 6)
    )
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{0}</Relationships>""".format(relationships)


def chart_text(text, font_size):
    return """<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="{0}"/><a:t>{1}</a:t></a:r></a:p></c:rich></c:tx>""".format(
        font_size, xml_escape(text)
    )


def chart_series_xml(series_index, item, records):
    column = item["column"]
    column_letter = excel_column_name(column)
    start_row = 5
    end_row = 4 + len(records)
    x_values = [record["total_time"] for record in records]
    color = CHART_COLORS[series_index % len(CHART_COLORS)]
    return """<c:ser><c:idx val="{0}"/><c:order val="{0}"/><c:tx><c:strRef><c:f>'History Data'!${1}$4</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{2}</c:v></c:pt></c:strCache></c:strRef></c:tx><c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="{3}"/></a:solidFill></a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker><c:xVal><c:numRef><c:f>'History Data'!$A${4}:$A${5}</c:f>{6}</c:numRef></c:xVal><c:yVal><c:numRef><c:f>'History Data'!${1}${4}:${1}${5}</c:f>{7}</c:numRef></c:yVal><c:smooth val="0"/></c:ser>""".format(
        series_index,
        column_letter,
        xml_escape(item["name"]),
        color,
        start_row,
        end_row,
        numeric_cache(x_values),
        numeric_cache(item["values"]),
    )


def value_axis_xml(axis_id, cross_axis_id, position, title, number_format):
    return """<c:valAx><c:axId val="{0}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="{1}"/><c:title>{2}<c:layout/><c:overlay val="0"/></c:title><c:numFmt formatCode="{3}" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="{4}"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>""".format(
        axis_id,
        position,
        chart_text(title, 1000),
        xml_escape(number_format),
        cross_axis_id,
    )


def xlsx_chart_xml(title, y_title, series, records, chart_index):
    x_axis = 88000000 + chart_index * 2
    y_axis = x_axis + 1
    series_xml = "".join(
        chart_series_xml(index, item, records)
        for index, item in enumerate(series)
    )
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/><c:style val="10"/><c:chart><c:title>{0}<c:layout/><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:scatterChart><c:scatterStyle val="line"/><c:varyColors val="0"/>{1}<c:axId val="{2}"/><c:axId val="{3}"/></c:scatterChart>{4}{5}</c:plotArea><c:legend><c:legendPos val="r"/><c:layout/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>""".format(
        chart_text(title, 1200),
        series_xml,
        x_axis,
        y_axis,
        value_axis_xml(x_axis, y_axis, "b", "Total Time", "0.000000"),
        value_axis_xml(y_axis, x_axis, "l", y_title, "0.000000E+00"),
    )


def write_excel(excel_path, odb_name, records, groups):
    worksheet_xml = xlsx_worksheet_xml(odb_name, records, groups)
    chart_titles = (
        ("PIPE ESF1 History", "ESF1"),
        ("Connector CTF1 History", "CTF1"),
        ("Spring S11 History", "S11"),
        ("Spring E11 History", "E11"),
        ("CTF1 + S11 at Shared Nodes", "CTF1 + S11"),
    )
    start_column = 5 + sum(len(items) for unused_variable, items in groups)
    with zipfile.ZipFile(excel_path, "w", zipfile.ZIP_DEFLATED) as workbook:
        workbook.writestr("[Content_Types].xml", xlsx_content_types_xml())
        workbook.writestr("_rels/.rels", xlsx_root_relationships_xml())
        workbook.writestr("docProps/core.xml", xlsx_core_properties_xml())
        workbook.writestr("docProps/app.xml", xlsx_app_properties_xml())
        workbook.writestr("xl/workbook.xml", xlsx_workbook_xml())
        workbook.writestr(
            "xl/_rels/workbook.xml.rels", xlsx_workbook_relationships_xml()
        )
        workbook.writestr("xl/styles.xml", xlsx_styles_xml())
        workbook.writestr("xl/worksheets/sheet1.xml", worksheet_xml)
        workbook.writestr(
            "xl/worksheets/_rels/sheet1.xml.rels",
            xlsx_worksheet_relationships_xml(),
        )
        workbook.writestr(
            "xl/drawings/drawing1.xml", xlsx_drawing_xml(start_column)
        )
        workbook.writestr(
            "xl/drawings/_rels/drawing1.xml.rels",
            xlsx_drawing_relationships_xml(),
        )
        for index, group in enumerate(groups, 1):
            title, y_title = chart_titles[index - 1]
            workbook.writestr(
                "xl/charts/chart{0}.xml".format(index),
                xlsx_chart_xml(
                    "{0} - {1}".format(title, odb_name),
                    y_title,
                    group[1],
                    records,
                    index,
                ),
            )


def process_odb(
    odb_path,
    report_path,
    excel_path,
    instance_name,
    requested_steps,
    step_range,
    group_options,
):
    odb = None
    try:
        print("Opening: {0}".format(odb_path))
        odb = openOdb(path=odb_path, readOnly=True)
        instance_key, instance = resolve_instance(odb, instance_name)
        element_by_label = dict(
            (int(element.label), element) for element in instance.elements
        )
        selections = {}
        descriptions = {}
        selection_notes = []
        for kind in ("pipe", "connector", "spring"):
            labels, group_descriptions, notes = resolve_group_selection(
                odb,
                instance,
                element_by_label,
                kind,
                group_options[kind][0],
                group_options[kind][1],
                group_options[kind][2],
            )
            selections[kind] = labels
            descriptions[kind] = group_descriptions
            selection_notes.extend(
                "{0}: {1}".format(kind, note) for note in notes
            )
        available_steps = list(odb.steps.keys())
        selected_steps, step_notes = select_steps(
            available_steps, requested_steps, step_range
        )
        if not selected_steps:
            raise ValueError(
                "No selected steps are available. Available steps: {0}".format(
                    ", ".join(available_steps)
                )
            )
        records = build_frame_records(odb, selected_steps)
        pairs = shared_node_pairs(
            element_by_label,
            selections["connector"],
            selections["spring"],
        )
        source_names, available_counts = extract_histories(
            records,
            instance,
            selections["pipe"],
            selections["connector"],
            selections["spring"],
            pairs,
        )
        groups = series_definitions(
            records,
            selections["pipe"],
            selections["connector"],
            selections["spring"],
            pairs,
        )
        write_report(
            report_path,
            odb_path,
            instance_key,
            selected_steps,
            step_notes,
            selections,
            selection_notes,
            element_by_label,
            records,
            groups,
            pairs,
            source_names,
            available_counts,
        )
        odb_stem = os.path.splitext(os.path.basename(odb_path))[0]
        write_excel(excel_path, odb_stem, records, groups)
        print(
            "Selected: {0} pipe, {1} connector, {2} spring element(s)".format(
                len(selections["pipe"]),
                len(selections["connector"]),
                len(selections["spring"]),
            )
        )
        print(
            "History: {0} frame(s), {1} shared-node sum pair(s)".format(
                len(records), len(pairs)
            )
        )
        for variable in VARIABLES:
            if (
                (variable == "ESF1" and selections["pipe"])
                or (variable == "CTF1" and selections["connector"])
                or (variable in ("S11", "E11") and selections["spring"])
            ) and not available_counts[variable]:
                print(
                    "Warning: no {0} values were found in the selected "
                    "steps/elements.".format(variable)
                )
        if selections["connector"] and selections["spring"] and not pairs:
            print(
                "Warning: no selected connector/spring pair shares exactly "
                "one node; the sum chart is empty."
            )
        print("Wrote: {0}".format(report_path))
        print("Excel: {0}".format(excel_path))
    finally:
        if odb is not None:
            odb.close()


def print_steps(odb_path):
    odb = None
    try:
        odb = openOdb(path=odb_path, readOnly=True)
        print(odb_path)
        for index, step_name in enumerate(odb.steps.keys(), 1):
            step = odb.steps[step_name]
            print(
                "  {0:>4}  {1}  frames={2}".format(
                    index, step_name, len(step.frames)
                )
            )
    finally:
        if odb is not None:
            odb.close()


def print_element_sets(odb_path, instance_name):
    odb = None
    try:
        odb = openOdb(path=odb_path, readOnly=True)
        unused_key, instance = resolve_instance(odb, instance_name)
        element_by_label = dict(
            (int(element.label), element) for element in instance.elements
        )
        instance_labels = set(element_by_label.keys())
        print(odb_path)
        print("  Scope     Set name                         PIPE  CONN  SPRING")
        listed = 0
        repositories = (
            ("instance", instance.elementSets),
            ("assembly", odb.rootAssembly.elementSets),
        )
        for scope, repository in repositories:
            for set_name in sorted(repository.keys()):
                labels = element_set_labels_for_instance(
                    repository[set_name], instance, instance_labels
                )
                counts = dict((kind, 0) for kind in ("pipe", "connector", "spring"))
                for label in labels:
                    kind = element_kind(element_by_label[label])
                    if kind is not None:
                        counts[kind] += 1
                if any(counts.values()):
                    print(
                        "  {0:<9} {1:<32} {2:>5} {3:>5} {4:>7}".format(
                            scope,
                            set_name,
                            counts["pipe"],
                            counts["connector"],
                            counts["spring"],
                        )
                    )
                    listed += 1
        if not listed:
            print("  (no matching element sets)")
    finally:
        if odb is not None:
            odb.close()


def write_execution_log(log_path, input_dir, output_dir, jobs, failures):
    with open(log_path, "w") as log_file:
        log_file.write(
            "extract_pipe_connector_spring_history.py version {0}\n".format(
                SCRIPT_VERSION
            )
        )
        log_file.write("Time: {0}\n".format(datetime.datetime.now()))
        log_file.write("sys.argv: {0}\n".format(repr(sys.argv)))
        log_file.write("Input directory: {0}\n".format(input_dir))
        log_file.write("Output directory: {0}\n".format(output_dir))
        log_file.write("ODB jobs: {0}\n".format(len(jobs)))
        if not failures:
            log_file.write("Completed without errors.\n")
            return
        log_file.write("Failures: {0}\n".format(len(failures)))
        for odb_path, message, traceback_text in failures:
            log_file.write("\n" + "=" * 80 + "\n")
            log_file.write("ODB: {0}\n".format(odb_path))
            log_file.write("Error: {0}\n\n".format(message))
            log_file.write(traceback_text)


def main():
    print(
        "extract_pipe_connector_spring_history.py version {0}".format(
            SCRIPT_VERSION
        )
    )
    args = parse_arguments()
    input_dir = os.path.abspath(args.input_dir)
    odb_paths = find_odb_files(input_dir, args.odb)
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
    group_options = {
        "pipe": (
            args.pipe_element_set,
            args.pipe_element,
            args.pipe_element_range,
        ),
        "connector": (
            args.connector_element_set,
            args.connector_element,
            args.connector_element_range,
        ),
        "spring": (
            args.spring_element_set,
            args.spring_element,
            args.spring_element_range,
        ),
    }
    failures = []
    for odb_path, report_path, excel_path in jobs:
        try:
            process_odb(
                odb_path,
                report_path,
                excel_path,
                args.instance,
                args.steps,
                args.step_range,
                group_options,
            )
        except Exception as exc:
            traceback_text = traceback.format_exc()
            failures.append((odb_path, str(exc), traceback_text))
            print("FAILED: {0}".format(odb_path))
            print("        {0}".format(exc))
            print(traceback_text)
    log_path = os.path.join(
        output_dir, "extract_pipe_connector_spring_history.log"
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
