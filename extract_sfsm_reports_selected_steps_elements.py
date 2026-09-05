from __future__ import print_function

"""Extract ESF1, SF, and SM for selected Abaqus elements and steps.

This is a self-contained companion to the original SFSM scripts. It does not
modify or import them. Element sets, individual labels, and inclusive label
ranges may be combined; their union is processed. When no element selection
is supplied, the original PART-1-1.START node-set behavior is retained.
"""

import argparse
import datetime
import math
import os
import sys
import traceback

from abaqusConstants import ELEMENT_NODAL, ON
from odbAccess import openOdb


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPT_VERSION = "2026-09-05-r1"
DEFAULT_INSTANCE = "PART-1-1"
DEFAULT_NODE_SET = "PART-1-1.START"
DEFAULT_FRAME_INDEX = -1
FIELD_NAMES = ("ESF1", "SF", "SM")
COLUMN_SPECS = (
    ("ESF1", 0, "ESF1"),
    ("SF", 0, "SF.SF1"),
    ("SF", 1, "SF.SF2"),
    ("SF", 2, "SF.SF3"),
    ("SM", 0, "SM.SM1"),
    ("SM", 1, "SM.SM2"),
    ("SM", 2, "SM.SM3"),
)


def parse_arguments():
    parser = argparse.ArgumentParser(
        description=(
            "Write ESF1/SF/SM nodal reports for user-selected element sets, "
            "element labels/ranges, and ODB steps."
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
        help=(
            "Instance containing explicitly selected elements "
            "(default: PART-1-1)."
        ),
    )
    parser.add_argument(
        "--node-set",
        default=DEFAULT_NODE_SET,
        help=(
            "Qualified fallback node set used only when no element option is "
            "supplied (default: PART-1-1.START)."
        ),
    )
    parser.add_argument(
        "--element-set",
        action="append",
        default=[],
        metavar="SET",
        help=(
            "Select an instance- or assembly-level element set. May be "
            "repeated; all element selections form a union."
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
            "Select one or more exact element labels. May be repeated; all "
            "element selections form a union."
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
            "Select an inclusive element-label range. May be repeated; all "
            "element selections form a union."
        ),
    )
    parser.add_argument(
        "--frame-index",
        type=int,
        default=DEFAULT_FRAME_INDEX,
        help=(
            "Zero-based frame index. Use -1 for the last frame containing "
            "ESF1, SF, and SM (default: -1)."
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
            "List instance- and assembly-level element sets for --instance "
            "and exit without writing reports."
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
            "Inclusive range using 1-based step numbers or exact first and "
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


def repository_key(repository, requested_name):
    if requested_name in repository:
        return requested_name
    requested_upper = requested_name.upper()
    for key in repository.keys():
        if str(key).upper() == requested_upper:
            return key
    return None


def default_report_name(odb_path):
    stem = os.path.splitext(os.path.basename(odb_path))[0]
    return stem + "_SFSM_SELECTED_ELEMENTS.rpt"


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
    instance_key = repository_key(instance.elementSets, requested_name)
    if instance_key is not None:
        labels = element_set_labels_for_instance(
            instance.elementSets[instance_key], instance, instance_labels
        )
        return labels, "instance set '{0}'".format(instance_key)

    assembly_key = repository_key(
        odb.rootAssembly.elementSets, requested_name
    )
    if assembly_key is not None:
        labels = element_set_labels_for_instance(
            odb.rootAssembly.elementSets[assembly_key],
            instance,
            instance_labels,
        )
        return labels, "assembly set '{0}'".format(assembly_key)

    available = sorted(
        set(
            list(instance.elementSets.keys())
            + list(odb.rootAssembly.elementSets.keys())
        )
    )
    raise ValueError(
        "Element set '{0}' was not found. Available element sets: {1}".format(
            requested_name, ", ".join(available) if available else "(none)"
        )
    )


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


def resolve_instance_node_set(odb, qualified_name):
    if "." not in qualified_name:
        raise ValueError(
            "--node-set must include instance and set names, for example "
            "PART-1-1.START"
        )
    instance_name, set_name = qualified_name.split(".", 1)
    instance_key, instance = resolve_instance(odb, instance_name)
    set_key = repository_key(instance.nodeSets, set_name)
    if set_key is None:
        raise ValueError(
            "Node set '{0}' is missing from instance '{1}'. Available sets: "
            "{2}".format(
                set_name,
                instance_key,
                ", ".join(instance.nodeSets.keys()),
            )
        )
    node_labels = sorted(
        set(int(node.label) for node in instance.nodeSets[set_key].nodes)
    )
    if not node_labels:
        raise ValueError("Node set '{0}' is empty.".format(qualified_name))
    return (
        instance_key,
        instance,
        node_labels,
        None,
        "node set {0}.{1}".format(instance_key, set_key),
        [],
    )


def flattened_requested_elements(requested):
    labels = []
    for group in requested:
        labels.extend(group)
    return labels


def resolve_output_region(
    odb,
    instance_name,
    node_set_name,
    requested_sets,
    requested_elements,
    requested_ranges,
):
    exact_labels = flattened_requested_elements(requested_elements)
    explicit_selection = bool(
        requested_sets or exact_labels or requested_ranges
    )
    if not explicit_selection:
        return resolve_instance_node_set(odb, node_set_name)

    instance_key, instance = resolve_instance(odb, instance_name)
    element_by_label = dict(
        (int(element.label), element) for element in instance.elements
    )
    instance_labels = set(element_by_label.keys())
    selected_labels = set()
    descriptions = []
    notes = []

    for requested_name in requested_sets:
        set_labels, resolved_name = resolve_element_set(
            odb, instance, requested_name, instance_labels
        )
        selected_labels.update(set_labels)
        descriptions.append(resolved_name)
        notes.append(
            "{0}: {1} element(s) selected".format(
                resolved_name, len(set_labels)
            )
        )

    missing_exact = []
    for label in exact_labels:
        if label in instance_labels:
            selected_labels.add(label)
        else:
            missing_exact.append(label)
    if exact_labels:
        descriptions.append(
            "element labels {0}".format(
                ", ".join(str(label) for label in exact_labels)
            )
        )
        notes.append(
            "exact labels: {0} existing element(s) selected".format(
                len(set(exact_labels).intersection(instance_labels))
            )
        )
    if missing_exact:
        raise ValueError(
            "Requested element labels were not found in instance '{0}': "
            "{1}".format(
                instance_key,
                ", ".join(str(label) for label in sorted(set(missing_exact))),
            )
        )

    for first_label, last_label in requested_ranges:
        range_labels = set(
            label
            for label in instance_labels
            if first_label <= label <= last_label
        )
        selected_labels.update(range_labels)
        descriptions.append(
            "element range {0}-{1}".format(first_label, last_label)
        )
        notes.append(
            "range {0}-{1}: {2} element(s) selected".format(
                first_label, last_label, len(range_labels)
            )
        )

    if not selected_labels:
        raise ValueError(
            "The requested element selections contain no elements in "
            "instance '{0}'.".format(instance_key)
        )

    node_labels = sorted(
        set(
            int(node_label)
            for element_label in selected_labels
            for node_label in element_by_label[element_label].connectivity
        )
    )
    if not node_labels:
        raise ValueError("The selected elements contain no nodes.")

    return (
        instance_key,
        instance,
        node_labels,
        selected_labels,
        "; ".join(descriptions),
        notes,
    )


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
                "range starts at {0}, but this ODB contains only {1} "
                "step(s)".format(start_number, len(available_steps))
            ]
        actual_end = min(end_number, len(available_steps))
        notes = []
        if end_number > len(available_steps):
            notes.append(
                "range end {0} was limited to the last available step "
                "({1})".format(end_number, len(available_steps))
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


def frame_has_fields(frame):
    return all(name in frame.fieldOutputs for name in FIELD_NAMES)


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
                    "frame index {0} is unavailable; the step has {1} "
                    "frame(s)".format(frame_index, len(frames)),
                )
            )
            continue
        else:
            candidate_indices = (frame_index,)

        selected_frame = None
        for candidate_index in candidate_indices:
            frame = frames[candidate_index]
            if frame_has_fields(frame):
                selected_frame = (step_name, candidate_index, frame)
                break
        if selected_frame is None:
            skipped.append(
                (step_name, "no requested frame contains ESF1, SF, and SM")
            )
        else:
            selected.append(selected_frame)
    return selected, skipped


def field_value_data(value):
    try:
        data = value.data
    except Exception:
        data = value.dataDouble
    if isinstance(data, (int, float)):
        return (float(data),)
    return tuple(float(item) for item in data)


def average_element_nodal_field(
    field_output,
    instance_name,
    node_labels,
    selected_element_labels,
):
    target_nodes = set(node_labels)
    contributions = dict((label, []) for label in node_labels)
    element_nodal = field_output.getSubset(
        position=ELEMENT_NODAL, readOnly=ON
    )

    for value in element_nodal.values:
        try:
            node_label = int(value.nodeLabel)
        except Exception:
            continue
        if node_label not in target_nodes:
            continue
        value_instance = getattr(value, "instance", None)
        if (
            value_instance is not None
            and value_instance.name.upper() != instance_name.upper()
        ):
            continue
        if selected_element_labels is not None:
            try:
                element_label = int(value.elementLabel)
            except Exception:
                continue
            if element_label not in selected_element_labels:
                continue
        contributions[node_label].append(field_value_data(value))

    averaged = {}
    counts = {}
    missing_nodes = []
    for node_label in node_labels:
        values = contributions[node_label]
        if not values:
            missing_nodes.append(node_label)
            continue
        component_count = len(values[0])
        if any(len(value) != component_count for value in values):
            raise ValueError(
                "Inconsistent component count for field '{0}' at node "
                "{1}.".format(field_output.name, node_label)
            )
        averaged[node_label] = tuple(
            sum(value[index] for value in values) / float(len(values))
            for index in range(component_count)
        )
        counts[node_label] = len(values)
    if missing_nodes:
        preview = ", ".join(str(label) for label in missing_nodes[:20])
        raise ValueError(
            "No selected element-nodal values were found for field '{0}' "
            "at {1} node(s), beginning with: {2}".format(
                field_output.name, len(missing_nodes), preview
            )
        )
    return averaged, counts


def extract_frame_data(
    frame, instance_name, node_labels, selected_element_labels
):
    fields = {}
    counts_by_field = {}
    for field_name in FIELD_NAMES:
        averaged, counts = average_element_nodal_field(
            frame.fieldOutputs[field_name],
            instance_name,
            node_labels,
            selected_element_labels,
        )
        fields[field_name] = averaged
        counts_by_field[field_name] = counts

    rows = []
    for node_label in node_labels:
        values = []
        for field_name, component_index, unused_title in COLUMN_SPECS:
            field_values = fields[field_name][node_label]
            if component_index >= len(field_values):
                raise ValueError(
                    "Field '{0}' at node {1} has {2} component(s); "
                    "component {3} was requested.".format(
                        field_name,
                        node_label,
                        len(field_values),
                        component_index + 1,
                    )
                )
            values.append(field_values[component_index])
        rows.append((node_label, values))
    return rows, counts_by_field


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
        digits_before_decimal = int(
            math.floor(math.log10(abs(mantissa)))
        ) + 1
        decimal_places = max(0, 6 - digits_before_decimal)
        mantissa_text = ("{0:.%df}" % decimal_places).format(mantissa)
    return "{0}E{1:+03d}".format(mantissa_text, exponent)


def write_table(report_file, rows):
    titles = [spec[2] for spec in COLUMN_SPECS]
    header = "{0:>16}".format("Node Label")
    header += "".join("{0:>16}".format(title) for title in titles)
    locations = "{0:>16}".format("")
    locations += "".join(
        "{0:>16}".format("@Loc 1") for unused in titles
    )
    report_file.write(header + "\n")
    report_file.write(locations + "\n")
    report_file.write("-" * len(header) + "\n")
    for node_label, values in rows:
        line = "{0:>16d}".format(node_label)
        line += "".join(
            "{0:>16}".format(engineering_format(value))
            for value in values
        )
        report_file.write(line + "\n")

    columns = list(zip(*(values for unused_label, values in rows)))
    minimums = [min(column) for column in columns]
    maximums = [max(column) for column in columns]
    totals = [sum(column) for column in columns]
    min_nodes = [
        rows[list(column).index(min(column))][0] for column in columns
    ]
    max_nodes = [
        rows[list(column).index(max(column))][0] for column in columns
    ]
    report_file.write("\n\n")
    report_file.write("{0:>16}".format("Minimum"))
    report_file.write(
        "".join(
            "{0:>16}".format(engineering_format(value))
            for value in minimums
        )
    )
    report_file.write("\n")
    report_file.write("{0:>16}".format("At Node"))
    report_file.write(
        "".join("{0:>16d}".format(value) for value in min_nodes)
    )
    report_file.write("\n\n")
    report_file.write("{0:>16}".format("Maximum"))
    report_file.write(
        "".join(
            "{0:>16}".format(engineering_format(value))
            for value in maximums
        )
    )
    report_file.write("\n")
    report_file.write("{0:>16}".format("At Node"))
    report_file.write(
        "".join("{0:>16d}".format(value) for value in max_nodes)
    )
    report_file.write("\n\n")
    report_file.write("{0:>16}".format("Total"))
    report_file.write(
        "".join(
            "{0:>16}".format(engineering_format(value))
            for value in totals
        )
    )
    report_file.write("\n\n\n")


def write_report(
    odb_path,
    report_path,
    instance_name,
    node_set_name,
    requested_sets,
    requested_elements,
    requested_ranges,
    requested_steps,
    step_range,
    frame_index,
):
    odb = None
    try:
        print("Opening: {0}".format(odb_path))
        odb = openOdb(path=odb_path, readOnly=True)
        (
            actual_instance_name,
            unused_instance,
            node_labels,
            selected_element_labels,
            region_description,
            region_notes,
        ) = resolve_output_region(
            odb,
            instance_name,
            node_set_name,
            requested_sets,
            requested_elements,
            requested_ranges,
        )

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
            raise ValueError(
                "No usable steps were found for ESF1, SF, and SM. {0}".format(
                    details
                )
            )

        extracted_frames = []
        max_contributions = 0
        for step_name, selected_frame_index, frame in selected_frames:
            rows, counts_by_field = extract_frame_data(
                frame,
                actual_instance_name,
                node_labels,
                selected_element_labels,
            )
            for field_counts in counts_by_field.values():
                if field_counts:
                    max_contributions = max(
                        max_contributions, max(field_counts.values())
                    )
            extracted_frames.append(
                (step_name, selected_frame_index, frame, rows)
            )

        with open(report_path, "w") as report_file:
            report_file.write("*" * 112 + "\n")
            report_file.write(
                "Selected-Element SFSM Field Output Report, written {0}\n".format(
                    datetime.datetime.now().strftime(
                        "%a %b %d %H:%M:%S %Y"
                    )
                )
            )
            report_file.write(
                "Script version: {0}\n".format(SCRIPT_VERSION)
            )
            report_file.write(
                "Instance: {0}\n".format(actual_instance_name)
            )
            report_file.write(
                "Region selection: {0}\n".format(region_description)
            )
            if selected_element_labels is None:
                report_file.write(
                    "Selection mode: fallback node set; contributions from "
                    "all connected elements may be averaged\n"
                )
            else:
                report_file.write(
                    "Selection mode: {0} selected element(s), {1} unique "
                    "node(s)\n".format(
                        len(selected_element_labels), len(node_labels)
                    )
                )
                report_file.write(
                    "Averaging scope: selected elements only\n"
                )
            for note in region_notes:
                report_file.write("Selection note: {0}\n".format(note))
            for note in step_notes:
                report_file.write("Step note: {0}\n".format(note))
            report_file.write("\n")

            for step_name, selected_frame_index, frame, rows in extracted_frames:
                report_file.write("Source 1\n")
                report_file.write("---------\n\n")
                report_file.write(
                    "   ODB: {0}\n".format(odb_path.replace("\\", "/"))
                )
                report_file.write("   Step: {0}\n".format(step_name))
                report_file.write(
                    "   Frame index: {0}\n".format(selected_frame_index)
                )
                report_file.write(
                    "   Frame: {0}\n\n".format(frame.description)
                )
                report_file.write(
                    "Loc 1 : Element-nodal values averaged at selected "
                    "nodes\n\n"
                )
                report_file.write(
                    'Output sorted by column "Node Label".\n\n'
                )
                report_file.write(
                    "Field Output reported for region: {0}\n".format(
                        region_description
                    )
                )
                report_file.write(
                    "   Computation algorithm: extrapolate to element "
                    "nodes, then arithmetic average\n"
                )
                if selected_element_labels is not None:
                    report_file.write(
                        "   Contributions restricted to selected elements\n"
                    )
                report_file.write("\n")
                write_table(report_file, rows)

        for note in step_notes:
            print("Note: {0}.".format(note))
        for skipped_step, reason in skipped_steps:
            print("Skipped step '{0}': {1}.".format(skipped_step, reason))
        for note in region_notes:
            print("Element selection: {0}.".format(note))
        print("Selected nodes: {0}".format(len(node_labels)))
        if selected_element_labels is not None:
            print(
                "Selected elements: {0}".format(
                    len(selected_element_labels)
                )
            )
        if max_contributions > 1:
            print(
                "Note: up to {0} selected element-nodal contributions were "
                "averaged at one node.".format(max_contributions)
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
        instance_labels = set(int(element.label) for element in instance.elements)
        print(odb_path)
        print("  Instance: {0}".format(instance_key))
        listed = 0
        for set_name in sorted(instance.elementSets.keys()):
            labels = element_set_labels_for_instance(
                instance.elementSets[set_name], instance, instance_labels
            )
            if labels:
                print(
                    "  {0:>8}  {1}  elements={2}".format(
                        "instance", set_name, len(labels)
                    )
                )
                listed += 1
        for set_name in sorted(odb.rootAssembly.elementSets.keys()):
            labels = element_set_labels_for_instance(
                odb.rootAssembly.elementSets[set_name],
                instance,
                instance_labels,
            )
            if labels:
                print(
                    "  {0:>8}  {1}  elements={2}".format(
                        "assembly", set_name, len(labels)
                    )
                )
                listed += 1
        if not listed:
            print("  (no element sets contain elements from this instance)")
    finally:
        if odb is not None:
            odb.close()


def write_execution_log(log_path, input_dir, output_dir, jobs, failures):
    with open(log_path, "w") as log_file:
        log_file.write(
            "extract_sfsm_reports_selected_steps_elements.py version "
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
        "extract_sfsm_reports_selected_steps_elements.py version {0}".format(
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
    failures = []

    for odb_path, report_path in jobs:
        try:
            write_report(
                odb_path=odb_path,
                report_path=report_path,
                instance_name=args.instance,
                node_set_name=args.node_set,
                requested_sets=args.element_set,
                requested_elements=args.element,
                requested_ranges=args.element_range,
                requested_steps=args.steps,
                step_range=args.step_range,
                frame_index=args.frame_index,
            )
        except Exception as exc:
            traceback_text = traceback.format_exc()
            failures.append((odb_path, str(exc), traceback_text))
            print("FAILED: {0}".format(odb_path))
            print("        {0}".format(exc))
            print(traceback_text)

    log_path = os.path.join(
        output_dir, "extract_sfsm_reports_selected_steps_elements.log"
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
