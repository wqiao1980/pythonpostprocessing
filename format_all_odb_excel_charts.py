from __future__ import print_function

"""Batch-format Excel charts with a built-in or user-selected template.

The script scans Excel workbooks created from ODB postprocessing, preserves
their chart data, formulas, titles, axis titles, and exact series names, and
changes only chart formatting. It uses only the Python standard library and
does not require Abaqus, Excel, openpyxl, or pywin32.
"""

import argparse
import copy
import datetime
import fnmatch
import os
import posixpath
import re
import shutil
import sys
import tempfile
import traceback
import zipfile
import xml.etree.ElementTree as ET


SCRIPT_VERSION = "2026-09-19-r4"
DEFAULT_PATTERNS = ("*_along_path.xlsx",)
DEFAULT_SUFFIX = "_formatted"
CHART_COLORS = (
    "1565C0",
    "C00000",
    "70AD47",
    "7030A0",
    "ED7D31",
    "00B0F0",
    "7F6000",
    "A5A5A5",
)

NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_XDR = (
    "http://schemas.openxmlformats.org/drawingml/2006/"
    "spreadsheetDrawing"
)

CHART_TYPE_NAMES = (
    "area3DChart",
    "areaChart",
    "bar3DChart",
    "barChart",
    "bubbleChart",
    "doughnutChart",
    "line3DChart",
    "lineChart",
    "ofPieChart",
    "pie3DChart",
    "pieChart",
    "radarChart",
    "scatterChart",
    "stockChart",
    "surface3DChart",
    "surfaceChart",
)
AXIS_NAMES = ("catAx", "dateAx", "serAx", "valAx")
SERIES_FORMAT_NAMES = (
    "spPr",
    "invertIfNegative",
    "pictureOptions",
    "marker",
    "dPt",
    "dLbls",
    "trendline",
    "errBars",
    "explosion",
    "bubble3D",
    "shape",
    "smooth",
)
SERIES_CHILD_ORDER = (
    "idx",
    "order",
    "tx",
    "spPr",
    "invertIfNegative",
    "pictureOptions",
    "marker",
    "explosion",
    "dPt",
    "dLbls",
    "trendline",
    "errBars",
    "cat",
    "xVal",
    "val",
    "yVal",
    "bubbleSize",
    "bubble3D",
    "shape",
    "smooth",
    "extLst",
)

for prefix, namespace in (
    ("c", NS_C),
    ("a", NS_A),
    ("r", NS_R),
    ("xdr", NS_XDR),
):
    try:
        ET.register_namespace(prefix, namespace)
    except AttributeError:
        pass


def q(namespace, local_name):
    return "{{{0}}}{1}".format(namespace, local_name)


def parse_arguments():
    script_name = globals().get("__file__") or (
        sys.argv[0] if sys.argv and sys.argv[0] else os.getcwd()
    )
    script_dir = os.path.dirname(os.path.abspath(script_name))
    parser = argparse.ArgumentParser(
        description=(
            "Batch-format native Excel charts in ODB postprocessing "
            "workbooks using the built-in style or a manually formatted "
            "template chart."
        )
    )
    parser.add_argument(
        "--input-dir",
        default=script_dir,
        help="Folder containing workbooks (default: script folder).",
    )
    parser.add_argument(
        "--pattern",
        action="append",
        default=None,
        metavar="WILDCARD",
        help=(
            "Filename wildcard to include. May be repeated. Default: "
            "*_along_path.xlsx"
        ),
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Search --input-dir and all subfolders.",
    )
    parser.add_argument(
        "--output-dir",
        default=None,
        help=(
            "Write formatted copies under this folder while preserving "
            "relative subfolders. Default: beside each source workbook."
        ),
    )
    parser.add_argument(
        "--suffix",
        default=DEFAULT_SUFFIX,
        help=(
            "Suffix added before .xlsx for formatted copies "
            "(default: _formatted)."
        ),
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Replace an existing formatted-copy output.",
    )
    parser.add_argument(
        "--in-place",
        action="store_true",
        help="Replace each source workbook after formatting it.",
    )
    parser.add_argument(
        "--no-backup",
        action="store_true",
        help=(
            "With --in-place, do not save a *_before_chart_format.xlsx "
            "backup. A backup is created by default."
        ),
    )
    parser.add_argument(
        "--no-resize",
        action="store_true",
        help="Keep every chart's existing size and anchor.",
    )
    parser.add_argument(
        "--chart-width-columns",
        type=int,
        default=17,
        help="Chart width in worksheet columns (default: 17).",
    )
    parser.add_argument(
        "--chart-height-rows",
        type=int,
        default=33,
        help="Chart height in worksheet rows (default: 33).",
    )
    parser.add_argument(
        "--font",
        default="Arial",
        help="Chart font family (default: Arial).",
    )
    parser.add_argument(
        "--template-workbook",
        default=None,
        help=(
            "Workbook containing a manually formatted native Excel chart. "
            "The selected chart's formatting is applied to compatible "
            "target charts."
        ),
    )
    parser.add_argument(
        "--template-chart",
        type=int,
        default=1,
        metavar="NUMBER",
        help=(
            "One-based chart number in --template-workbook (default: 1). "
            "Use --list-template-charts to see the available charts."
        ),
    )
    parser.add_argument(
        "--target-chart-title",
        action="append",
        default=None,
        metavar="WILDCARD",
        help=(
            "In template mode, format only charts whose title matches this "
            "wildcard. May be repeated; for example, *CPRESS*."
        ),
    )
    parser.add_argument(
        "--list-template-charts",
        action="store_true",
        help=(
            "List chart numbers, titles, types, and approximate sizes in "
            "the template workbook, then exit."
        ),
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="List input/output paths without creating or changing files.",
    )
    args = parser.parse_args()

    if args.in_place and args.output_dir:
        parser.error("--in-place and --output-dir cannot be used together")
    if not args.in_place and not args.suffix:
        parser.error("--suffix cannot be empty unless --in-place is used")
    if args.chart_width_columns < 4:
        parser.error("--chart-width-columns must be at least 4")
    if args.chart_height_rows < 8:
        parser.error("--chart-height-rows must be at least 8")
    if args.template_chart < 1:
        parser.error("--template-chart must be at least 1")
    if args.list_template_charts and not args.template_workbook:
        parser.error("--list-template-charts requires --template-workbook")
    if args.target_chart_title and not args.template_workbook:
        parser.error("--target-chart-title requires --template-workbook")
    return args


def matches_patterns(filename, patterns):
    lower_name = filename.lower()
    return any(
        fnmatch.fnmatch(lower_name, pattern.lower())
        for pattern in patterns
    )


def find_workbooks(
    input_dir, patterns, recursive, suffix, excluded_paths=None
):
    workbooks = []
    excluded = set(
        os.path.normcase(os.path.abspath(path))
        for path in (excluded_paths or ())
    )
    if recursive:
        directory_iterator = os.walk(input_dir)
    else:
        directory_iterator = ((input_dir, (), os.listdir(input_dir)),)

    suffix_lower = suffix.lower()
    for directory, unused_subdirs, filenames in directory_iterator:
        for filename in filenames:
            if filename.startswith("~$"):
                continue
            if not filename.lower().endswith(".xlsx"):
                continue
            if filename[:-5].lower().endswith("_before_chart_format"):
                continue
            if suffix_lower and filename[:-5].lower().endswith(suffix_lower):
                continue
            if not matches_patterns(filename, patterns):
                continue
            full_path = os.path.abspath(os.path.join(directory, filename))
            if os.path.normcase(full_path) in excluded:
                continue
            workbooks.append(full_path)
    return sorted(set(workbooks), key=lambda value: value.lower())


def output_path_for(source_path, input_dir, output_dir, suffix, in_place):
    if in_place:
        return source_path
    stem, extension = os.path.splitext(os.path.basename(source_path))
    output_name = stem + suffix + extension
    if output_dir:
        relative_folder = os.path.dirname(os.path.relpath(source_path, input_dir))
        target_folder = os.path.join(output_dir, relative_folder)
    else:
        target_folder = os.path.dirname(source_path)
    return os.path.abspath(os.path.join(target_folder, output_name))


def remove_children(parent, tags):
    for child in list(parent):
        if child.tag in tags:
            parent.remove(child)


def insert_before(parent, element, before_tags):
    children = list(parent)
    for index, child in enumerate(children):
        if child.tag in before_tags:
            parent.insert(index, element)
            return
    parent.append(element)


def solid_fill(parent, color):
    fill = ET.SubElement(parent, q(NS_A, "solidFill"))
    ET.SubElement(fill, q(NS_A, "srgbClr"), {"val": color})
    return fill


def line_element(color, width):
    line = ET.Element(q(NS_A, "ln"), {"w": str(width)})
    solid_fill(line, color)
    ET.SubElement(line, q(NS_A, "prstDash"), {"val": "solid"})
    return line


def shape_properties(fill_color, line_color, line_width):
    shape = ET.Element(q(NS_C, "spPr"))
    if fill_color:
        solid_fill(shape, fill_color)
    else:
        ET.SubElement(shape, q(NS_A, "noFill"))
    if line_color:
        shape.append(line_element(line_color, line_width))
    return shape


def replace_shape_properties(parent, fill_color, line_color, line_width, before):
    existing = parent.find(q(NS_C, "spPr"))
    if existing is not None:
        parent.remove(existing)
    replacement = shape_properties(fill_color, line_color, line_width)
    insert_before(parent, replacement, before)


def add_typeface(parent, font_name):
    remove_children(
        parent,
        (q(NS_A, "latin"), q(NS_A, "ea"), q(NS_A, "cs")),
    )
    ET.SubElement(parent, q(NS_A, "latin"), {"typeface": font_name})
    ET.SubElement(parent, q(NS_A, "ea"), {"typeface": font_name})
    ET.SubElement(parent, q(NS_A, "cs"), {"typeface": font_name})


def text_properties(font_name, point_size):
    text = ET.Element(q(NS_C, "txPr"))
    ET.SubElement(text, q(NS_A, "bodyPr"), {"anchorCtr": "1"})
    ET.SubElement(text, q(NS_A, "lstStyle"))
    paragraph = ET.SubElement(text, q(NS_A, "p"))
    paragraph_properties = ET.SubElement(paragraph, q(NS_A, "pPr"))
    default_run = ET.SubElement(
        paragraph_properties,
        q(NS_A, "defRPr"),
        {"sz": str(int(point_size * 100))},
    )
    add_typeface(default_run, font_name)
    end_run = ET.SubElement(
        paragraph,
        q(NS_A, "endParaRPr"),
        {"lang": "en-US", "sz": str(int(point_size * 100))},
    )
    add_typeface(end_run, font_name)
    return text


def replace_text_properties(parent, font_name, point_size, before):
    existing = parent.find(q(NS_C, "txPr"))
    if existing is not None:
        parent.remove(existing)
    replacement = text_properties(font_name, point_size)
    insert_before(parent, replacement, before)


def style_existing_rich_text(parent, font_name, point_size):
    size_text = str(int(point_size * 100))
    run_tags = (
        q(NS_A, "rPr"),
        q(NS_A, "defRPr"),
        q(NS_A, "endParaRPr"),
    )
    for element in parent.iter():
        if element.tag in run_tags:
            element.set("sz", size_text)
            add_typeface(element, font_name)


def ensure_value_child(parent, local_name, value, before=()):
    tag = q(NS_C, local_name)
    element = parent.find(tag)
    if element is None:
        element = ET.Element(tag)
        insert_before(parent, element, tuple(q(NS_C, item) for item in before))
    element.set("val", str(value))
    return element


def style_series(series, color):
    replace_shape_properties(
        series,
        None,
        color,
        28575,
        (
            q(NS_C, "marker"),
            q(NS_C, "dPt"),
            q(NS_C, "dLbls"),
            q(NS_C, "xVal"),
            q(NS_C, "cat"),
            q(NS_C, "val"),
            q(NS_C, "yVal"),
        ),
    )

    marker = series.find(q(NS_C, "marker"))
    if marker is None:
        marker = ET.Element(q(NS_C, "marker"))
        insert_before(
            series,
            marker,
            (
                q(NS_C, "dPt"),
                q(NS_C, "dLbls"),
                q(NS_C, "xVal"),
                q(NS_C, "cat"),
                q(NS_C, "val"),
                q(NS_C, "yVal"),
            ),
        )
    remove_children(marker, tuple(child.tag for child in list(marker)))
    ET.SubElement(marker, q(NS_C, "symbol"), {"val": "circle"})
    ET.SubElement(marker, q(NS_C, "size"), {"val": "5"})
    marker_shape = ET.SubElement(marker, q(NS_C, "spPr"))
    solid_fill(marker_shape, color)
    marker_shape.append(line_element(color, 12700))
    ensure_value_child(series, "smooth", "0")


def chart_text_upper(root):
    values = []
    for element in root.iter():
        if element.text:
            values.append(element.text)
    return " ".join(values).upper()


def set_axis_minimum(axis, minimum):
    scaling = axis.find(q(NS_C, "scaling"))
    if scaling is None:
        scaling = ET.Element(q(NS_C, "scaling"))
        insert_before(axis, scaling, (q(NS_C, "delete"), q(NS_C, "axPos")))
    minimum_element = scaling.find(q(NS_C, "min"))
    if minimum_element is None:
        minimum_element = ET.SubElement(scaling, q(NS_C, "min"))
    minimum_element.set("val", str(minimum))


def style_axis(axis, chart_text, font_name):
    remove_children(
        axis,
        (q(NS_C, "majorGridlines"), q(NS_C, "minorGridlines")),
    )
    ensure_value_child(axis, "majorTickMark", "out", ("minorTickMark", "tickLblPos"))
    ensure_value_child(axis, "minorTickMark", "none", ("tickLblPos",))
    ensure_value_child(axis, "tickLblPos", "nextTo")

    replace_shape_properties(
        axis,
        None,
        "555555",
        12700,
        (q(NS_C, "txPr"), q(NS_C, "crossAx")),
    )
    replace_text_properties(
        axis,
        font_name,
        9,
        (q(NS_C, "crossAx"),),
    )

    title = axis.find(q(NS_C, "title"))
    if title is not None:
        style_existing_rich_text(title, font_name, 11)

    axis_position = axis.find(q(NS_C, "axPos"))
    position = axis_position.get("val") if axis_position is not None else ""
    num_fmt = axis.find(q(NS_C, "numFmt"))
    if num_fmt is None:
        num_fmt = ET.Element(q(NS_C, "numFmt"))
        insert_before(
            axis,
            num_fmt,
            (q(NS_C, "majorTickMark"), q(NS_C, "minorTickMark")),
        )
    if position == "b":
        num_fmt.set("formatCode", "#,##0")
        set_axis_minimum(axis, 0)
    elif "COPEN" in chart_text:
        num_fmt.set("formatCode", "0.0000E+00")
    elif "CPRESS" in chart_text:
        num_fmt.set("formatCode", "#,##0")
        set_axis_minimum(axis, 0)
    else:
        num_fmt.set("formatCode", "#,##0.00")
    num_fmt.set("sourceLinked", "0")


def format_chart_xml(xml_data, font_name):
    root = ET.fromstring(xml_data)
    chart = root.find(q(NS_C, "chart"))
    if chart is None:
        return xml_data, 0

    ensure_value_child(root, "roundedCorners", "0", ("style", "chart"))
    ensure_value_child(root, "style", "10", ("chart",))
    replace_shape_properties(
        root,
        "FFFFFF",
        "808080",
        12700,
        (q(NS_C, "txPr"), q(NS_C, "externalData"), q(NS_C, "printSettings")),
    )

    title = chart.find(q(NS_C, "title"))
    if title is not None:
        style_existing_rich_text(title, font_name, 14)
        replace_text_properties(
            title,
            font_name,
            14,
            (q(NS_C, "extLst"),),
        )

    plot_area = chart.find(q(NS_C, "plotArea"))
    if plot_area is None:
        return xml_data, 0
    replace_shape_properties(
        plot_area,
        "FFFFFF",
        "222222",
        12700,
        (q(NS_C, "extLst"),),
    )

    chart_types = (
        "scatterChart",
        "lineChart",
        "barChart",
        "areaChart",
        "radarChart",
    )
    series = []
    for chart_type_name in chart_types:
        chart_type = plot_area.find(q(NS_C, chart_type_name))
        if chart_type is None:
            continue
        if chart_type_name == "scatterChart":
            ensure_value_child(
                chart_type,
                "scatterStyle",
                "lineMarker",
                ("varyColors", "ser", "dLbls", "axId"),
            )
        ensure_value_child(
            chart_type,
            "varyColors",
            "0",
            ("ser", "dLbls", "axId"),
        )
        series.extend(chart_type.findall(q(NS_C, "ser")))

    for series_index, series_element in enumerate(series):
        style_series(
            series_element,
            CHART_COLORS[series_index % len(CHART_COLORS)],
        )

    text_upper = chart_text_upper(root)
    for axis_name in ("valAx", "catAx", "dateAx"):
        for axis in plot_area.findall(q(NS_C, axis_name)):
            style_axis(axis, text_upper, font_name)

    legend = chart.find(q(NS_C, "legend"))
    if legend is None:
        legend = ET.Element(q(NS_C, "legend"))
        insert_before(
            chart,
            legend,
            (
                q(NS_C, "plotVisOnly"),
                q(NS_C, "dispBlanksAs"),
                q(NS_C, "showDLblsOverMax"),
            ),
        )
    ensure_value_child(legend, "legendPos", "t", ("legendEntry", "layout", "overlay"))
    ensure_value_child(legend, "overlay", "0")
    replace_text_properties(legend, font_name, 10, (q(NS_C, "extLst"),))
    ensure_value_child(chart, "plotVisOnly", "1", ("dispBlanksAs", "showDLblsOverMax"))
    ensure_value_child(chart, "dispBlanksAs", "gap", ("showDLblsOverMax",))

    xml_body = ET.tostring(root, encoding="utf-8")
    declaration = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    return declaration + xml_body, len(series)


def local_name(tag):
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def chart_type_elements(plot_area):
    if plot_area is None:
        return []
    return [
        child
        for child in list(plot_area)
        if local_name(child.tag) in CHART_TYPE_NAMES
    ]


def chart_signature(root):
    chart = root.find(q(NS_C, "chart"))
    plot_area = chart.find(q(NS_C, "plotArea")) if chart is not None else None
    return tuple(
        local_name(element.tag)
        for element in chart_type_elements(plot_area)
    )


def chart_title_text(root):
    chart = root.find(q(NS_C, "chart"))
    title = chart.find(q(NS_C, "title")) if chart is not None else None
    if title is None:
        return "(no title)"
    texts = []
    for element in title.findall(".//" + q(NS_A, "t")):
        if element.text:
            texts.append(element.text)
    if not texts:
        for element in title.findall(".//" + q(NS_C, "v")):
            if element.text:
                texts.append(element.text)
    title_text = " ".join(texts).strip()
    return title_text or "(no title)"


def replace_child(parent, old_child, new_child):
    children = list(parent)
    if old_child is None:
        parent.append(new_child)
        return
    index = children.index(old_child)
    parent.remove(old_child)
    parent.insert(index, new_child)


def copy_named_child(target_parent, template_parent, name):
    tag = q(NS_C, name)
    existing = target_parent.find(tag)
    source = template_parent.find(tag)
    if source is None:
        if existing is not None:
            target_parent.remove(existing)
        return
    replacement = copy.deepcopy(source)
    if existing is not None:
        replace_child(target_parent, existing, replacement)
        return

    template_children = list(template_parent)
    source_index = template_children.index(source)
    target_children = list(target_parent)
    for following in template_children[source_index + 1 :]:
        following_tag = following.tag
        for index, target_child in enumerate(target_children):
            if target_child.tag == following_tag:
                target_parent.insert(index, replacement)
                return
    for previous in reversed(template_children[:source_index]):
        previous_tag = previous.tag
        for index in range(len(target_children) - 1, -1, -1):
            if target_children[index].tag == previous_tag:
                target_parent.insert(index + 1, replacement)
                return
    target_parent.append(replacement)


def copy_rich_text_style(target_text, template_text):
    if target_text is None or template_text is None:
        return
    template_properties = [
        element
        for element in template_text.iter()
        if local_name(element.tag) in ("rPr", "defRPr", "endParaRPr")
    ]
    if not template_properties:
        return
    target_properties = [
        element
        for element in target_text.iter()
        if local_name(element.tag) in ("rPr", "defRPr", "endParaRPr")
    ]
    for index, target_property in enumerate(target_properties):
        source = template_properties[
            min(index, len(template_properties) - 1)
        ]
        target_property.attrib.clear()
        target_property.attrib.update(source.attrib)
        for child in list(target_property):
            target_property.remove(child)
        for child in list(source):
            target_property.append(copy.deepcopy(child))

    source_run_properties = template_properties[0]
    for run in target_text.findall(".//" + q(NS_A, "r")):
        if run.find(q(NS_A, "rPr")) is not None:
            continue
        new_properties = copy.deepcopy(source_run_properties)
        new_properties.tag = q(NS_A, "rPr")
        run.insert(0, new_properties)


def merge_title(template_title, target_title):
    if target_title is None:
        return None
    if template_title is None:
        return copy.deepcopy(target_title)
    merged = copy.deepcopy(template_title)
    target_text = target_title.find(q(NS_C, "tx"))
    template_text = template_title.find(q(NS_C, "tx"))
    if target_text is not None:
        replacement_text = copy.deepcopy(target_text)
        copy_rich_text_style(replacement_text, template_text)
        replace_child(merged, merged.find(q(NS_C, "tx")), replacement_text)
    return merged


def reorder_series_children(series):
    order = dict(
        (name, index) for index, name in enumerate(SERIES_CHILD_ORDER)
    )
    children = list(series)
    decorated = []
    for original_index, child in enumerate(children):
        decorated.append(
            (
                order.get(local_name(child.tag), len(order)),
                original_index,
                child,
            )
        )
    decorated.sort(key=lambda item: (item[0], item[1]))
    for child in children:
        series.remove(child)
    for unused_priority, unused_index, child in decorated:
        series.append(child)


def apply_series_template(target_series, template_series):
    for name in SERIES_FORMAT_NAMES:
        tag = q(NS_C, name)
        for element in list(target_series.findall(tag)):
            target_series.remove(element)
        for element in template_series.findall(tag):
            target_series.append(copy.deepcopy(element))
    reorder_series_children(target_series)


def apply_chart_type_template(target_type, template_type):
    target_series = list(target_type.findall(q(NS_C, "ser")))
    template_series = list(template_type.findall(q(NS_C, "ser")))
    if not template_series:
        raise ValueError("The selected template chart contains no series.")
    for index, series in enumerate(target_series):
        apply_series_template(
            series, template_series[index % len(template_series)]
        )

    target_axis_ids = [
        copy.deepcopy(element)
        for element in target_type.findall(q(NS_C, "axId"))
    ]
    replacement_children = []
    series_inserted = False
    axes_inserted = False
    for child in list(template_type):
        name = local_name(child.tag)
        if name == "ser":
            if not series_inserted:
                replacement_children.extend(target_series)
                series_inserted = True
            continue
        if name == "axId":
            if not axes_inserted:
                replacement_children.extend(target_axis_ids)
                axes_inserted = True
            continue
        replacement_children.append(copy.deepcopy(child))
    if not series_inserted:
        replacement_children.extend(target_series)
    if not axes_inserted:
        replacement_children.extend(target_axis_ids)

    for child in list(target_type):
        target_type.remove(child)
    for child in replacement_children:
        target_type.append(child)
    return len(target_series)


def axis_position(axis):
    position = axis.find(q(NS_C, "axPos"))
    return position.get("val") if position is not None else None


def merge_axis(template_axis, target_axis):
    merged = copy.deepcopy(template_axis)
    for name in ("axId", "crossAx"):
        target_value = target_axis.find(q(NS_C, name))
        merged_value = merged.find(q(NS_C, name))
        if target_value is not None:
            replacement = copy.deepcopy(target_value)
            replace_child(merged, merged_value, replacement)

    target_title = target_axis.find(q(NS_C, "title"))
    template_title = template_axis.find(q(NS_C, "title"))
    merged_title = merge_title(template_title, target_title)
    current_title = merged.find(q(NS_C, "title"))
    if merged_title is None:
        if current_title is not None:
            merged.remove(current_title)
    elif current_title is None:
        insert_before(
            merged,
            merged_title,
            (
                q(NS_C, "numFmt"),
                q(NS_C, "majorTickMark"),
                q(NS_C, "minorTickMark"),
                q(NS_C, "tickLblPos"),
                q(NS_C, "spPr"),
                q(NS_C, "txPr"),
                q(NS_C, "crossAx"),
                q(NS_C, "extLst"),
            ),
        )
    else:
        replace_child(merged, current_title, merged_title)

    target_extensions = target_axis.find(q(NS_C, "extLst"))
    merged_extensions = merged.find(q(NS_C, "extLst"))
    if target_extensions is not None:
        replace_child(
            merged, merged_extensions, copy.deepcopy(target_extensions)
        )
    elif merged_extensions is not None:
        merged.remove(merged_extensions)
    return merged


def apply_plot_area_template(target_plot, template_plot):
    target_types = chart_type_elements(target_plot)
    template_types = chart_type_elements(template_plot)
    target_names = tuple(local_name(element.tag) for element in target_types)
    template_names = tuple(
        local_name(element.tag) for element in template_types
    )
    if target_names != template_names:
        return 0, False

    series_count = 0
    for target_type, template_type in zip(target_types, template_types):
        series_count += apply_chart_type_template(
            target_type, template_type
        )

    for name in ("layout", "dTable", "spPr"):
        copy_named_child(target_plot, template_plot, name)

    for axis_name in AXIS_NAMES:
        target_axes = list(target_plot.findall(q(NS_C, axis_name)))
        template_axes = list(template_plot.findall(q(NS_C, axis_name)))
        if not template_axes:
            continue
        used_indexes = set()
        for target_axis in target_axes:
            selected_index = None
            target_position = axis_position(target_axis)
            for index, template_axis in enumerate(template_axes):
                if index in used_indexes:
                    continue
                if axis_position(template_axis) == target_position:
                    selected_index = index
                    break
            if selected_index is None:
                for index in range(len(template_axes)):
                    if index not in used_indexes:
                        selected_index = index
                        break
            if selected_index is None:
                selected_index = len(used_indexes) % len(template_axes)
            used_indexes.add(selected_index)
            replacement = merge_axis(
                template_axes[selected_index], target_axis
            )
            replace_child(target_plot, target_axis, replacement)
    return series_count, True


def apply_template_chart_xml(target_xml, template_xml):
    target_root = ET.fromstring(target_xml)
    template_root = ET.fromstring(template_xml)
    target_signature = chart_signature(target_root)
    template_signature = chart_signature(template_root)
    if not target_signature or target_signature != template_signature:
        return target_xml, 0, False, target_signature

    target_chart = target_root.find(q(NS_C, "chart"))
    template_chart = template_root.find(q(NS_C, "chart"))
    target_plot = target_chart.find(q(NS_C, "plotArea"))
    template_plot = template_chart.find(q(NS_C, "plotArea"))
    series_count, compatible = apply_plot_area_template(
        target_plot, template_plot
    )
    if not compatible:
        return target_xml, 0, False, target_signature

    target_title = target_chart.find(q(NS_C, "title"))
    template_title = template_chart.find(q(NS_C, "title"))
    merged_title = merge_title(template_title, target_title)
    if merged_title is None:
        if target_title is not None:
            target_chart.remove(target_title)
    else:
        replace_child(target_chart, target_title, merged_title)

    for name in (
        "autoTitleDeleted",
        "view3D",
        "floor",
        "sideWall",
        "backWall",
        "legend",
        "plotVisOnly",
        "dispBlanksAs",
        "showDLblsOverMax",
        "spPr",
        "txPr",
    ):
        copy_named_child(target_chart, template_chart, name)

    for name in ("roundedCorners", "style", "clrMapOvr", "spPr", "txPr"):
        copy_named_child(target_root, template_root, name)

    xml_body = ET.tostring(target_root, encoding="utf-8")
    declaration = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    return declaration + xml_body, series_count, True, target_signature


def integer_child(parent, local_name):
    element = parent.find(q(NS_XDR, local_name))
    if element is None or element.text is None:
        return None, element
    try:
        return int(element.text), element
    except ValueError:
        return None, element


def format_drawing_xml(xml_data, width_columns, height_rows):
    root = ET.fromstring(xml_data)
    resized = 0
    for anchor in root.findall(q(NS_XDR, "twoCellAnchor")):
        if anchor.find(".//" + q(NS_C, "chart")) is None:
            continue
        from_element = anchor.find(q(NS_XDR, "from"))
        to_element = anchor.find(q(NS_XDR, "to"))
        if from_element is None or to_element is None:
            continue
        start_column, unused_column_element = integer_child(from_element, "col")
        start_row, unused_row_element = integer_child(from_element, "row")
        unused_end_column, end_column_element = integer_child(to_element, "col")
        unused_end_row, end_row_element = integer_child(to_element, "row")
        if None in (start_column, start_row, end_column_element, end_row_element):
            continue
        end_column_element.text = str(start_column + width_columns)
        end_row_element.text = str(start_row + height_rows)
        for offset_name in ("colOff", "rowOff"):
            offset = to_element.find(q(NS_XDR, offset_name))
            if offset is not None:
                offset.text = "0"
        resized += 1

    for anchor_name in ("oneCellAnchor", "absoluteAnchor"):
        for anchor in root.findall(q(NS_XDR, anchor_name)):
            if anchor.find(".//" + q(NS_C, "chart")) is None:
                continue
            extent = anchor.find(q(NS_XDR, "ext"))
            if extent is None:
                continue
            extent.set("cx", str(width_columns * 64 * 9525))
            extent.set("cy", str(height_rows * 20 * 9525))
            resized += 1

    if not resized:
        return xml_data, 0
    xml_body = ET.tostring(root, encoding="utf-8")
    declaration = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    return declaration + xml_body, resized


def is_chart_xml(member_name):
    normalized = member_name.replace("\\", "/").lower()
    basename = normalized.rsplit("/", 1)[-1]
    return "/charts/" in normalized and basename.startswith("chart") and basename.endswith(".xml")


def is_drawing_xml(member_name):
    normalized = member_name.replace("\\", "/").lower()
    basename = normalized.rsplit("/", 1)[-1]
    return "/drawings/" in normalized and basename.startswith("drawing") and basename.endswith(".xml")


def natural_member_key(member_name):
    return [
        int(part) if part.isdigit() else part.lower()
        for part in re.split(r"([0-9]+)", member_name)
    ]


def relationship_target(base_part, target):
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(
        posixpath.join(posixpath.dirname(base_part), target)
    )


def chart_anchor_size(workbook_zip, chart_member):
    relationship_members = [
        name
        for name in workbook_zip.namelist()
        if "/drawings/_rels/" in name.replace("\\", "/").lower()
        and name.lower().endswith(".xml.rels")
    ]
    for relationship_member in relationship_members:
        normalized = relationship_member.replace("\\", "/")
        drawing_member = normalized.replace("/_rels/", "/")[:-5]
        if drawing_member not in workbook_zip.namelist():
            continue
        relationships = ET.fromstring(
            workbook_zip.read(relationship_member)
        )
        relationship_id = None
        for relationship in relationships.findall(q(NS_REL, "Relationship")):
            relationship_type = relationship.get("Type", "")
            if not relationship_type.endswith("/chart"):
                continue
            resolved = relationship_target(
                drawing_member, relationship.get("Target", "")
            )
            if resolved == chart_member:
                relationship_id = relationship.get("Id")
                break
        if not relationship_id:
            continue

        drawing = ET.fromstring(workbook_zip.read(drawing_member))
        for anchor in list(drawing):
            chart_reference = anchor.find(".//" + q(NS_C, "chart"))
            if chart_reference is None:
                continue
            if chart_reference.get(q(NS_R, "id")) != relationship_id:
                continue
            anchor_name = local_name(anchor.tag)
            if anchor_name == "twoCellAnchor":
                from_element = anchor.find(q(NS_XDR, "from"))
                to_element = anchor.find(q(NS_XDR, "to"))
                if from_element is None or to_element is None:
                    return None, None
                start_column, unused = integer_child(from_element, "col")
                start_row, unused = integer_child(from_element, "row")
                end_column, unused = integer_child(to_element, "col")
                end_row, unused = integer_child(to_element, "row")
                if None in (
                    start_column,
                    start_row,
                    end_column,
                    end_row,
                ):
                    return None, None
                return (
                    max(1, end_column - start_column),
                    max(1, end_row - start_row),
                )
            extent = anchor.find(q(NS_XDR, "ext"))
            if extent is None:
                return None, None
            try:
                width_columns = int(
                    round(float(extent.get("cx")) / (64.0 * 9525.0))
                )
                height_rows = int(
                    round(float(extent.get("cy")) / (20.0 * 9525.0))
                )
                return max(1, width_columns), max(1, height_rows)
            except (TypeError, ValueError):
                return None, None
    return None, None


def template_chart_catalog(workbook_path):
    workbook_zip = zipfile.ZipFile(workbook_path, "r")
    try:
        bad_member = workbook_zip.testzip()
        if bad_member:
            raise ValueError(
                "Template workbook CRC validation failed for {0}".format(
                    bad_member
                )
            )
        chart_members = sorted(
            (
                name
                for name in workbook_zip.namelist()
                if is_chart_xml(name)
            ),
            key=natural_member_key,
        )
        catalog = []
        for index, member_name in enumerate(chart_members, 1):
            xml_data = workbook_zip.read(member_name)
            root = ET.fromstring(xml_data)
            width_columns, height_rows = chart_anchor_size(
                workbook_zip, member_name
            )
            catalog.append(
                {
                    "index": index,
                    "member_name": member_name,
                    "title": chart_title_text(root),
                    "signature": chart_signature(root),
                    "width_columns": width_columns,
                    "height_rows": height_rows,
                    "xml_data": xml_data,
                }
            )
        return catalog
    finally:
        workbook_zip.close()


def select_template_chart(workbook_path, chart_number):
    catalog = template_chart_catalog(workbook_path)
    if not catalog:
        raise ValueError(
            "The template workbook contains no native Excel charts: {0}".format(
                workbook_path
            )
        )
    if chart_number > len(catalog):
        raise ValueError(
            "Template chart {0} does not exist; the workbook contains {1} "
            "chart(s).".format(chart_number, len(catalog))
        )
    return catalog[chart_number - 1]


def format_chart_signature(signature):
    return "+".join(signature) if signature else "unknown"


def print_template_charts(workbook_path):
    catalog = template_chart_catalog(workbook_path)
    if not catalog:
        print("The template workbook contains no native Excel charts.")
        return 2
    print("Template workbook: {0}".format(workbook_path))
    print("Native Excel charts: {0}".format(len(catalog)))
    for item in catalog:
        size_text = "unknown size"
        if item["width_columns"] and item["height_rows"]:
            size_text = "approximately {0} columns x {1} rows".format(
                item["width_columns"], item["height_rows"]
            )
        print(
            "  {0}: {1} | {2} | {3}".format(
                item["index"],
                item["title"],
                format_chart_signature(item["signature"]),
                size_text,
            )
        )
    return 0


def ensure_directory(path):
    if not os.path.isdir(path):
        os.makedirs(path)


def replace_file(source_path, destination_path):
    replace_function = getattr(os, "replace", None)
    if replace_function is not None:
        replace_function(source_path, destination_path)
        return
    if os.path.exists(destination_path):
        os.remove(destination_path)
    os.rename(source_path, destination_path)


def unique_backup_path(source_path):
    stem, extension = os.path.splitext(source_path)
    candidate = stem + "_before_chart_format" + extension
    if not os.path.exists(candidate):
        return candidate
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    return stem + "_before_chart_format_" + timestamp + extension


def validate_workbook(path):
    workbook_zip = zipfile.ZipFile(path, "r")
    try:
        bad_member = workbook_zip.testzip()
        if bad_member:
            raise ValueError("CRC validation failed for {0}".format(bad_member))
        chart_members = [
            name for name in workbook_zip.namelist() if is_chart_xml(name)
        ]
        for member_name in chart_members:
            ET.fromstring(workbook_zip.read(member_name))
        return len(chart_members)
    finally:
        workbook_zip.close()


def format_workbook(
    source_path,
    output_path,
    font_name,
    resize,
    width_columns,
    height_rows,
    overwrite,
    in_place,
    create_backup,
    template_chart=None,
    target_chart_title_patterns=None,
):
    output_folder = os.path.dirname(output_path)
    ensure_directory(output_folder)
    if not in_place and os.path.exists(output_path) and not overwrite:
        raise ValueError(
            "Output already exists; use --overwrite: {0}".format(output_path)
        )

    descriptor, temporary_path = tempfile.mkstemp(
        prefix="chart_format_", suffix=".xlsx", dir=output_folder
    )
    os.close(descriptor)
    chart_parts = 0
    series_count = 0
    resized_charts = 0
    templated_charts = 0
    incompatible_charts = 0
    unselected_charts = 0
    resize_suppressed = False
    try:
        source_zip = zipfile.ZipFile(source_path, "r")
        effective_resize = resize
        chart_selection = {}
        if template_chart is not None:
            for source_member in source_zip.namelist():
                if not is_chart_xml(source_member):
                    continue
                source_root = ET.fromstring(source_zip.read(source_member))
                source_title = chart_title_text(source_root)
                selected = not target_chart_title_patterns or matches_patterns(
                    source_title, target_chart_title_patterns
                )
                source_signature = chart_signature(source_root)
                chart_selection[source_member] = (
                    selected,
                    source_signature,
                    source_title,
                )
            if resize and any(
                not selected
                or signature != template_chart["signature"]
                for selected, signature, unused_title in chart_selection.values()
            ):
                effective_resize = False
                resize_suppressed = True
        target_zip = zipfile.ZipFile(
            temporary_path,
            "w",
            compression=zipfile.ZIP_DEFLATED,
            allowZip64=True,
        )
        try:
            for member in source_zip.infolist():
                data = source_zip.read(member.filename)
                if is_chart_xml(member.filename):
                    chart_parts += 1
                    if template_chart is None:
                        data, formatted_series = format_chart_xml(
                            data, font_name
                        )
                        series_count += formatted_series
                        templated_charts += 1
                    else:
                        selected, unused_signature, unused_title = chart_selection[
                            member.filename
                        ]
                        if not selected:
                            unselected_charts += 1
                            target_zip.writestr(member, data)
                            continue
                        (
                            data,
                            formatted_series,
                            applied,
                            unused_signature,
                        ) = apply_template_chart_xml(
                            data, template_chart["xml_data"]
                        )
                        if applied:
                            series_count += formatted_series
                            templated_charts += 1
                        else:
                            incompatible_charts += 1
                elif effective_resize and is_drawing_xml(member.filename):
                    data, resized = format_drawing_xml(
                        data, width_columns, height_rows
                    )
                    resized_charts += resized
                target_zip.writestr(member, data)
        finally:
            target_zip.close()
            source_zip.close()

        if chart_parts == 0:
            raise ValueError("The workbook contains no native Excel charts.")
        if template_chart is not None and templated_charts == 0:
            raise ValueError(
                "No chart has the selected template type ({0}).".format(
                    format_chart_signature(template_chart["signature"])
                )
            )
        validate_workbook(temporary_path)

        backup_path = None
        if in_place and create_backup:
            backup_path = unique_backup_path(source_path)
            shutil.copy2(source_path, backup_path)
        replace_file(temporary_path, output_path)
        temporary_path = None
        return {
            "chart_parts": chart_parts,
            "series_count": series_count,
            "resized_charts": resized_charts,
            "backup_path": backup_path,
            "templated_charts": templated_charts,
            "incompatible_charts": incompatible_charts,
            "unselected_charts": unselected_charts,
            "resize_suppressed": resize_suppressed,
        }
    finally:
        if temporary_path and os.path.exists(temporary_path):
            os.remove(temporary_path)


def write_log(log_path, args, jobs, successes, failures, interrupted=False):
    with open(log_path, "w") as log_file:
        log_file.write(
            "format_all_odb_excel_charts.py version {0}\n".format(
                SCRIPT_VERSION
            )
        )
        log_file.write("Time: {0}\n".format(datetime.datetime.now()))
        log_file.write("Arguments: {0}\n".format(repr(sys.argv)))
        log_file.write("Jobs: {0}\n".format(len(jobs)))
        log_file.write("Successful: {0}\n".format(len(successes)))
        log_file.write("Failed: {0}\n\n".format(len(failures)))
        log_file.write("Interrupted by user: {0}\n\n".format(interrupted))
        for source_path, output_path, result in successes:
            log_file.write("SOURCE: {0}\n".format(source_path))
            log_file.write("OUTPUT: {0}\n".format(output_path))
            log_file.write(
                "CHARTS: {0}; FORMATTED: {1}; INCOMPATIBLE: {2}; "
                "TITLE-FILTERED: {3}; SERIES: {4}; RESIZED: {5}\n".format(
                    result["chart_parts"],
                    result["templated_charts"],
                    result["incompatible_charts"],
                    result["unselected_charts"],
                    result["series_count"],
                    result["resized_charts"],
                )
            )
            if result["resize_suppressed"]:
                log_file.write(
                    "RESIZE SUPPRESSED: workbook contains charts that were "
                    "not selected or were incompatible.\n"
                )
            if result["backup_path"]:
                log_file.write(
                    "BACKUP: {0}\n".format(result["backup_path"])
                )
            log_file.write("\n")
        for source_path, output_path, error_text, traceback_text in failures:
            log_file.write("FAILED SOURCE: {0}\n".format(source_path))
            log_file.write("INTENDED OUTPUT: {0}\n".format(output_path))
            log_file.write("ERROR: {0}\n".format(error_text))
            log_file.write(traceback_text + "\n")


def option_was_supplied(option_name):
    return any(
        argument == option_name or argument.startswith(option_name + "=")
        for argument in sys.argv[1:]
    )


def main():
    args = parse_arguments()
    template_workbook = None
    template_chart = None
    if args.template_workbook:
        template_workbook = os.path.abspath(args.template_workbook)
        if not os.path.isfile(template_workbook):
            print(
                "Template workbook does not exist: {0}".format(
                    template_workbook
                )
            )
            return 2
        if args.list_template_charts:
            try:
                return print_template_charts(template_workbook)
            except Exception as exc:
                print("Could not read template workbook: {0}".format(exc))
                return 2
        try:
            template_chart = select_template_chart(
                template_workbook, args.template_chart
            )
        except Exception as exc:
            print("Could not select template chart: {0}".format(exc))
            return 2

    input_dir = os.path.abspath(args.input_dir)
    if not os.path.isdir(input_dir):
        print("Input directory does not exist: {0}".format(input_dir))
        return 2
    patterns = args.pattern or list(DEFAULT_PATTERNS)
    output_dir = (
        os.path.abspath(args.output_dir) if args.output_dir else None
    )
    workbooks = find_workbooks(
        input_dir,
        patterns,
        args.recursive,
        args.suffix,
        excluded_paths=(template_workbook,) if template_workbook else (),
    )
    if not workbooks:
        print(
            "No matching .xlsx files were found in {0}. Patterns: {1}".format(
                input_dir, ", ".join(patterns)
            )
        )
        return 2

    jobs = [
        (
            source_path,
            output_path_for(
                source_path,
                input_dir,
                output_dir,
                args.suffix,
                args.in_place,
            ),
        )
        for source_path in workbooks
    ]
    print(
        "format_all_odb_excel_charts.py version {0}".format(
            SCRIPT_VERSION
        )
    )
    if template_chart is not None:
        print("Template workbook: {0}".format(template_workbook))
        print(
            "Template chart {0}: {1} ({2})".format(
                template_chart["index"],
                template_chart["title"],
                format_chart_signature(template_chart["signature"]),
            )
        )
    print("Matched workbooks: {0}".format(len(jobs)))
    if args.dry_run:
        for source_path, output_path in jobs:
            print("  {0}".format(source_path))
            print("    -> {0}".format(output_path))
        return 0

    successes = []
    failures = []
    interrupted = False
    width_columns = args.chart_width_columns
    height_rows = args.chart_height_rows
    if template_chart is not None and not args.no_resize:
        if (
            template_chart["width_columns"]
            and not option_was_supplied("--chart-width-columns")
        ):
            width_columns = template_chart["width_columns"]
        if (
            template_chart["height_rows"]
            and not option_was_supplied("--chart-height-rows")
        ):
            height_rows = template_chart["height_rows"]
    for source_path, output_path in jobs:
        print("Formatting: {0}".format(source_path))
        try:
            result = format_workbook(
                source_path=source_path,
                output_path=output_path,
                font_name=args.font,
                resize=not args.no_resize,
                width_columns=width_columns,
                height_rows=height_rows,
                overwrite=args.overwrite,
                in_place=args.in_place,
                create_backup=not args.no_backup,
                template_chart=template_chart,
                target_chart_title_patterns=args.target_chart_title,
            )
            successes.append((source_path, output_path, result))
            print(
                "  Wrote: {0} ({1} chart(s), {2} series)".format(
                    output_path,
                    result["chart_parts"],
                    result["series_count"],
                )
            )
            if result["incompatible_charts"]:
                print(
                    "  Left unchanged: {0} chart(s) with a different "
                    "chart type.".format(result["incompatible_charts"])
                )
            if result["unselected_charts"]:
                print(
                    "  Left unchanged: {0} chart(s) excluded by the "
                    "title filter.".format(result["unselected_charts"])
                )
            if result["resize_suppressed"]:
                print(
                    "  Chart resizing was skipped because this workbook "
                    "also contains charts that were not selected or were "
                    "incompatible."
                )
            if result["backup_path"]:
                print("  Backup: {0}".format(result["backup_path"]))
        except KeyboardInterrupt:
            interrupted = True
            print("\nInterrupted by user. The current temporary file was removed.")
            break
        except Exception as exc:
            traceback_text = traceback.format_exc()
            failures.append(
                (source_path, output_path, str(exc), traceback_text)
            )
            print("  FAILED: {0}".format(exc))

    log_folder = output_dir or input_dir
    ensure_directory(log_folder)
    log_path = os.path.join(log_folder, "format_all_odb_excel_charts.log")
    write_log(
        log_path,
        args,
        jobs,
        successes,
        failures,
        interrupted=interrupted,
    )
    print(
        "Completed: {0} succeeded, {1} failed.".format(
            len(successes), len(failures)
        )
    )
    print("Log: {0}".format(log_path))
    if interrupted:
        return 130
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
