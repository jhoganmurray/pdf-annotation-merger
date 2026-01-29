#!/usr/bin/env python3
"""
XFDF Importer using pikepdf

Imports XFDF annotation data directly into a PDF file without requiring
Adobe Acrobat. Uses pikepdf (based on qpdf) for robust PDF structure handling.

Requirements:
    pip install pikepdf

Usage:
    from xfdf_importer import import_xfdf_to_pdf
    import_xfdf_to_pdf("base.pdf", "comments.xfdf", "output.pdf")
"""

import xml.etree.ElementTree as ET
from typing import List, Dict, Any, Optional, Tuple
from datetime import datetime
import pikepdf
from pikepdf import Array, Dictionary, Name, String


def parse_xfdf(xfdf_path: str) -> Tuple[str, List[Dict[str, Any]]]:
    """
    Parse an XFDF file and extract annotation data.

    Returns:
        Tuple of (source_pdf_name, list of annotation dicts)
    """
    tree = ET.parse(xfdf_path)
    root = tree.getroot()

    # Handle namespace
    ns = {'xfdf': 'http://ns.adobe.com/xfdf/'}

    # Get source PDF name
    f_elem = root.find('xfdf:f', ns) or root.find('f')
    source_pdf = f_elem.get('href', '') if f_elem is not None else ''

    # Find annots element
    annots_elem = root.find('xfdf:annots', ns) or root.find('annots')
    if annots_elem is None:
        return source_pdf, []

    annotations = []
    for elem in annots_elem:
        # Get tag name without namespace
        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag

        annot = {
            'type': tag,
            'page': int(elem.get('page', '0')),
            'rect': parse_rect(elem.get('rect', '0,0,0,0')),
            'name': elem.get('name', ''),
            'color': parse_color(elem.get('color')),
            'title': elem.get('title', ''),
            'subject': elem.get('subject', ''),
            'opacity': float(elem.get('opacity', '1.0')),
            'width': float(elem.get('width', '1.0')),
            'flags': int(elem.get('flags', '4')),  # Default: Print flag
        }

        # Contents element
        contents_elem = elem.find('xfdf:contents', ns) or elem.find('contents')
        if contents_elem is not None and contents_elem.text:
            annot['contents'] = contents_elem.text

        # Type-specific data
        if tag == 'ink':
            annot['inklist'] = parse_inklist(elem, ns)
        elif tag == 'line':
            annot['start'] = parse_point(elem.get('start', '0,0'))
            annot['end'] = parse_point(elem.get('end', '0,0'))
        elif tag in ('highlight', 'underline', 'strikeout', 'squiggly'):
            coords = elem.get('coords')
            if coords:
                annot['quadpoints'] = parse_coords(coords)
        elif tag in ('polygon', 'polyline'):
            vertices = elem.get('vertices')
            if vertices:
                annot['vertices'] = parse_coords(vertices)

        annotations.append(annot)

    return source_pdf, annotations


def parse_rect(rect_str: str) -> List[float]:
    """Parse rect string "x1,y1,x2,y2" into list of floats."""
    return [float(x) for x in rect_str.split(',')]


def parse_point(point_str: str) -> List[float]:
    """Parse point string "x,y" into list of floats."""
    return [float(x) for x in point_str.split(',')]


def parse_color(color_str: Optional[str]) -> List[float]:
    """Parse hex color "#RRGGBB" into RGB array (0-1 range)."""
    if not color_str or not color_str.startswith('#'):
        return [1.0, 1.0, 0.0]  # Default yellow

    hex_color = color_str[1:]
    if len(hex_color) == 6:
        r = int(hex_color[0:2], 16) / 255.0
        g = int(hex_color[2:4], 16) / 255.0
        b = int(hex_color[4:6], 16) / 255.0
        return [r, g, b]
    return [1.0, 1.0, 0.0]


def parse_coords(coords_str: str) -> List[float]:
    """Parse coords string "x1,y1,x2,y2,..." into list of floats."""
    return [float(x) for x in coords_str.split(',')]


def parse_inklist(elem, ns: dict) -> List[List[Tuple[float, float]]]:
    """Parse ink annotation gesture data."""
    inklist_elem = elem.find('xfdf:inklist', ns) or elem.find('inklist')
    if inklist_elem is None:
        return []

    strokes = []
    for gesture in inklist_elem.findall('xfdf:gesture', ns) or inklist_elem.findall('gesture'):
        if gesture.text:
            points = []
            for point_str in gesture.text.split(';'):
                coords = point_str.split(',')
                if len(coords) >= 2:
                    points.append((float(coords[0]), float(coords[1])))
            if points:
                strokes.append(points)
    return strokes


def create_annotation_dict(annot: Dict[str, Any]) -> Dictionary:
    """
    Create a PDF annotation dictionary from parsed XFDF data.
    """
    annot_type = annot['type'].lower()

    # Map XFDF types to PDF annotation subtypes
    type_map = {
        'text': Name.Text,
        'freetext': Name.FreeText,
        'highlight': Name.Highlight,
        'underline': Name.Underline,
        'strikeout': Name.StrikeOut,
        'squiggly': Name.Squiggly,
        'ink': Name.Ink,
        'line': Name.Line,
        'square': Name.Square,
        'circle': Name.Circle,
        'polygon': Name.Polygon,
        'polyline': Name.PolyLine,
        'stamp': Name.Stamp,
        'caret': Name.Caret,
        'fileattachment': Name.FileAttachment,
    }

    subtype = type_map.get(annot_type, Name.Text)

    # Build base annotation dictionary
    annot_dict = Dictionary({
        Name.Type: Name.Annot,
        Name.Subtype: subtype,
        Name.Rect: Array([float(x) for x in annot['rect']]),
        Name.F: annot.get('flags', 4),  # Flags (4 = Print)
    })

    # Color
    if annot.get('color'):
        annot_dict[Name.C] = Array(annot['color'])

    # Name/identifier
    if annot.get('name'):
        annot_dict[Name.NM] = String(annot['name'])

    # Title (author)
    if annot.get('title'):
        annot_dict[Name.T] = String(annot['title'])

    # Subject
    if annot.get('subject'):
        annot_dict[Name.Subj] = String(annot['subject'])

    # Contents
    if annot.get('contents'):
        annot_dict[Name.Contents] = String(annot['contents'])

    # Opacity (CA = stroke opacity)
    if annot.get('opacity', 1.0) < 1.0:
        annot_dict[Name.CA] = annot['opacity']

    # Creation date
    annot_dict[Name.CreationDate] = String(format_pdf_date(datetime.now()))
    annot_dict[Name.M] = String(format_pdf_date(datetime.now()))

    # Type-specific handling
    if annot_type == 'ink':
        ink_list = build_ink_list(annot.get('inklist', []))
        if ink_list:
            annot_dict[Name.InkList] = ink_list
        # Border style for ink
        annot_dict[Name.BS] = Dictionary({
            Name.W: annot.get('width', 1.0),
            Name.S: Name.S,  # Solid
        })

    elif annot_type == 'line':
        if annot.get('start') and annot.get('end'):
            annot_dict[Name.L] = Array(annot['start'] + annot['end'])

    elif annot_type in ('highlight', 'underline', 'strikeout', 'squiggly'):
        if annot.get('quadpoints'):
            annot_dict[Name.QuadPoints] = Array(annot['quadpoints'])

    elif annot_type in ('polygon', 'polyline'):
        if annot.get('vertices'):
            annot_dict[Name.Vertices] = Array(annot['vertices'])

    elif annot_type == 'square' or annot_type == 'circle':
        # Border style
        annot_dict[Name.BS] = Dictionary({
            Name.W: annot.get('width', 1.0),
            Name.S: Name.S,
        })

    elif annot_type == 'text':
        # Note annotation (sticky note)
        annot_dict[Name.Open] = False
        annot_dict[Name.Name] = Name.Comment

    return annot_dict


def build_ink_list(strokes: List[List[Tuple[float, float]]]) -> Array:
    """Build InkList array from stroke data."""
    ink_list = Array()
    for stroke in strokes:
        stroke_array = Array()
        for x, y in stroke:
            stroke_array.append(float(x))
            stroke_array.append(float(y))
        ink_list.append(stroke_array)
    return ink_list


def format_pdf_date(dt: datetime) -> str:
    """Format datetime as PDF date string."""
    return dt.strftime("D:%Y%m%d%H%M%S")


def import_xfdf_to_pdf(
    pdf_path: str,
    xfdf_path: str,
    output_path: str,
    linearize: bool = False
) -> int:
    """
    Import XFDF annotations into a PDF file.

    Args:
        pdf_path: Path to the source PDF file
        xfdf_path: Path to the XFDF file with annotations
        output_path: Path to save the modified PDF
        linearize: Whether to linearize (optimize for web) the output

    Returns:
        Number of annotations imported
    """
    # Parse XFDF
    _, annotations = parse_xfdf(xfdf_path)

    if not annotations:
        # No annotations to import - just copy the file
        with open(pdf_path, 'rb') as src, open(output_path, 'wb') as dst:
            dst.write(src.read())
        return 0

    # Open PDF
    pdf = pikepdf.open(pdf_path, allow_overwriting_input=True)

    # Group annotations by page
    by_page: Dict[int, List[Dict]] = {}
    for annot in annotations:
        page_num = annot['page']
        by_page.setdefault(page_num, []).append(annot)

    imported_count = 0

    # Add annotations to each page
    for page_num, page_annots in by_page.items():
        if page_num >= len(pdf.pages):
            print(f"Warning: Page {page_num} does not exist, skipping annotations")
            continue

        page = pdf.pages[page_num]

        # Get existing annots array, handling indirect references
        annots_array = []
        if Name.Annots in page:
            existing = page[Name.Annots]
            try:
                for item in existing:
                    annots_array.append(item)
            except TypeError:
                pass

        # Add each annotation
        for annot_data in page_annots:
            annot_dict = create_annotation_dict(annot_data)

            # Create indirect object for annotation
            annot_obj = pdf.make_indirect(annot_dict)

            # Note: P (page reference) is optional, skipping to avoid compatibility issues

            annots_array.append(annot_obj)
            imported_count += 1

        # Update page's Annots array
        page[Name.Annots] = Array(annots_array)

    # Save the modified PDF
    save_options = {}
    if linearize:
        save_options['linearize'] = True

    pdf.save(output_path, **save_options)
    pdf.close()

    return imported_count


def import_annotations_direct(
    pdf_path: str,
    annotations: List[Dict[str, Any]],
    output_path: str
) -> int:
    """
    Import annotations directly (without XFDF file) into a PDF.

    This allows the Comment Collector to bypass XFDF generation
    and write annotations directly.

    Args:
        pdf_path: Path to the source PDF file
        annotations: List of annotation dictionaries (same format as extract_annotations)
        output_path: Path to save the modified PDF

    Returns:
        Number of annotations imported
    """
    if not annotations:
        with open(pdf_path, 'rb') as src, open(output_path, 'wb') as dst:
            dst.write(src.read())
        return 0

    pdf = pikepdf.open(pdf_path, allow_overwriting_input=True)

    # Group by page
    by_page: Dict[int, List[Dict]] = {}
    for annot in annotations:
        page_num = annot['page']
        by_page.setdefault(page_num, []).append(annot)

    imported_count = 0

    for page_num, page_annots in by_page.items():
        if page_num >= len(pdf.pages):
            continue

        page = pdf.pages[page_num]

        # Get page height from MediaBox
        mediabox = page.get(Name.MediaBox)
        if mediabox is not None:
            page_height = float(mediabox[3]) - float(mediabox[1])
        else:
            page_height = 792.0  # Default letter size

        # Get existing annots array, handling indirect references
        annots_array = []
        if Name.Annots in page:
            existing = page[Name.Annots]
            # Iterate through existing annotations
            try:
                for item in existing:
                    annots_array.append(item)
            except TypeError:
                # If not iterable, it might be a single annotation or invalid
                pass

        for annot_data in page_annots:
            # Convert from PyMuPDF format to PDF format
            pdf_annot = convert_pymupdf_to_pdf_annot(annot_data, page_height)

            # Create indirect object for annotation
            annot_obj = pdf.make_indirect(pdf_annot)

            # Note: P (page reference) is optional and can cause issues, skipping it
            # Most PDF readers derive the page from the Annots array location

            annots_array.append(annot_obj)
            imported_count += 1

        # Set the new Annots array
        page[Name.Annots] = Array(annots_array)

    pdf.save(output_path)
    pdf.close()

    return imported_count


def convert_pymupdf_to_pdf_annot(annot: Dict[str, Any], page_height: float) -> Dictionary:
    """
    Convert a PyMuPDF-extracted annotation dict to a PDF annotation dictionary.

    PyMuPDF uses screen coordinates (origin top-left), PDF uses
    page coordinates (origin bottom-left).
    """
    annot_type = annot['type']

    type_map = {
        'Text': Name.Text,
        'FreeText': Name.FreeText,
        'Highlight': Name.Highlight,
        'Underline': Name.Underline,
        'StrikeOut': Name.StrikeOut,
        'Squiggly': Name.Squiggly,
        'Ink': Name.Ink,
        'Line': Name.Line,
        'Square': Name.Square,
        'Circle': Name.Circle,
        'Polygon': Name.Polygon,
        'PolyLine': Name.PolyLine,
        'Stamp': Name.Stamp,
        'Caret': Name.Caret,
        'FileAttachment': Name.FileAttachment,
    }

    subtype = type_map.get(annot_type, Name.Text)

    # Convert rect from screen coords to PDF coords
    rect = annot['rect']
    pdf_rect = [
        rect[0],                    # x0 unchanged
        page_height - rect[3],      # y0 = page_height - screen_y1
        rect[2],                    # x1 unchanged
        page_height - rect[1],      # y1 = page_height - screen_y0
    ]

    annot_dict = Dictionary({
        Name.Type: Name.Annot,
        Name.Subtype: subtype,
        Name.Rect: Array(pdf_rect),
        Name.F: annot.get('flags', 4),
    })

    # Color
    if annot.get('stroke_color'):
        annot_dict[Name.C] = Array(annot['stroke_color'])
    elif annot.get('fill_color'):
        annot_dict[Name.C] = Array(annot['fill_color'])

    # Author
    if annot.get('author'):
        annot_dict[Name.T] = String(annot['author'])

    # Subject
    if annot.get('subject'):
        annot_dict[Name.Subj] = String(annot['subject'])

    # Contents
    if annot.get('content'):
        annot_dict[Name.Contents] = String(annot['content'])

    # Opacity
    if annot.get('opacity', 1.0) < 1.0:
        annot_dict[Name.CA] = annot['opacity']

    # Dates
    annot_dict[Name.CreationDate] = String(format_pdf_date(datetime.now()))
    annot_dict[Name.M] = String(format_pdf_date(datetime.now()))

    # Type-specific handling
    if annot_type == 'Ink':
        if annot.get('vertices'):
            ink_list = Array()
            for stroke in annot['vertices']:
                stroke_array = Array()
                for pt in stroke:
                    if isinstance(pt, (list, tuple)) and len(pt) >= 2:
                        stroke_array.append(float(pt[0]))
                        stroke_array.append(float(page_height - pt[1]))
                ink_list.append(stroke_array)
            annot_dict[Name.InkList] = ink_list

        width = annot.get('border_width', 1.0)
        annot_dict[Name.BS] = Dictionary({
            Name.W: width,
            Name.S: Name.S,
        })

    elif annot_type == 'Line':
        if annot.get('vertices') and len(annot['vertices']) >= 2:
            v = annot['vertices']
            start = v[0] if isinstance(v[0], (list, tuple)) else v[:2]
            end = v[1] if isinstance(v[1], (list, tuple)) else v[2:4]
            annot_dict[Name.L] = Array([
                start[0], page_height - start[1],
                end[0], page_height - end[1]
            ])

    elif annot_type in ('Highlight', 'Underline', 'StrikeOut', 'Squiggly'):
        if annot.get('vertices'):
            quad_points = []
            for v in annot['vertices']:
                if isinstance(v, (list, tuple)) and len(v) >= 2:
                    quad_points.append(float(v[0]))
                    quad_points.append(float(page_height - v[1]))
            if quad_points:
                annot_dict[Name.QuadPoints] = Array(quad_points)

    elif annot_type in ('Polygon', 'PolyLine'):
        if annot.get('vertices'):
            vertices = []
            for v in annot['vertices']:
                if isinstance(v, (list, tuple)) and len(v) >= 2:
                    vertices.append(float(v[0]))
                    vertices.append(float(page_height - v[1]))
            if vertices:
                annot_dict[Name.Vertices] = Array(vertices)

    elif annot_type == 'Text':
        annot_dict[Name.Open] = False
        annot_dict[Name.Name] = Name.Comment

    return annot_dict


# Command-line interface
if __name__ == "__main__":
    import sys

    if len(sys.argv) < 4:
        print("Usage: python xfdf_importer.py <input.pdf> <annotations.xfdf> <output.pdf>")
        print("\nImports XFDF annotations into a PDF file using pikepdf.")
        sys.exit(1)

    pdf_path = sys.argv[1]
    xfdf_path = sys.argv[2]
    output_path = sys.argv[3]

    try:
        count = import_xfdf_to_pdf(pdf_path, xfdf_path, output_path)
        print(f"Successfully imported {count} annotation(s) to {output_path}")
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)
