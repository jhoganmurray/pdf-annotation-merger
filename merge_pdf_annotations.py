#!/usr/bin/env python3
"""
PDF Annotation Merger

Merges annotations from multiple versions of the same PDF into a single file.
Designed to solve OneDrive sync conflicts where multiple users annotate the same
document and create divergent copies.

Usage:
    python merge_pdf_annotations.py output.pdf input1.pdf input2.pdf [input3.pdf ...]

The first input file is used as the base, and unique annotations from all other
files are merged into it.
"""

import fitz  # PyMuPDF
import sys
import os
import io
import tempfile
import shutil
from typing import List, Dict, Tuple, Set
from dataclasses import dataclass
from copy import deepcopy


@dataclass
class AnnotationKey:
    """Hashable key for identifying unique annotations"""
    page: int
    annot_type: str
    rect: Tuple[float, float, float, float]
    content: str

    def __hash__(self):
        return hash((self.page, self.annot_type, self.rect, self.content))

    def __eq__(self, other):
        return (self.page == other.page and
                self.annot_type == other.annot_type and
                self.rect == other.rect and
                self.content == other.content)


def round_rect(rect, precision=1) -> Tuple[float, float, float, float]:
    """Round rectangle coordinates for comparison (handles minor float differences)"""
    return tuple(round(x, precision) for x in rect)


def repair_pdf(input_path: str, temp_dir: str) -> str:
    """
    Repair a PDF by opening and re-saving it with cleanup options.
    This fixes xref table issues common in OneDrive-edited PDFs.
    """
    base_name = os.path.basename(input_path)
    temp_path = os.path.join(temp_dir, f"repaired_{base_name}")

    old_stderr = sys.stderr
    sys.stderr = io.StringIO()
    try:
        doc = fitz.open(input_path)
        doc.save(temp_path, garbage=4, deflate=True, clean=True)
        doc.close()
    finally:
        sys.stderr = old_stderr

    return temp_path


def open_pdf_safe(filepath: str):
    """Open a PDF file, suppressing recoverable MuPDF warnings."""
    old_stderr = sys.stderr
    sys.stderr = io.StringIO()
    try:
        doc = fitz.open(filepath)
    finally:
        sys.stderr = old_stderr
    return doc


def extract_annotation_keys(doc: fitz.Document) -> Dict[AnnotationKey, dict]:
    """Extract all annotations and their full data from a document"""
    annotations = {}

    for page_num in range(doc.page_count):
        page = doc[page_num]
        annots = list(page.annots()) if page.annots() else []

        for annot in annots:
            info = annot.info
            key = AnnotationKey(
                page=page_num,
                annot_type=annot.type[1],
                rect=round_rect(annot.rect),
                content=info.get('content', '') or ''
            )

            # Store the full annotation data for later reconstruction
            annotations[key] = {
                'type': annot.type,
                'rect': fitz.Rect(annot.rect),
                'info': dict(info),
                'colors': annot.colors,
                'border': annot.border,
                'opacity': annot.opacity,
                'vertices': annot.vertices if hasattr(annot, 'vertices') else None,
                'line_ends': annot.line_ends if annot.type[0] == fitz.PDF_ANNOT_LINE else None,
            }

            # For ink annotations, store the vertices (drawing paths)
            if annot.type[0] == fitz.PDF_ANNOT_INK:
                try:
                    annotations[key]['vertices'] = annot.vertices
                except:
                    pass

    return annotations


def copy_annotation(source_page: fitz.Page, target_page: fitz.Page, annot: fitz.Annot) -> bool:
    """Copy an annotation from source page to target page"""
    try:
        annot_type = annot.type[0]
        info = annot.info
        rect = annot.rect

        new_annot = None

        # Handle different annotation types
        if annot_type == fitz.PDF_ANNOT_FREE_TEXT:
            new_annot = target_page.add_freetext_annot(
                rect,
                info.get('content', ''),
                fontsize=11,
                text_color=(0, 0, 0),
                fill_color=annot.colors.get('fill') if annot.colors else None,
            )

        elif annot_type == fitz.PDF_ANNOT_TEXT:
            new_annot = target_page.add_text_annot(
                rect.tl,  # Top-left point
                info.get('content', ''),
            )

        elif annot_type == fitz.PDF_ANNOT_HIGHLIGHT:
            # For highlights, we need the quad points
            quads = annot.vertices
            if quads:
                new_annot = target_page.add_highlight_annot(quads=quads)
            else:
                new_annot = target_page.add_highlight_annot(rect)

        elif annot_type == fitz.PDF_ANNOT_UNDERLINE:
            quads = annot.vertices
            if quads:
                new_annot = target_page.add_underline_annot(quads=quads)
            else:
                new_annot = target_page.add_underline_annot(rect)

        elif annot_type == fitz.PDF_ANNOT_STRIKE_OUT:
            quads = annot.vertices
            if quads:
                new_annot = target_page.add_strikeout_annot(quads=quads)
            else:
                new_annot = target_page.add_strikeout_annot(rect)

        elif annot_type == fitz.PDF_ANNOT_INK:
            try:
                # Ink annotation paths are stored in vertices
                ink_list = annot.vertices
                if ink_list:
                    new_annot = target_page.add_ink_annot(ink_list)
            except Exception as e:
                print(f"    Warning: Could not copy ink annotation: {e}")
                return False

        elif annot_type == fitz.PDF_ANNOT_LINE:
            # Line annotation needs start and end points
            vertices = annot.vertices
            if vertices and len(vertices) >= 2:
                new_annot = target_page.add_line_annot(
                    fitz.Point(vertices[0]),
                    fitz.Point(vertices[1])
                )

        elif annot_type == fitz.PDF_ANNOT_SQUARE:
            new_annot = target_page.add_rect_annot(rect)

        elif annot_type == fitz.PDF_ANNOT_CIRCLE:
            new_annot = target_page.add_circle_annot(rect)

        elif annot_type == fitz.PDF_ANNOT_POLYGON:
            vertices = annot.vertices
            if vertices:
                new_annot = target_page.add_polygon_annot(vertices)

        elif annot_type == fitz.PDF_ANNOT_POLYLINE:
            vertices = annot.vertices
            if vertices:
                new_annot = target_page.add_polyline_annot(vertices)

        elif annot_type == fitz.PDF_ANNOT_STAMP:
            new_annot = target_page.add_stamp_annot(rect, stamp=0)

        elif annot_type == fitz.PDF_ANNOT_CARET:
            new_annot = target_page.add_caret_annot(rect.tl)

        else:
            print(f"    Warning: Unsupported annotation type: {annot.type[1]}")
            return False

        # Apply common properties if annotation was created
        if new_annot:
            # Set colors
            if annot.colors:
                if annot.colors.get('stroke'):
                    new_annot.set_colors(stroke=annot.colors['stroke'])
                if annot.colors.get('fill'):
                    new_annot.set_colors(fill=annot.colors['fill'])

            # Set content/info
            if info.get('content'):
                new_annot.set_info(content=info['content'])
            if info.get('title'):
                new_annot.set_info(title=info['title'])

            # Set opacity
            if annot.opacity is not None and annot.opacity < 1.0:
                new_annot.set_opacity(annot.opacity)

            new_annot.update()
            return True

        return False

    except Exception as e:
        print(f"    Error copying annotation: {e}")
        return False


def merge_pdf_annotations(output_path: str, input_paths: List[str], verbose: bool = True) -> dict:
    """
    Merge annotations from multiple PDFs into a single output file.

    Args:
        output_path: Path for the merged output PDF
        input_paths: List of input PDF paths (first is used as base)
        verbose: Print progress information

    Returns:
        Dictionary with merge statistics
    """
    if len(input_paths) < 2:
        raise ValueError("Need at least 2 input files to merge")

    stats = {
        'base_annotations': 0,
        'merged_annotations': 0,
        'failed_annotations': 0,
        'duplicate_annotations': 0,
        'files_processed': len(input_paths),
    }

    # Create temp directory for repaired PDFs
    temp_dir = tempfile.mkdtemp(prefix="pdf_merger_")

    try:
        # Repair and open the base document (first file)
        base_path = input_paths[0]
        if verbose:
            print(f"Repairing base file: {os.path.basename(base_path)}")

        repaired_base = repair_pdf(base_path, temp_dir)
        base_doc = open_pdf_safe(repaired_base)
        base_keys = extract_annotation_keys(base_doc)
        stats['base_annotations'] = len(base_keys)

        if verbose:
            print(f"  Base annotations: {len(base_keys)}")

        # Track all annotations we've seen
        all_seen_keys: Set[AnnotationKey] = set(base_keys.keys())

        # Process each additional file
        for input_path in input_paths[1:]:
            if verbose:
                print(f"\nRepairing: {os.path.basename(input_path)}")

            repaired_source = repair_pdf(input_path, temp_dir)

            if verbose:
                print(f"Processing: {os.path.basename(input_path)}")

            source_doc = open_pdf_safe(repaired_source)
            source_keys = extract_annotation_keys(source_doc)

            if verbose:
                print(f"  Total annotations: {len(source_keys)}")

            # Find annotations unique to this file
            unique_keys = set(source_keys.keys()) - all_seen_keys

            if verbose:
                print(f"  Unique annotations to merge: {len(unique_keys)}")

            # Copy unique annotations to base document
            merged_count = 0
            failed_count = 0

            for page_num in range(source_doc.page_count):
                source_page = source_doc[page_num]
                target_page = base_doc[page_num]

                annots = list(source_page.annots()) if source_page.annots() else []

                for annot in annots:
                    info = annot.info
                    key = AnnotationKey(
                        page=page_num,
                        annot_type=annot.type[1],
                        rect=round_rect(annot.rect),
                        content=info.get('content', '') or ''
                    )

                    if key in unique_keys:
                        if copy_annotation(source_page, target_page, annot):
                            merged_count += 1
                            all_seen_keys.add(key)
                            if verbose:
                                content_preview = key.content[:30] + "..." if len(key.content) > 30 else key.content
                                print(f"    + Page {page_num + 1}: {key.annot_type} - {content_preview or '(drawing)'}")
                        else:
                            failed_count += 1

            source_doc.close()
            stats['merged_annotations'] += merged_count
            stats['failed_annotations'] += failed_count
            stats['duplicate_annotations'] += len(source_keys) - len(unique_keys)

        # Save the merged document
        if verbose:
            print(f"\nSaving merged document: {output_path}")

        base_doc.save(output_path, garbage=4, deflate=True, clean=True)
        base_doc.close()

    finally:
        # Clean up temp directory
        try:
            shutil.rmtree(temp_dir)
        except:
            pass

    if verbose:
        print("\n" + "=" * 50)
        print("MERGE COMPLETE")
        print("=" * 50)
        print(f"Files processed: {stats['files_processed']}")
        print(f"Base annotations: {stats['base_annotations']}")
        print(f"New annotations merged: {stats['merged_annotations']}")
        print(f"Duplicates skipped: {stats['duplicate_annotations']}")
        if stats['failed_annotations'] > 0:
            print(f"Failed to copy: {stats['failed_annotations']}")
        print(f"Output: {output_path}")

    return stats


def main():
    if len(sys.argv) < 4:
        print(__doc__)
        print("\nError: Need at least output file and 2 input files")
        print("Example: python merge_pdf_annotations.py merged.pdf file1.pdf file2.pdf")
        sys.exit(1)

    output_path = sys.argv[1]
    input_paths = sys.argv[2:]

    # Validate inputs
    for path in input_paths:
        if not os.path.exists(path):
            print(f"Error: File not found: {path}")
            sys.exit(1)

    # Check output won't overwrite input
    output_abs = os.path.abspath(output_path)
    for path in input_paths:
        if os.path.abspath(path) == output_abs:
            print(f"Error: Output file cannot be the same as an input file")
            sys.exit(1)

    try:
        stats = merge_pdf_annotations(output_path, input_paths)
        sys.exit(0)
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
