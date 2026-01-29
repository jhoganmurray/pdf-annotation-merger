#!/usr/bin/env python3
"""
PDF Annotation Merger - Windows GUI Application

Merges annotations from multiple versions of the same PDF by:
1. Extracting annotations to a data format (similar to Adobe's XFDF export)
2. Deduplicating based on position, type, and content
3. Applying merged annotations to a clean base document

Requirements:
    pip install pymupdf

Usage:
    Double-click this file to run, or: python PDF_Annotation_Merger.pyw
"""

import os
import sys
import json
import datetime
import tempfile
import shutil
import hashlib
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import List, Dict, Set, Any, Optional

# Check for PyMuPDF
try:
    import fitz
except ImportError:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "Missing Dependency",
        "PyMuPDF is not installed.\n\n"
        "Please open Command Prompt and run:\n"
        "pip install pymupdf\n\n"
        "Then try again."
    )
    sys.exit(1)


# =============================================================================
# ANNOTATION DATA MODEL (similar to XFDF concept)
# =============================================================================

def annotation_to_data(annot, page_num: int) -> Dict[str, Any]:
    """
    Extract annotation properties to a dictionary.
    This is similar to Adobe's XFDF export - separating annotation data from PDF.
    """
    data = {
        'page': page_num,
        'type': annot.type[1],
        'type_code': annot.type[0],
        'rect': [round(x, 2) for x in annot.rect],
        'content': annot.info.get('content', '') or '',
        'author': annot.info.get('title', '') or '',
        'subject': annot.info.get('subject', '') or '',
        'opacity': annot.opacity if annot.opacity >= 0 else 1.0,
    }

    # Colors
    if annot.colors:
        if annot.colors.get('stroke'):
            data['stroke_color'] = list(annot.colors['stroke'])
        if annot.colors.get('fill'):
            data['fill_color'] = list(annot.colors['fill'])

    # Vertices for ink, line, polygon, polyline annotations
    if annot.vertices:
        # Flatten nested structure and round coordinates
        vertices = []
        for item in annot.vertices:
            if isinstance(item, (list, tuple)):
                if isinstance(item[0], (list, tuple)):
                    # Nested list of points (ink annotations)
                    vertices.append([[round(p[0], 2), round(p[1], 2)] for p in item])
                else:
                    # Single point
                    vertices.append([round(item[0], 2), round(item[1], 2)])
            else:
                vertices.append(item)
        data['vertices'] = vertices

    return data


def compute_annotation_key(annot_data: Dict[str, Any]) -> str:
    """
    Compute a unique key for deduplication.
    Two annotations are considered duplicates if they have the same:
    - Page number
    - Type
    - Position (rect)
    - Content text
    """
    key_parts = [
        str(annot_data['page']),
        annot_data['type'],
        json.dumps(annot_data['rect']),
        annot_data['content'],
    ]
    key_string = '|'.join(key_parts)
    return hashlib.md5(key_string.encode()).hexdigest()


def data_to_annotation(page, annot_data: Dict[str, Any]) -> bool:
    """
    Create an annotation on a page from data dictionary.
    This is similar to Adobe's XFDF import.
    Returns True if successful.
    """
    try:
        annot_type = annot_data['type_code']
        rect = fitz.Rect(annot_data['rect'])
        content = annot_data.get('content', '')
        new_annot = None

        if annot_type == fitz.PDF_ANNOT_FREE_TEXT:
            fill = annot_data.get('fill_color')
            new_annot = page.add_freetext_annot(
                rect, content,
                fontsize=11,
                text_color=(0, 0, 0),
                fill_color=fill if fill else None,
            )

        elif annot_type == fitz.PDF_ANNOT_TEXT:
            new_annot = page.add_text_annot(rect.tl, content)

        elif annot_type == fitz.PDF_ANNOT_HIGHLIGHT:
            vertices = annot_data.get('vertices')
            if vertices:
                new_annot = page.add_highlight_annot(quads=vertices)
            else:
                new_annot = page.add_highlight_annot(rect)

        elif annot_type == fitz.PDF_ANNOT_UNDERLINE:
            vertices = annot_data.get('vertices')
            if vertices:
                new_annot = page.add_underline_annot(quads=vertices)
            else:
                new_annot = page.add_underline_annot(rect)

        elif annot_type == fitz.PDF_ANNOT_STRIKE_OUT:
            vertices = annot_data.get('vertices')
            if vertices:
                new_annot = page.add_strikeout_annot(quads=vertices)
            else:
                new_annot = page.add_strikeout_annot(rect)

        elif annot_type == fitz.PDF_ANNOT_INK:
            vertices = annot_data.get('vertices')
            if vertices:
                new_annot = page.add_ink_annot(vertices)

        elif annot_type == fitz.PDF_ANNOT_LINE:
            vertices = annot_data.get('vertices')
            if vertices and len(vertices) >= 2:
                new_annot = page.add_line_annot(
                    fitz.Point(vertices[0]),
                    fitz.Point(vertices[1])
                )

        elif annot_type == fitz.PDF_ANNOT_SQUARE:
            new_annot = page.add_rect_annot(rect)

        elif annot_type == fitz.PDF_ANNOT_CIRCLE:
            new_annot = page.add_circle_annot(rect)

        elif annot_type == fitz.PDF_ANNOT_POLYGON:
            vertices = annot_data.get('vertices')
            if vertices:
                new_annot = page.add_polygon_annot(vertices)

        elif annot_type == fitz.PDF_ANNOT_POLYLINE:
            vertices = annot_data.get('vertices')
            if vertices:
                new_annot = page.add_polyline_annot(vertices)

        elif annot_type == fitz.PDF_ANNOT_STAMP:
            new_annot = page.add_stamp_annot(rect, stamp=0)

        elif annot_type == fitz.PDF_ANNOT_CARET:
            new_annot = page.add_caret_annot(rect.tl)

        else:
            return False

        # Apply common properties
        if new_annot:
            if annot_data.get('stroke_color'):
                new_annot.set_colors(stroke=annot_data['stroke_color'])
            if annot_data.get('fill_color'):
                new_annot.set_colors(fill=annot_data['fill_color'])
            if content:
                new_annot.set_info(content=content)
            if annot_data.get('author'):
                new_annot.set_info(title=annot_data['author'])
            if annot_data.get('opacity', 1.0) < 1.0:
                new_annot.set_opacity(annot_data['opacity'])
            new_annot.update()
            return True

        return False

    except Exception:
        return False


# =============================================================================
# PDF OPERATIONS
# =============================================================================

def open_pdf_safe(filepath: str):
    """Open PDF suppressing recoverable MuPDF warnings."""
    import io
    old_stderr = sys.stderr
    sys.stderr = io.StringIO()
    try:
        doc = fitz.open(filepath)
    except Exception as e:
        sys.stderr = old_stderr
        if "xref" not in str(e).lower():
            raise
        doc = fitz.open(filepath)
    finally:
        sys.stderr = old_stderr
    return doc


def extract_annotations(filepath: str) -> List[Dict[str, Any]]:
    """
    Extract all annotations from a PDF as data dictionaries.
    Similar to Adobe Acrobat's "Export Comments" feature.
    """
    doc = open_pdf_safe(filepath)
    annotations = []

    for page_num in range(doc.page_count):
        page = doc[page_num]
        for annot in page.annots() or []:
            annot_data = annotation_to_data(annot, page_num)
            annot_data['_key'] = compute_annotation_key(annot_data)
            annot_data['_source'] = os.path.basename(filepath)
            annotations.append(annot_data)

    doc.close()
    return annotations


def merge_annotations(annotation_lists: List[List[Dict]]) -> List[Dict[str, Any]]:
    """
    Merge multiple annotation lists, removing duplicates.
    First list is considered the "base" - its annotations take precedence.
    """
    seen_keys: Set[str] = set()
    merged: List[Dict[str, Any]] = []

    for annot_list in annotation_lists:
        for annot_data in annot_list:
            key = annot_data['_key']
            if key not in seen_keys:
                seen_keys.add(key)
                merged.append(annot_data)

    return merged


def apply_annotations_to_pdf(input_path: str, output_path: str,
                              annotations: List[Dict[str, Any]],
                              progress_callback=None) -> Dict[str, int]:
    """
    Apply annotation data to a PDF, creating a new output file.
    Similar to Adobe Acrobat's "Import Comments" feature.
    """
    import io

    # Open and repair source PDF
    old_stderr = sys.stderr
    sys.stderr = io.StringIO()

    try:
        doc = fitz.open(input_path)

        # First, remove all existing annotations to start clean
        for page in doc:
            # Get list of annotations first to avoid modification during iteration
            annots_to_delete = list(page.annots() or [])
            for annot in annots_to_delete:
                page.delete_annot(annot)

    finally:
        sys.stderr = old_stderr

    # Group annotations by page for efficient processing
    by_page: Dict[int, List[Dict]] = {}
    for annot_data in annotations:
        page_num = annot_data['page']
        if page_num not in by_page:
            by_page[page_num] = []
        by_page[page_num].append(annot_data)

    # Apply annotations
    stats = {'applied': 0, 'failed': 0}
    total = len(annotations)
    done = 0

    for page_num in sorted(by_page.keys()):
        if page_num >= doc.page_count:
            continue

        page = doc[page_num]
        for annot_data in by_page[page_num]:
            if data_to_annotation(page, annot_data):
                stats['applied'] += 1
            else:
                stats['failed'] += 1

            done += 1
            if progress_callback:
                progress_callback(done / total * 100)

    # Save
    old_stderr = sys.stderr
    sys.stderr = io.StringIO()
    try:
        doc.save(output_path, garbage=4, deflate=True, clean=True)
    finally:
        sys.stderr = old_stderr

    doc.close()
    return stats


# =============================================================================
# GUI APPLICATION
# =============================================================================

class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Annotation Merger")
        self.root.geometry("650x500")
        self.root.resizable(True, True)
        self.root.minsize(550, 450)

        self.selected_files: List[str] = []
        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)

        # Title
        title_label = ttk.Label(
            main_frame,
            text="PDF Annotation Merger",
            font=('Segoe UI', 14, 'bold')
        )
        title_label.grid(row=0, column=0, pady=(0, 5), sticky="w")

        # Instructions
        instructions = ttk.Label(
            main_frame,
            text="Select multiple versions of the same PDF to merge their annotations.\n"
                 "Works like Adobe's Export/Import Comments, with automatic deduplication.",
            font=('Segoe UI', 9),
            foreground='gray'
        )
        instructions.grid(row=1, column=0, pady=(0, 10), sticky="w")

        # File list
        list_frame = ttk.LabelFrame(main_frame, text="Selected Files (first = base document)", padding="5")
        list_frame.grid(row=2, column=0, sticky="nsew", pady=(0, 10))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.file_listbox = tk.Listbox(
            list_frame,
            selectmode=tk.EXTENDED,
            font=('Consolas', 9)
        )
        self.file_listbox.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.file_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.file_listbox.config(yscrollcommand=scrollbar.set)

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, sticky="ew", pady=(0, 10))

        ttk.Button(button_frame, text="Add Files...", command=self.add_files).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="Remove Selected", command=self.remove_selected).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="Clear All", command=self.clear_all).pack(side=tk.LEFT)
        ttk.Button(button_frame, text="Move Up", command=self.move_up).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="Move Down", command=self.move_down).pack(side=tk.RIGHT)

        # Progress
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=4, column=0, sticky="ew", pady=(0, 5))

        # Status
        self.status_var = tk.StringVar(value="Select at least 2 PDF files to merge")
        self.status_label = ttk.Label(main_frame, textvariable=self.status_var, font=('Segoe UI', 9))
        self.status_label.grid(row=5, column=0, sticky="w", pady=(0, 10))

        # Action buttons frame
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=6, column=0, sticky="ew")

        # Export JSON button (optional feature)
        self.export_btn = ttk.Button(action_frame, text="Export Data...", command=self.export_json)
        self.export_btn.pack(side=tk.LEFT)

        # Merge button
        self.merge_button = ttk.Button(action_frame, text="Merge Annotations", command=self.merge_files)
        self.merge_button.pack(side=tk.RIGHT)

        try:
            style = ttk.Style()
            style.configure('Accent.TButton', font=('Segoe UI', 10, 'bold'))
        except:
            pass

    def add_files(self):
        files = filedialog.askopenfilenames(
            title="Select PDF Files to Merge",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        for f in files:
            if f not in self.selected_files:
                self.selected_files.append(f)
                self.file_listbox.insert(tk.END, os.path.basename(f))
        self.update_status()

    def remove_selected(self):
        for i in reversed(self.file_listbox.curselection()):
            self.file_listbox.delete(i)
            del self.selected_files[i]
        self.update_status()

    def clear_all(self):
        self.file_listbox.delete(0, tk.END)
        self.selected_files.clear()
        self.update_status()

    def move_up(self):
        selected = self.file_listbox.curselection()
        if not selected or selected[0] == 0:
            return
        for i in selected:
            if i > 0:
                self.selected_files[i], self.selected_files[i-1] = self.selected_files[i-1], self.selected_files[i]
                text = self.file_listbox.get(i)
                self.file_listbox.delete(i)
                self.file_listbox.insert(i-1, text)
                self.file_listbox.selection_set(i-1)

    def move_down(self):
        selected = self.file_listbox.curselection()
        if not selected or selected[-1] == self.file_listbox.size() - 1:
            return
        for i in reversed(selected):
            if i < self.file_listbox.size() - 1:
                self.selected_files[i], self.selected_files[i+1] = self.selected_files[i+1], self.selected_files[i]
                text = self.file_listbox.get(i)
                self.file_listbox.delete(i)
                self.file_listbox.insert(i+1, text)
                self.file_listbox.selection_set(i+1)

    def update_status(self):
        count = len(self.selected_files)
        if count == 0:
            self.status_var.set("Select at least 2 PDF files to merge")
        elif count == 1:
            self.status_var.set("Select at least 1 more PDF file")
        else:
            self.status_var.set(f"{count} files selected - Ready to merge")

    def export_json(self):
        """Export merged annotation data to JSON for inspection."""
        if len(self.selected_files) < 1:
            messagebox.showwarning("No Files", "Please select at least 1 PDF file.")
            return

        save_path = filedialog.asksaveasfilename(
            title="Export Annotation Data",
            defaultextension=".json",
            filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")]
        )
        if not save_path:
            return

        self.merge_button.config(state='disabled')
        self.export_btn.config(state='disabled')

        try:
            all_annotations = []
            for i, filepath in enumerate(self.selected_files):
                self.status_var.set(f"Extracting from file {i+1} of {len(self.selected_files)}...")
                self.progress_var.set((i / len(self.selected_files)) * 100)
                self.root.update()
                all_annotations.append(extract_annotations(filepath))

            merged = merge_annotations(all_annotations)

            with open(save_path, 'w', encoding='utf-8') as f:
                json.dump(merged, f, indent=2, ensure_ascii=False)

            self.progress_var.set(100)
            messagebox.showinfo(
                "Export Complete",
                f"Exported {len(merged)} unique annotations to:\n{save_path}"
            )

        except Exception as e:
            messagebox.showerror("Export Failed", f"Error: {e}")

        finally:
            self.merge_button.config(state='normal')
            self.export_btn.config(state='normal')
            self.progress_var.set(0)
            self.update_status()

    def merge_files(self):
        if len(self.selected_files) < 2:
            messagebox.showwarning("Not Enough Files", "Please select at least 2 PDF files to merge.")
            return

        self.merge_button.config(state='disabled')
        self.export_btn.config(state='disabled')
        self.progress_var.set(0)

        try:
            # Phase 1: Extract annotations from all files
            all_annotations = []
            for i, filepath in enumerate(self.selected_files):
                self.status_var.set(f"Extracting annotations from file {i+1} of {len(self.selected_files)}...")
                self.progress_var.set((i / len(self.selected_files)) * 30)
                self.root.update()
                all_annotations.append(extract_annotations(filepath))

            # Phase 2: Merge and deduplicate
            self.status_var.set("Deduplicating annotations...")
            self.progress_var.set(40)
            self.root.update()

            base_count = len(all_annotations[0])
            merged = merge_annotations(all_annotations)
            new_count = len(merged) - base_count

            # Phase 3: Apply to output
            script_dir = os.path.dirname(os.path.abspath(__file__))
            base_name = os.path.splitext(os.path.basename(self.selected_files[0]))[0]
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = os.path.join(script_dir, f"{base_name}_MERGED_{timestamp}.pdf")

            self.status_var.set("Creating merged PDF...")
            self.root.update()

            def progress_cb(pct):
                self.progress_var.set(50 + pct * 0.5)
                self.root.update()

            stats = apply_annotations_to_pdf(
                self.selected_files[0],
                output_path,
                merged,
                progress_callback=progress_cb
            )

            self.progress_var.set(100)
            self.status_var.set("Merge complete!")

            messagebox.showinfo(
                "Merge Complete",
                f"Successfully merged annotations!\n\n"
                f"Files processed: {len(self.selected_files)}\n"
                f"Total unique annotations: {len(merged)}\n"
                f"New annotations added: {new_count}\n"
                f"Duplicates removed: {sum(len(a) for a in all_annotations) - len(merged)}\n\n"
                f"Output saved to:\n{output_path}"
            )

        except Exception as e:
            messagebox.showerror("Merge Failed", f"An error occurred:\n\n{str(e)}")
            self.status_var.set("Merge failed")

        finally:
            self.merge_button.config(state='normal')
            self.export_btn.config(state='normal')
            self.progress_var.set(0)


def main():
    root = tk.Tk()
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
    app = PDFMergerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
