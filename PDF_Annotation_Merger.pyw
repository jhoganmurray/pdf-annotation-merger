#!/usr/bin/env python3
"""
PDF Annotation Merger - Windows GUI Application

A simple tool to merge annotations from multiple versions of the same PDF
into a single consolidated file. Designed to solve OneDrive sync conflicts.

Requirements:
    pip install pymupdf

Usage:
    Double-click this file to run, or:
    python PDF_Annotation_Merger.pyw
"""

import os
import sys
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import List, Tuple, Set
from dataclasses import dataclass

# Check for PyMuPDF
try:
    import fitz
except ImportError:
    # Show error in GUI if possible
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
    """Round rectangle coordinates for comparison"""
    return tuple(round(x, precision) for x in rect)


def extract_annotation_keys(doc) -> dict:
    """Extract all annotations from a document"""
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
            annotations[key] = True
    return annotations


def copy_annotation(source_page, target_page, annot) -> bool:
    """Copy an annotation from source page to target page"""
    try:
        annot_type = annot.type[0]
        info = annot.info
        rect = annot.rect
        new_annot = None

        if annot_type == fitz.PDF_ANNOT_FREE_TEXT:
            new_annot = target_page.add_freetext_annot(
                rect,
                info.get('content', ''),
                fontsize=11,
                text_color=(0, 0, 0),
                fill_color=annot.colors.get('fill') if annot.colors else None,
            )

        elif annot_type == fitz.PDF_ANNOT_TEXT:
            new_annot = target_page.add_text_annot(rect.tl, info.get('content', ''))

        elif annot_type == fitz.PDF_ANNOT_HIGHLIGHT:
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
            ink_list = annot.vertices
            if ink_list:
                new_annot = target_page.add_ink_annot(ink_list)

        elif annot_type == fitz.PDF_ANNOT_LINE:
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
            return False

        if new_annot:
            if annot.colors:
                if annot.colors.get('stroke'):
                    new_annot.set_colors(stroke=annot.colors['stroke'])
                if annot.colors.get('fill'):
                    new_annot.set_colors(fill=annot.colors['fill'])
            if info.get('content'):
                new_annot.set_info(content=info['content'])
            if info.get('title'):
                new_annot.set_info(title=info['title'])
            if annot.opacity is not None and annot.opacity < 1.0:
                new_annot.set_opacity(annot.opacity)
            new_annot.update()
            return True

        return False

    except Exception:
        return False


class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Annotation Merger")
        self.root.geometry("600x450")
        self.root.resizable(True, True)

        # Set minimum size
        self.root.minsize(500, 400)

        # Store selected files
        self.selected_files: List[str] = []

        self.create_widgets()

    def create_widgets(self):
        # Main frame with padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")

        # Configure grid weights for resizing
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)

        # Title and instructions
        title_label = ttk.Label(
            main_frame,
            text="PDF Annotation Merger",
            font=('Segoe UI', 14, 'bold')
        )
        title_label.grid(row=0, column=0, pady=(0, 5), sticky="w")

        instructions = ttk.Label(
            main_frame,
            text="Select multiple versions of the same PDF to merge their annotations.\n"
                 "The first file selected will be used as the base document.",
            font=('Segoe UI', 9),
            foreground='gray'
        )
        instructions.grid(row=1, column=0, pady=(0, 10), sticky="w")

        # File list frame
        list_frame = ttk.LabelFrame(main_frame, text="Selected Files", padding="5")
        list_frame.grid(row=2, column=0, sticky="nsew", pady=(0, 10))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        # Listbox with scrollbar
        self.file_listbox = tk.Listbox(
            list_frame,
            selectmode=tk.EXTENDED,
            font=('Consolas', 9)
        )
        self.file_listbox.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.file_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.file_listbox.config(yscrollcommand=scrollbar.set)

        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, sticky="ew", pady=(0, 10))

        ttk.Button(
            button_frame,
            text="Add Files...",
            command=self.add_files
        ).pack(side=tk.LEFT, padx=(0, 5))

        ttk.Button(
            button_frame,
            text="Remove Selected",
            command=self.remove_selected
        ).pack(side=tk.LEFT, padx=(0, 5))

        ttk.Button(
            button_frame,
            text="Clear All",
            command=self.clear_all
        ).pack(side=tk.LEFT)

        ttk.Button(
            button_frame,
            text="Move Up",
            command=self.move_up
        ).pack(side=tk.RIGHT, padx=(5, 0))

        ttk.Button(
            button_frame,
            text="Move Down",
            command=self.move_down
        ).pack(side=tk.RIGHT)

        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            main_frame,
            variable=self.progress_var,
            maximum=100
        )
        self.progress_bar.grid(row=4, column=0, sticky="ew", pady=(0, 5))

        # Status label
        self.status_var = tk.StringVar(value="Select at least 2 PDF files to merge")
        self.status_label = ttk.Label(
            main_frame,
            textvariable=self.status_var,
            font=('Segoe UI', 9)
        )
        self.status_label.grid(row=5, column=0, sticky="w", pady=(0, 10))

        # Merge button
        self.merge_button = ttk.Button(
            main_frame,
            text="Merge Annotations",
            command=self.merge_files,
            style='Accent.TButton'
        )
        self.merge_button.grid(row=6, column=0, sticky="e")

        # Try to style the merge button (works on Windows 10+)
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
        selected_indices = list(self.file_listbox.curselection())
        # Remove in reverse order to maintain correct indices
        for i in reversed(selected_indices):
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
                # Swap in list
                self.selected_files[i], self.selected_files[i-1] = \
                    self.selected_files[i-1], self.selected_files[i]

                # Swap in listbox
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
                # Swap in list
                self.selected_files[i], self.selected_files[i+1] = \
                    self.selected_files[i+1], self.selected_files[i]

                # Swap in listbox
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

    def merge_files(self):
        if len(self.selected_files) < 2:
            messagebox.showwarning(
                "Not Enough Files",
                "Please select at least 2 PDF files to merge."
            )
            return

        # Disable UI during merge
        self.merge_button.config(state='disabled')
        self.progress_var.set(0)
        self.status_var.set("Merging annotations...")
        self.root.update()

        try:
            result = self.perform_merge()

            self.progress_var.set(100)
            self.status_var.set("Merge complete!")

            messagebox.showinfo(
                "Merge Complete",
                f"Successfully merged annotations!\n\n"
                f"Files processed: {result['files_processed']}\n"
                f"Base annotations: {result['base_annotations']}\n"
                f"New annotations added: {result['merged_annotations']}\n"
                f"Duplicates skipped: {result['duplicate_annotations']}\n\n"
                f"Output saved to:\n{result['output_path']}"
            )

        except Exception as e:
            messagebox.showerror(
                "Merge Failed",
                f"An error occurred during merge:\n\n{str(e)}"
            )
            self.status_var.set("Merge failed - see error message")

        finally:
            self.merge_button.config(state='normal')
            self.progress_var.set(0)

    def perform_merge(self) -> dict:
        """Perform the actual merge operation"""
        stats = {
            'base_annotations': 0,
            'merged_annotations': 0,
            'failed_annotations': 0,
            'duplicate_annotations': 0,
            'files_processed': len(self.selected_files),
            'output_path': ''
        }

        # Determine output path (same directory as script)
        script_dir = os.path.dirname(os.path.abspath(__file__))
        base_name = os.path.splitext(os.path.basename(self.selected_files[0]))[0]
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"{base_name}_MERGED_{timestamp}.pdf"
        output_path = os.path.join(script_dir, output_filename)
        stats['output_path'] = output_path

        # Open base document
        base_doc = fitz.open(self.selected_files[0])
        base_keys = extract_annotation_keys(base_doc)
        stats['base_annotations'] = len(base_keys)

        all_seen_keys: Set[AnnotationKey] = set(base_keys.keys())

        total_files = len(self.selected_files)

        # Process each additional file
        for file_idx, input_path in enumerate(self.selected_files[1:], start=2):
            # Update progress
            progress = (file_idx - 1) / total_files * 100
            self.progress_var.set(progress)
            self.status_var.set(f"Processing file {file_idx} of {total_files}...")
            self.root.update()

            source_doc = fitz.open(input_path)
            source_keys = extract_annotation_keys(source_doc)

            # Find unique annotations
            unique_keys = set(source_keys.keys()) - all_seen_keys

            # Copy unique annotations
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
                            stats['merged_annotations'] += 1
                            all_seen_keys.add(key)
                        else:
                            stats['failed_annotations'] += 1

            stats['duplicate_annotations'] += len(source_keys) - len(unique_keys)
            source_doc.close()

        # Save merged document
        self.status_var.set("Saving merged document...")
        self.root.update()

        base_doc.save(output_path, garbage=4, deflate=True)
        base_doc.close()

        return stats


def main():
    root = tk.Tk()

    # Set DPI awareness for Windows 10+
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass

    app = PDFMergerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
