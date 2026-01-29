#!/usr/bin/env python3
"""
PDF Comment Collector

Collects unique annotations from multiple PDF versions and exports them
as an XFDF file that can be imported into Adobe Acrobat.

This tool does NOT modify any PDFs. It only reads and produces XFDF output.

Workflow:
1. Select your BASE PDF (the one you want to add comments to)
2. Select OTHER PDFs (the divergent copies with additional comments)
3. Click "Create XFDF" to generate a file with only the NEW comments
4. In Adobe Acrobat, open your base PDF and use:
   Comment → Import Comments → select the XFDF file

Requirements:
    pip install pymupdf
"""

import os
import sys
import hashlib
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import List, Dict, Any, Optional
import xml.etree.ElementTree as ET
from xml.dom import minidom

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
# ANNOTATION EXTRACTION
# =============================================================================

def extract_annotations(filepath: str) -> List[Dict[str, Any]]:
    """Extract all annotations from a PDF as dictionaries."""
    doc = fitz.open(filepath)
    annotations = []

    for page_num in range(doc.page_count):
        page = doc[page_num]
        page_height = page.rect.height

        for annot in page.annots() or []:
            data = {
                'page': page_num,
                'type': annot.type[1],
                'rect': list(annot.rect),
                'content': annot.info.get('content', '') or '',
                'author': annot.info.get('title', '') or '',
                'subject': annot.info.get('subject', '') or '',
                'opacity': annot.opacity if annot.opacity >= 0 else 1.0,
                'page_height': page_height,
                'flags': annot.flags,
            }

            # Border style (line width)
            if annot.border and annot.border.get('width', 0) > 0:
                data['border_width'] = annot.border['width']

            # Colors
            if annot.colors:
                if annot.colors.get('stroke'):
                    data['stroke_color'] = list(annot.colors['stroke'])
                if annot.colors.get('fill'):
                    data['fill_color'] = list(annot.colors['fill'])

            # Vertices for ink, line, polygon annotations
            if annot.vertices:
                data['vertices'] = annot.vertices

            # Unique key for deduplication
            key_parts = [
                str(page_num),
                annot.type[1],
                ','.join(f"{x:.1f}" for x in annot.rect),
                data['content'][:100],
            ]
            data['_key'] = hashlib.md5('|'.join(key_parts).encode()).hexdigest()

            annotations.append(data)

    doc.close()
    return annotations


# =============================================================================
# XFDF GENERATION
# =============================================================================

def rgb_to_hex(rgb: List[float]) -> str:
    """Convert RGB (0-1 range) to hex color."""
    if not rgb or len(rgb) < 3:
        return "#FFFF00"
    r, g, b = [int(c * 255) for c in rgb[:3]]
    return f"#{r:02X}{g:02X}{b:02X}"


def rect_to_xfdf(rect: List[float], page_height: float) -> str:
    """Convert rect to XFDF format (PDF coordinates)."""
    x0, y0, x1, y1 = rect
    pdf_y0 = page_height - y1
    pdf_y1 = page_height - y0
    return f"{x0:.2f},{pdf_y0:.2f},{x1:.2f},{pdf_y1:.2f}"


def ink_to_gestures(vertices, page_height: float) -> List[str]:
    """Convert ink vertices to XFDF gesture strings."""
    gestures = []
    for stroke in vertices:
        if isinstance(stroke, list) and stroke:
            points = []
            for pt in stroke:
                if isinstance(pt, (list, tuple)) and len(pt) >= 2:
                    x, y = pt[0], pt[1]
                    pdf_y = page_height - y
                    points.append(f"{x:.2f},{pdf_y:.2f}")
            if points:
                gestures.append(";".join(points))
    return gestures


def vertices_to_coords(vertices, page_height: float) -> str:
    """Convert vertices to XFDF coords format."""
    coords = []
    for v in vertices:
        if isinstance(v, (list, tuple)) and len(v) >= 2:
            x, y = v[0], v[1]
            pdf_y = page_height - y
            coords.extend([f"{x:.2f}", f"{pdf_y:.2f}"])
    return ",".join(coords)


def create_xfdf(annotations: List[Dict], base_pdf_name: str) -> str:
    """Create XFDF XML string from annotations."""
    ns = "http://ns.adobe.com/xfdf/"
    ET.register_namespace('', ns)

    root = ET.Element('xfdf', xmlns=ns)
    root.set('xml:space', 'preserve')

    f_elem = ET.SubElement(root, 'f')
    f_elem.set('href', base_pdf_name)

    annots = ET.SubElement(root, 'annots')

    type_map = {
        'Text': 'text', 'FreeText': 'freetext', 'Highlight': 'highlight',
        'Underline': 'underline', 'StrikeOut': 'strikeout', 'Squiggly': 'squiggly',
        'Ink': 'ink', 'Line': 'line', 'Square': 'square', 'Circle': 'circle',
        'Polygon': 'polygon', 'PolyLine': 'polyline', 'Stamp': 'stamp',
        'Caret': 'caret', 'FileAttachment': 'fileattachment',
    }

    for i, a in enumerate(annotations):
        annot_type = a['type']
        xfdf_type = type_map.get(annot_type, annot_type.lower())
        page_height = a.get('page_height', 792)

        elem = ET.SubElement(annots, xfdf_type)
        elem.set('page', str(a['page']))
        elem.set('rect', rect_to_xfdf(a['rect'], page_height))
        elem.set('name', f"annot_{i}_{a['_key'][:8]}")

        # Color
        color = a.get('stroke_color') or a.get('fill_color')
        if color:
            elem.set('color', rgb_to_hex(color))

        # Line width
        if a.get('border_width'):
            elem.set('width', f"{a['border_width']:.2f}")

        # Opacity
        if a.get('opacity', 1.0) < 1.0:
            elem.set('opacity', f"{a['opacity']:.2f}")

        # Flags
        if a.get('flags'):
            elem.set('flags', str(a['flags']))

        # Author
        if a.get('author'):
            elem.set('title', a['author'])

        # Subject
        if a.get('subject'):
            elem.set('subject', a['subject'])

        # Contents
        if a.get('content'):
            contents = ET.SubElement(elem, 'contents')
            contents.text = a['content']

        # Type-specific handling
        if annot_type in ('Highlight', 'Underline', 'StrikeOut', 'Squiggly'):
            if a.get('vertices'):
                elem.set('coords', vertices_to_coords(a['vertices'], page_height))

        elif annot_type == 'Ink':
            if a.get('vertices'):
                inklist = ET.SubElement(elem, 'inklist')
                for gesture_str in ink_to_gestures(a['vertices'], page_height):
                    gesture = ET.SubElement(inklist, 'gesture')
                    gesture.text = gesture_str

        elif annot_type == 'Line':
            if a.get('vertices') and len(a['vertices']) >= 2:
                v = a['vertices']
                start = v[0] if isinstance(v[0], (list, tuple)) else v[:2]
                end = v[1] if isinstance(v[1], (list, tuple)) else v[2:4]
                elem.set('start', f"{start[0]:.2f},{page_height - start[1]:.2f}")
                elem.set('end', f"{end[0]:.2f},{page_height - end[1]:.2f}")

        elif annot_type in ('Polygon', 'PolyLine'):
            if a.get('vertices'):
                elem.set('vertices', vertices_to_coords(a['vertices'], page_height))

    xml_str = ET.tostring(root, encoding='unicode')
    dom = minidom.parseString(xml_str)
    return dom.toprettyxml(indent="  ", encoding=None)


# =============================================================================
# GUI APPLICATION
# =============================================================================

class CommentCollectorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Comment Collector")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        self.root.minsize(500, 400)

        self.base_file: Optional[str] = None
        self.other_files: List[str] = []

        self.create_widgets()

    def create_widgets(self):
        main = ttk.Frame(self.root, padding="10")
        main.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main.columnconfigure(0, weight=1)

        # Title
        ttk.Label(main, text="PDF Comment Collector",
                  font=('Segoe UI', 14, 'bold')).grid(row=0, column=0, sticky="w")

        # Instructions
        instructions = (
            "Collects NEW comments from divergent PDF copies and exports them\n"
            "as an XFDF file you can import into Adobe Acrobat.\n\n"
            "1. Select your BASE PDF (the master you want to update)\n"
            "2. Select OTHER PDFs (copies with additional comments)\n"
            "3. Click 'Create XFDF' to generate the import file\n"
            "4. In Acrobat: Comment → Import Comments → select the XFDF"
        )
        ttk.Label(main, text=instructions, font=('Segoe UI', 9),
                  foreground='gray').grid(row=1, column=0, sticky="w", pady=(5, 15))

        # BASE FILE section
        base_frame = ttk.LabelFrame(main, text="Step 1: Base PDF", padding="5")
        base_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        base_frame.columnconfigure(0, weight=1)

        self.base_label = ttk.Label(base_frame, text="No file selected",
                                     font=('Consolas', 9), foreground='gray')
        self.base_label.grid(row=0, column=0, sticky="w")

        ttk.Button(base_frame, text="Select Base PDF...",
                   command=self.select_base).grid(row=0, column=1, padx=(10, 0))

        # OTHER FILES section
        other_frame = ttk.LabelFrame(main, text="Step 2: Other PDFs (with new comments)", padding="5")
        other_frame.grid(row=3, column=0, sticky="nsew", pady=(0, 10))
        other_frame.columnconfigure(0, weight=1)
        other_frame.rowconfigure(0, weight=1)
        main.rowconfigure(3, weight=1)

        self.other_listbox = tk.Listbox(other_frame, font=('Consolas', 9), height=6)
        self.other_listbox.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(other_frame, orient="vertical",
                                   command=self.other_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.other_listbox.config(yscrollcommand=scrollbar.set)

        btn_frame = ttk.Frame(other_frame)
        btn_frame.grid(row=1, column=0, columnspan=2, sticky="w", pady=(5, 0))
        ttk.Button(btn_frame, text="Add Files...",
                   command=self.add_others).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="Remove Selected",
                   command=self.remove_selected).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="Clear All",
                   command=self.clear_others).pack(side=tk.LEFT)

        # Progress & Status
        self.progress_var = tk.DoubleVar()
        ttk.Progressbar(main, variable=self.progress_var,
                        maximum=100).grid(row=4, column=0, sticky="ew", pady=(0, 5))

        self.status_var = tk.StringVar(value="Select a base PDF and at least one other PDF")
        ttk.Label(main, textvariable=self.status_var,
                  font=('Segoe UI', 9)).grid(row=5, column=0, sticky="w", pady=(0, 10))

        # Action buttons
        action_frame = ttk.Frame(main)
        action_frame.grid(row=6, column=0, sticky="ew")

        self.create_btn = ttk.Button(action_frame, text="Create XFDF",
                                      command=self.create_xfdf)
        self.create_btn.pack(side=tk.RIGHT)

        ttk.Button(action_frame, text="Preview",
                   command=self.preview_changes).pack(side=tk.RIGHT, padx=(0, 10))

    def select_base(self):
        filepath = filedialog.askopenfilename(
            title="Select Base PDF",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        if filepath:
            self.base_file = filepath
            self.base_label.config(text=os.path.basename(filepath), foreground='black')
            self.update_status()

    def add_others(self):
        files = filedialog.askopenfilenames(
            title="Select Other PDFs",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        for f in files:
            if f not in self.other_files and f != self.base_file:
                self.other_files.append(f)
                self.other_listbox.insert(tk.END, os.path.basename(f))
        self.update_status()

    def remove_selected(self):
        for i in reversed(self.other_listbox.curselection()):
            self.other_listbox.delete(i)
            del self.other_files[i]
        self.update_status()

    def clear_others(self):
        self.other_listbox.delete(0, tk.END)
        self.other_files.clear()
        self.update_status()

    def update_status(self):
        if not self.base_file:
            self.status_var.set("Select a base PDF")
        elif not self.other_files:
            self.status_var.set("Select at least one other PDF")
        else:
            self.status_var.set(f"Ready: 1 base + {len(self.other_files)} other file(s)")

    def collect_new_annotations(self):
        """Extract base and find new annotations from other files."""
        base_annots = extract_annotations(self.base_file)
        seen_keys = {a['_key'] for a in base_annots}
        all_new = []

        for i, filepath in enumerate(self.other_files):
            self.status_var.set(f"Reading file {i+1} of {len(self.other_files)}...")
            self.progress_var.set((i + 1) / len(self.other_files) * 80)
            self.root.update()

            other_annots = extract_annotations(filepath)
            for a in other_annots:
                if a['_key'] not in seen_keys:
                    a['_source'] = os.path.basename(filepath)
                    all_new.append(a)
                    seen_keys.add(a['_key'])

        return base_annots, all_new

    def preview_changes(self):
        """Show a summary of what would be exported."""
        if not self.base_file or not self.other_files:
            messagebox.showwarning("Missing Files",
                "Please select a base PDF and at least one other PDF.")
            return

        self.create_btn.config(state='disabled')
        self.progress_var.set(0)

        try:
            self.status_var.set("Reading base PDF...")
            self.root.update()

            base_annots, all_new = self.collect_new_annotations()
            self.progress_var.set(100)

            # Build summary
            summary = f"BASE: {os.path.basename(self.base_file)}\n"
            summary += f"  Existing comments: {len(base_annots)}\n\n"
            summary += f"NEW COMMENTS FOUND: {len(all_new)}\n\n"

            if all_new:
                by_source = {}
                for a in all_new:
                    src = a.get('_source', 'Unknown')
                    by_source.setdefault(src, []).append(a)

                for src, annots in by_source.items():
                    summary += f"From {src}: {len(annots)} new\n"
                    for a in annots[:3]:
                        content = (a['content'][:35] + "...") if len(a['content']) > 35 else a['content']
                        summary += f"  • Page {a['page']+1}: {a['type']}"
                        if content:
                            summary += f" - {content}"
                        summary += "\n"
                    if len(annots) > 3:
                        summary += f"  ... and {len(annots)-3} more\n"
                    summary += "\n"
            else:
                summary += "No new comments found."

            messagebox.showinfo("Preview", summary)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to analyze: {e}")

        finally:
            self.create_btn.config(state='normal')
            self.progress_var.set(0)
            self.update_status()

    def create_xfdf(self):
        """Create the XFDF file with new annotations."""
        if not self.base_file or not self.other_files:
            messagebox.showwarning("Missing Files",
                "Please select a base PDF and at least one other PDF.")
            return

        default_name = os.path.splitext(os.path.basename(self.base_file))[0]
        save_path = filedialog.asksaveasfilename(
            title="Save XFDF File",
            defaultextension=".xfdf",
            initialfile=f"{default_name}_new_comments.xfdf",
            filetypes=[("XFDF Files", "*.xfdf"), ("All Files", "*.*")]
        )
        if not save_path:
            return

        self.create_btn.config(state='disabled')
        self.progress_var.set(0)

        try:
            self.status_var.set("Reading base PDF...")
            self.root.update()

            base_annots, all_new = self.collect_new_annotations()

            if not all_new:
                messagebox.showinfo("No New Comments",
                    "No new comments were found.\n"
                    "All comments already exist in the base PDF.")
                return

            self.status_var.set("Creating XFDF file...")
            self.progress_var.set(90)
            self.root.update()

            xfdf_content = create_xfdf(all_new, os.path.basename(self.base_file))

            with open(save_path, 'w', encoding='utf-8') as f:
                f.write(xfdf_content)

            self.progress_var.set(100)
            self.status_var.set("XFDF created!")

            messagebox.showinfo("Success",
                f"Created XFDF with {len(all_new)} new comment(s).\n\n"
                f"Saved to:\n{save_path}\n\n"
                "To import:\n"
                "1. Open base PDF in Adobe Acrobat\n"
                "2. Comment → Import Comments\n"
                "3. Select the XFDF file\n"
                "4. Save your PDF")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to create XFDF:\n\n{e}")
            self.status_var.set("Failed")

        finally:
            self.create_btn.config(state='normal')
            self.progress_var.set(0)
            self.update_status()


def main():
    root = tk.Tk()
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
    CommentCollectorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
