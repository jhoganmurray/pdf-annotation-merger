PDF ANNOTATION MERGER
=====================

A simple tool to merge annotations from multiple versions of the same PDF
into a single consolidated file. Designed to solve OneDrive sync conflicts
where multiple users annotate the same document and create divergent copies.


REQUIREMENTS
------------
- Windows 10 or later
- Python 3.8 or later (https://www.python.org/downloads/)
  * IMPORTANT: During Python installation, check "Add Python to PATH"


HOW TO USE
----------
1. Double-click "Run_PDF_Merger.bat"
   - First run will automatically install the required PyMuPDF library

2. Click "Add Files..." and select all the PDF versions you want to merge
   - Select the file you want to use as the BASE first (or reorder with Move Up/Down)
   - The base file's annotations will be preserved, and unique annotations
     from other files will be added to it

3. Click "Merge Annotations"

4. The merged file will be saved in the same folder as the script with
   the name: OriginalName_MERGED_YYYYMMDD_HHMMSS.pdf


HOW IT WORKS
------------
- The tool compares annotations across all selected PDFs
- It identifies which annotations are duplicates (same position, type, content)
- Only unique annotations are copied to the merged output
- The base PDF content remains unchanged; only annotations are merged


SUPPORTED ANNOTATION TYPES
--------------------------
- Text boxes (FreeText)
- Highlights, underlines, strikeouts
- Ink drawings (freehand marks, circles, arrows)
- Lines, rectangles, circles, polygons
- Stamps, carets, sticky notes


TROUBLESHOOTING
---------------
"Python is not installed":
   Download and install from https://www.python.org/downloads/
   Make sure to check "Add Python to PATH" during installation

"Failed to install PyMuPDF":
   Open Command Prompt as Administrator and run:
   pip install pymupdf

The app doesn't open:
   Try running from Command Prompt:
   python PDF_Annotation_Merger.pyw


FILES INCLUDED
--------------
- Run_PDF_Merger.bat      - Double-click this to start
- PDF_Annotation_Merger.pyw - The main application
- README.txt              - This file
