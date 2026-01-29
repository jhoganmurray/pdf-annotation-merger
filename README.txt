PDF COMMENT COLLECTOR & MERGER
==============================

Collects unique comments/annotations from multiple PDF versions and either:
- Creates an XFDF file for manual import into Adobe Acrobat, OR
- Merges them directly into a new PDF (no Acrobat needed!)


THE PROBLEM
-----------
When multiple people annotate the same PDF via OneDrive, sync conflicts
create divergent copies. Each copy has different annotations that need
to be consolidated into the master file.


HOW TO USE
----------
1. Double-click "Run_Comment_Collector.bat"

2. Select your BASE PDF (the master file)

3. Select OTHER PDFs (the divergent copies with new comments)

4. Click "Preview" to see what new comments were found

5. Choose your export method:

   OPTION A - Merge & Save PDF (Recommended)
   * Click "Merge & Save PDF"
   * Choose where to save the merged PDF
   * Done! The new PDF contains all comments.

   OPTION B - Create XFDF (Manual import)
   * Click "Create XFDF" to generate the import file
   * Open your base PDF in Adobe Acrobat
   * Comment menu -> Import Comments -> select the XFDF file
   * Save your PDF


REQUIREMENTS
------------
- Windows 10 or later
- Python 3.8+ (https://www.python.org/downloads/)
  * Check "Add Python to PATH" during installation

Dependencies (auto-installed on first run):
- pymupdf  - For reading PDF annotations
- pikepdf  - For direct PDF merging (optional but recommended)


FILES
-----
- Run_Comment_Collector.bat  - Double-click to start
- PDF_Comment_Collector.pyw  - The main application
- xfdf_importer.py           - PDF merge engine (pikepdf-based)
- README.txt                 - This file
