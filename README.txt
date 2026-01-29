PDF COMMENT COLLECTOR
=====================

Collects unique comments/annotations from multiple PDF versions and creates
an XFDF file that you can import into Adobe Acrobat.

This tool does NOT modify your PDFs - it only reads them and produces an
XFDF import file. Adobe Acrobat handles the actual import.


THE PROBLEM
-----------
When multiple people annotate the same PDF via OneDrive, sync conflicts
create divergent copies (e.g., "File.pdf", "File-JohnSmith.pdf"). Each
copy has different annotations that need to be consolidated.


THE SOLUTION
------------
1. Run this tool to extract the "delta" - comments that are NEW
2. Import the delta into your master PDF using Adobe Acrobat


HOW TO USE
----------
1. Double-click "Run_Comment_Collector.bat"

2. Select your BASE PDF (the master file you want to update)

3. Select OTHER PDFs (the divergent copies with additional comments)

4. Click "Preview Changes" to see what new comments were found

5. Click "Create XFDF File" to generate the import file

6. Open your base PDF in Adobe Acrobat and import the XFDF:
   - Acrobat Pro: Comment menu → Import Comments
   - Acrobat Reader: This feature may be limited

7. Save your PDF


REQUIREMENTS
------------
- Windows 10 or later
- Python 3.8+ (https://www.python.org/downloads/)
  * Check "Add Python to PATH" during installation
- Adobe Acrobat Pro (for importing XFDF or automatic merge)

Optional (for automatic "Merge & Save PDF" button):
- pywin32: pip install pywin32
  This allows the tool to automatically import comments via Acrobat


FILES INCLUDED
--------------
- Run_Comment_Collector.bat  - Double-click to start
- PDF_Comment_Collector.pyw  - The application
- README.txt                 - This file


WHY XFDF?
---------
XFDF is Adobe's standard format for exchanging PDF comments/annotations.
By outputting XFDF instead of directly modifying PDFs:

- We avoid PDF corruption issues
- Adobe's own import handles the merge correctly
- You maintain full control over your master PDF
- The process is reversible (just don't save if something looks wrong)


TROUBLESHOOTING
---------------
"Python is not installed":
   Download from https://www.python.org/downloads/
   Check "Add Python to PATH" during installation

"Import Comments" not available in Acrobat Reader:
   This feature requires Adobe Acrobat Pro. As a workaround,
   you can use the free trial of Acrobat Pro.

Comments don't appear after import:
   - Make sure you're importing into the correct PDF
   - Check that the XFDF file references the right PDF filename
   - Try opening the XFDF in a text editor to verify it has content
