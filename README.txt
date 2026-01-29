PDF COMMENT COLLECTOR
=====================

Collects unique comments/annotations from multiple PDF versions and creates
an XFDF file that you can import into Adobe Acrobat.

This tool does NOT modify your PDFs - it only reads them and outputs XFDF.


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

5. Click "Create XFDF" to generate the import file

6. Open your base PDF in Adobe Acrobat:
   Comment menu → Import Comments → select the XFDF file

7. Save your PDF


REQUIREMENTS
------------
- Windows 10 or later
- Python 3.8+ (https://www.python.org/downloads/)
  * Check "Add Python to PATH" during installation
- Adobe Acrobat Pro (for importing the XFDF)


FILES
-----
- Run_Comment_Collector.bat  - Double-click to start
- PDF_Comment_Collector.pyw  - The application
- README.txt                 - This file
