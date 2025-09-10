# Excel Data Automation Tool

### Overview
This is a Python-based desktop application designed for cybersecurity professionals and analysts to streamline the tedious process of managing and analyzing vulnerability data in Excel. Forget manual data entry, filtering, and repetitive copy-pasting; this tool automates key tasks, allowing you to focus on threat remediation.

### Features
* **Compare & Split:** Compares two Excel files (e.g., a "tracking" file and a "raw" data file) and separates new and old data into two distinct sheets in a new output file.
* **Highlight Matching Rows:** Automatically highlights rows in one Excel sheet that match data in a second sheet based on user-defined key columns.
* **Add Site & Product:** Merges and enriches data by taking a list of IP addresses and adding corresponding "Site" and "Product" information from a separate mapping file.
* **Add New Data:** Appends new data from one sheet to the end of an existing tracking file and highlights the newly added rows for easy identification.

### Prerequisites
To run this application, you need to have Python installed on your system.
The following libraries are required and can be installed using `pip`:
```sh
pip install pandas openpyxl numpy
