<br />
<p align="center"><img src="https://github.com/saltyFamiliar/ExPat/blob/master/preview.gif?raw=true")</p>
<br />

# ExPat - Patient Data Extractor

This utility creates Word files (.docx) based on a given Word template for each row of a given Excel file (.xlsx). The first row of the Excel file is excluded and is used to identify the title of each column of data. After the relevant data is extracted from the Excel file, the Word template is searched for tags to be replaced by corresponding values from the selected Excel file.


# Tags

Tags are case-senstive, alpha-numeric and prepended with a "#". In order for this utility to correctly match a tag with a value from an input Excel file, the tag must match a value from the first row of the Excel file and be located inside of a Word table cell. 


# Naming Convention

Title columns are identified upon selection of an Excel data file and used to populate a set of drop-down lists. Output files will be named according to the values set in the drop-down lists and the text input into the adjacent text fields.

# Dependencies

* PyQt5
* python-docx
* openpxl
