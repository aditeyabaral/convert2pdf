# Convert2PDF
A Python3 application that converts multiple Office files into their PDF versions automatically. Convert2PDF takes in a file type as input and exports all matching file extensions for that Office format (such as, a word document may have the extension .doc or .docx) and saves them in a separate directory, thus saving you the hassle of converting them all manually or looking for online converters. Made during the COVID-19 outbreak to kill Quarantine Boredom :P

# How to use Convert2PDF
### To convert all files in a directory
```python Convert2PDF.py``` or ```python Convert2PDF.py -f *```
### To convert specific formats
You can also explicitly mention the file format using command line arguments. 
```python Convert2PDF.py -f word``` for all Word Document files

```python Convert2PDF.py -f ppt``` for all Powerpoint files

```python Convert2PDF.py -f excel``` for all Excel Spreadsheets

A list of various file formats has been declared at the top section of the code. Don't see a file extension you need? You can add it in!
