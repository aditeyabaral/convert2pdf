# Convert2PDF
A Python3 application that converts multiple Office files into their PDF versions automatically. Convert2PDF takes in a file type as input and exports all matching file extensions for that Office format (such as, a word document may have the extension .doc or .docx) and saves them in a separate directory, thus saving you the hassle of converting them all manually or looking for online converters. Made during the COVID-19 outbreak to kill Quarantine Boredom :P

# How to use Convert2PDF
### To convert all files in a directory
```python Convert2PDF.py```
## To convert specific formats
You can also explicityly mention the file format using command line arguments. 
```
python Convert2PDF.py -f word
python Convert2PDF.py -f ppt
python Convert2PDF.py -f excel
python Convert2PDF.py -f *
```
