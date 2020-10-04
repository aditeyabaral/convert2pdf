# Convert2PDF
A Python3 application that converts multiple Office files into their PDF versions automatically. 

Convert2PDF takes in a file type as input and exports all matching file extensions for that Office format (such as, a word document may have the extension .doc or .docx) and saves them in a separate directory, thus saving you the hassle of converting them all manually or looking for online converters. Made during the COVID-19 outbreak to kill Quarantine Boredom :P

*Since comtypes primarily supports Windows, Convert2PDF will not work on other platforms* :(

# Getting Started
Convert2PDF requires **comtypes** and **img2pdf** to be installed. This can be done using a simple pip command.

```pip install -r requirements.txt``` 

That's it. You're good to go!

# How to use Convert2PDF
### To convert all files in a directory
You can convert all files in a directory using 
```python Convert2PDF.py```
or
```python Convert2PDF.py -f *```

### To convert specific formats
You can also explicitly mention which files you would like to convert. To specify a particular type, pass in the respective format paramter as a command line argument. 

##### For all Word Document files
```python Convert2PDF.py -f word``` 

##### For all Powerpoint files
```python Convert2PDF.py -f ppt``` 

##### For all Excel Spreadsheets
```python Convert2PDF.py -f excel``` 

##### For all Image files
```python Convert2PDF.py -f img``` 

### Missing file formats 
A list of various file formats has been declared at the top section of the code. Don't see a file extension you need? You can add it in!
