""" converts files of following type 2 pdf"""
import os
import sys
import shutil
import comtypes.client
import img2pdf

ppt_formats = ['.ppt', '.pptx']
word_formats = ['.doc', '.docx']
excel_formats = ['.xls', '.xlsx', '.csv']
img_formats = ['.png', '.jpg', '.gif', 'webp']
format_dictionary = {
    'WORD': word_formats,
    'PPT': ppt_formats,
    'EXCEL': excel_formats,
    'IMG': img_formats
}


def PPTtoPDF(inputFileName, outputFileName, formatType=32):
    """ converts ppts 2 pdf """
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    outputFileName = outputFileName + ".pdf"
    deck = powerpoint.Presentations.Open(inputFileName, WithWindow=False)
    deck.SaveAs(outputFileName, formatType)
    deck.Close()
    powerpoint.Quit()


def WordtoPDF(inputFileName, outputFileName, formatType=17):
    """ converts word 2 pdf """
    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = False
    outputFileName = outputFileName + ".pdf"
    deck = word.Documents.Open(inputFileName)
    deck.SaveAs(outputFileName, formatType)
    deck.Close()
    word.Quit()


def ExceltoPDF(inputFileName, outputFileName, formatType=56):
    """ coverts excel 2 pdf """
    excel = comtypes.client.CreateObject("Excel.Application")
    excel.Visible = False
    outputFileName = outputFileName + ".pdf"
    deck = excel.Workbooks.Open(inputFileName)
    deck.ExportAsFixedFormat(0, outputFileName, 1, 0)
    deck.Close()
    excel.Quit()


def ImagetoPDF(inputFileName, outputFileName):
    """ converts image 2 pdf """
    with open(f"{outputFileName}.pdf","wb") as f:
        f.write(img2pdf.convert(inputFileName))


def convert():
    """ main method """
    cmd = sys.argv[1:]
    formats = []
    if len(cmd) == 0:
        formats = ppt_formats+word_formats+excel_formats
    elif len(cmd) == 2 and cmd[0] == '-f':
        if cmd[1] == 'word':
            formats = word_formats
        elif cmd[1] == 'ppt':
            formats = ppt_formats
        elif cmd[1] == 'excel':
            formats = excel_formats
        elif cmd[1] == 'img':
            formats = img_formats
        elif cmd[1] == '*':
            formats = ppt_formats+word_formats+excel_formats+img_formats
        else:
            print("Invalid format.\nUse: python -f <word/ppt/excel/img/*>")
    else:
        print("Invalid format.\nUse: python -f <word/ppt/excel/img/*>")

    out_path = os.path.abspath("PDF")
    files = os.listdir()
    files.sort()

    for i in files:
        pos = i.rfind('.')
        if pos != -1:
            file, extension = out_path+r'\\'+i[:pos], i[pos:]
            if extension in formats:
                if i.startswith('~$') and i[2:] in files:
                    continue
                if extension in format_dictionary['WORD']:
                    WordtoPDF(os.path.abspath(i), file)
                elif extension in format_dictionary['PPT']:
                    PPTtoPDF(os.path.abspath(i), file)
                elif extension in format_dictionary['IMG']:
                    ImagetoPDF(os.path.abspath(i), file)
                elif extension in format_dictionary['EXCEL']:
                    ExceltoPDF(os.path.abspath(i), file)
                print(i, ": CONVERTED")


if __name__ == '__main__':
    if "PDF" not in os.listdir():
        os.mkdir("PDF")
    convert()
