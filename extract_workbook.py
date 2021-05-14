#!/usr/bin/python
import os
import sys
import shutil
import zipfile
from oletools.olevba import VBA_Parser


def writeVBA(workbookName):
    vbaparser = VBA_Parser('./'+workbookName)

    VBA_path = './src/VBA'
    if os.path.exists(VBA_path):
        shutil.rmtree(VBA_path)
    os.mkdir(VBA_path)

    for (filename, stream_path, vba_filename, vba_code) in vbaparser.extract_macros():
        f = open('./src/VBA/' + vba_filename, "w")
        f.write(vba_code)
        f.close()

    vbaparser.close()


def extractWorkbook(workbookName):
    with zipfile.ZipFile('./'+workbookName, 'r') as zip_ref:
        zip_ref.extractall('./src/workbook')


if __name__ == "__main__":
    if len(sys.argv) == 2:
        workbookName = sys.argv[1]
        writeVBA(workbookName)
        extractWorkbook(workbookName)
