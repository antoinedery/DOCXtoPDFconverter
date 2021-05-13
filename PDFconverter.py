#####################################################################################
# Convert docx files to pdf files                                                   #
# File:     PDFConverter.py                                                         #
# Authors:  Antoine Dery                                                            #
# Thanks to Al Johri for the docx2pdf package (https://github.com/AlJohri/docx2pdf) #
##################################################################################### 

import os
import sys
import time

from docx2pdf import convert

print("Converting to PDF... This window will close once the conversion is completed")
convert(sys.argv[1])    # sys.argv[1] contains the file path name from which the python script was called
