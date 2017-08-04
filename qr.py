# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import pyqrcode
import png
from xlutils.copy import copy # http://pypi.python.org/pypi/xlutils
from xlrd import open_workbook # http://pypi.python.org/pypi/xlrd
from xlwt import easyxf # http://pypi.python.org/pypi/xlwt
from xlwt import *
import glob
import os
from PIL import Image



rb = open_workbook("qrcodes.xlsx",)

directory = "/Users/haripriyamehta/Desktop/qrcodes2"
os.mkdir(directory)
os.chdir(directory)  

r_sheet = rb.sheet_by_index(0) # read only copy to introspect the file
wb = copy(rb) # a writable copy (I can't read values out of this, only write to it)
w_sheet = wb.get_sheet(0) # the sheet to write to within the writable copy

                     
for i in range(1, r_sheet.nrows-1):
  strtotal = str(r_sheet.cell_value(i, 3))
  pic = strtotal[:6]
  picpng = pic +".png"
  picbmp = pic+".bmp"
  qr = pyqrcode.create(strtotal)
  qr.png(picpng, scale=24)
  Image.open(picpng).convert("RGB").save(picbmp)  
  w_sheet.insert_bitmap(picbmp,i,4)

wb.save("qrcodes3.xls")



files=glob.glob('*.png')
for filename in files:
    os.unlink(filename)
    
files=glob.glob('*.bmp')
for filename in files:
    os.unlink(filename)