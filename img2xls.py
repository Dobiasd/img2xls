#!/usr/bin/python3
"""Convert images to colored cells in an Excel spreadsheet.
"""
import sys
import xlwt
import os
import re
from PIL import Image

c_width = 0

def img2xls(c_width, img_path, xls_path):
    # Load image.
    im = Image.open(img_path)
    im = im.convert('RGB')
    width, height = im.size

    # Create Table.
    book = xlwt.Workbook()

    # Create new sheet with valid name.
    sheet1 = book.add_sheet(re.sub('[^\.0-9a-zA-Z]+', '',
        os.path.basename(img_path)))

    # Scale image down if needed.
    if width > 256 or height > 256:
        factor = 256.0 / max(width, height)
        im = im.resize((int(factor * width), int(factor * height)),
            Image.BILINEAR)
        width, height = im.size

    #  Reduce image colors.
    colCnt = 63-8 # Excel does not allow more custom colors.
    palImg = im.convert('P', palette=Image.ADAPTIVE, colors=colCnt)

    # Pixel access to image.
    pix = im.load()
    palPix = palImg.load()

    # Gather color values.
    palVals = set()
    for y in range(height):
        for x in range(width):
            palVals.add(palPix[x,y])

    # Generate an index for every color.
    pallookup = {}
    for idx, val in enumerate(palVals, start=8):
        pallookup[val] = idx

    # Generate custom palette for Excel
    def gen_col_name(idx):
        return "custom_colour_" + str(idx)
    alreadyHadCol = set()
    for y in range(height):
        for x in range(width):
            palcolnum = pallookup[palPix[x,y]]
            if palcolnum in alreadyHadCol:
                continue
            alreadyHadCol.add(palcolnum)
            r, g, b = pix[x,y]
            xlwt.add_palette_colour(gen_col_name(palcolnum), palcolnum)
            book.set_colour_RGB(palcolnum, r, g, b)

    # Generate cell styles with custom colors.
    style_lookup = {}
    for key, val in pallookup.items():
        style = xlwt.easyxf(
            'pattern: pattern solid, fore_colour ' + gen_col_name(val))
        style.pattern.pattern_fore_colour = val
        style_lookup[val] = style

    # Scale table cells.
    maxEdge = max(width, height)
    colWidth = int(c_width / maxEdge)
    rowHeight = int(10000 / maxEdge)
    for x in range(width):
        col = sheet1.col(x).width = colWidth
    for y in range(height):
        row = sheet1.row(y).height = rowHeight

    # Set cell colors.
    for y in range(height):
        for x in range(width):
            palcolnum = pallookup[palPix[x,y]]
            style = style_lookup[palcolnum]
            sheet1.write(y, x, ' ', style)

    # Save finished work of art.
    book.save(xls_path)
    print('saved', xls_path)

def main():
    if len(sys.argv) != 3:
        print("Usage: python img2xls.py args image")
        print("Usage: -l switch = LibreOffice xls")
        print("Usage: -m switch = Microsoft Office xls")
        print("Usage: -a switch = Mac Office xls")
        sys.exit(2)

    switch = sys.argv[1]
    img_path = sys.argv[2]
    xls_path = img_path + ".xls"

    if switch in ("-l", "--libre"):
        c_width = 25000
    elif switch in ("-m", "--ms"):
        c_width = 135000
    elif switch in ("-a", "--mac"):
        c_width = 135000
    else:
        print("Usage: python img2xls.py args image")
        print("Usage: -l or --libre switch = LibreOffice xls")
        print("Usage: -o or --office switch = Microsoft Office xls")
        sys.exit(2)

    img2xls(c_width, img_path, xls_path)

if __name__ == "__main__":
    sys.exit(main())
