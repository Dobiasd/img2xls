#!/usr/bin/python3
"""Convert images to colored cells in an Excel spreadsheet.
"""
import sys
import xlwt
import os
import re
from PIL import Image

def load_image_rgb(path):
    im = Image.open(path)
    return im.convert('RGB')

def prepare_image(im):
    """Scales down if needed"""
    width, height = im.size
    if width > 256 or height > 256:
        fact = 256.0 / max(width, height)
        im = im.resize((int(fact*width), int(fact*height)), Image.BILINEAR)
    return im

def map2d(size, func):
    width, height = size
    for y in range(height):
        for x in range(width):
            func(x, y)

def get_col_reduced_palette_image(im):
    # Excel does not allow more custom colors.
    cust_col_num_range = (8, 64)
    colCnt = cust_col_num_range[1] - cust_col_num_range[0]
    palImg = im.convert('P', palette=Image.ADAPTIVE, colors=colCnt)
    palPix = palImg.load()
    def add_col_offset(x, y):
        palPix[x,y] += cust_col_num_range[0]
    map2d(palImg.size, add_col_offset)
    return palImg

def scale_table_cells(sheet1, imgSize, c_width):
    width, height = imgSize
    maxEdge = max(width, height)
    colWidth = int(c_width / maxEdge)
    rowHeight = int(10000 / maxEdge)
    for x in range(width):
        col = sheet1.col(x).width = colWidth
    for y in range(height):
        row = sheet1.row(y).height = rowHeight

def create_workbook_with_sheet(name_suggestion):
    book = xlwt.Workbook()
    valid_name = re.sub('[^\.0-9a-zA-Z]+', '',
        os.path.basename(name_suggestion))
    sheet1 = book.add_sheet(valid_name)
    return book, sheet1

def gen_style_lookup(im, palImg, book):
    pix = im.load()
    palPix = palImg.load()
    assert(im.size == palImg.size)
    alreadyUsedColors = set()
    style_lookup = {}

    def add_style_lookup(x, y):
        palcolnum = palPix[x,y]
        if palcolnum in alreadyUsedColors:
            return
        alreadyUsedColors.add(palcolnum)
        col_name = "custom_colour_" + str(palcolnum)
        xlwt.add_palette_colour(col_name, palcolnum)
        book.set_colour_RGB(palcolnum, *pix[x,y])
        style = xlwt.easyxf('pattern: pattern solid, fore_colour ' + col_name)
        style.pattern.pattern_fore_colour = palcolnum
        style_lookup[palcolnum] = style

    map2d(im.size, add_style_lookup)

    return style_lookup

def set_cell_colors(palImg, style_lookup, sheet):
    palPix = palImg.load()
    def write_sheet_cell(x, y):
        sheet.write(y, x, ' ', style_lookup[palPix[x,y]])
    map2d(palImg.size, write_sheet_cell)

def img2xls(c_width, img_path, xls_path):
    im = load_image_rgb(img_path)
    im = prepare_image(im)
    palImg = get_col_reduced_palette_image(im)

    book, sheet1 = create_workbook_with_sheet(img_path)

    style_lookup = gen_style_lookup(im, palImg, book)

    set_cell_colors(palImg, style_lookup, sheet1)

    scale_table_cells(sheet1, im.size, c_width)

    book.save(xls_path)
    print('saved', xls_path)

def print_usage():
    print("Usage: python img2xls.py format image")
    print("                         format = libre -> LibreOffice xls")
    print("                         format = ms    -> Microsoft Office xls")
    print("                         format = mac   -> Mac Office xls")

def abort_with_usage():
    print_usage()
    sys.exit(2)

def main():
    if len(sys.argv) != 3:
        abort_with_usage()

    switch = sys.argv[1]
    img_path = sys.argv[2]
    xls_path = img_path + ".xls"

    size_dict = { "libre": 25000
                , "ms": 135000
                , "mac": 135000 }

    if not switch in size_dict:
        abort_with_usage()

    img2xls(size_dict[switch], img_path, xls_path)

if __name__ == "__main__":
    sys.exit(main())
