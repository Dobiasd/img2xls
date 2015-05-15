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
        factor = 256.0 / max(width, height)
        im = im.resize((int(factor * width), int(factor * height)),
            Image.BILINEAR)
    return im

def get_col_reduced_palette_image(im):
    colCnt = 63-8 # Excel does not allow more custom colors.
    palImg = im.convert('P', palette=Image.ADAPTIVE, colors=colCnt)
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

def get_pal_lookup(palImg):
    # Gather color values.
    palVals = set()
    width, height = palImg.size
    palPix = palImg.load()
    for y in range(height):
        for x in range(width):
            palVals.add(palPix[x,y])

    # Generate an index for every color.
    pallookup = {}
    for idx, val in enumerate(palVals, start=8):
        pallookup[val] = idx

    return pallookup

def gen_col_name(idx):
        return "custom_colour_" + str(idx)

def gen_custom_palette(im, palImg, pallookup, book):
    pix = im.load()
    palPix = palImg.load()
    width, height = im.size
    assert(im.size == palImg.size)
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

def set_cell_colors(palImg, pallookup, style_lookup, sheet):
    palPix = palImg.load()
    width, height = palImg.size
    for y in range(height):
        for x in range(width):
            palcolnum = pallookup[palPix[x,y]]
            style = style_lookup[palcolnum]
            sheet.write(y, x, ' ', style)

def gen_custom_colored_cell_styles(pallookup):
    style_lookup = {}
    for key, val in pallookup.items():
        style = xlwt.easyxf(
            'pattern: pattern solid, fore_colour ' + gen_col_name(val))
        style.pattern.pattern_fore_colour = val
        style_lookup[val] = style
    return style_lookup

def img2xls(c_width, img_path, xls_path):
    im = load_image_rgb(img_path)
    im = prepare_image(im)
    palImg = get_col_reduced_palette_image(im)

    book, sheet1 = create_workbook_with_sheet(img_path)

    pallookup = get_pal_lookup(palImg)

    gen_custom_palette(im, palImg, pallookup, book)

    style_lookup = gen_custom_colored_cell_styles(pallookup)

    scale_table_cells(sheet1, im.size, c_width)

    set_cell_colors(palImg, pallookup, style_lookup, sheet1)

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
