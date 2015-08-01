#!/usr/bin/python3
"""Convert images to colored cells in an Excel spreadsheet.
"""
import sys
import xlwt
import os
import re
from PIL import Image
import argparse

# constants for accessing the correct style-variation
STYLE_CNT = 4
STYLE_OFFSET_BORDER_NONE = 0
STYLE_OFFSET_BORDER_TOP = 1
STYLE_OFFSET_BORDER_LEFT = 2
STYLE_OFFSET_BORDER_TOPLEFT = 3

def load_image_rgb(path):
    """Ensures the image to be in RGB format."""
    img = Image.open(path)
    return img.convert('RGB')

def prepare_image(img):
    """Scales down if needed"""
    width, height = img.size
    if width > 256 or height > 256:
        fact = 256.0 / max(width, height)
        img = img.resize((int(fact*width), int(fact*height)), Image.BILINEAR)
    return img

def map2d(size, func):
    """Apply function to every point in [(0,0) ... (width-1, height-1)]."""
    width, height = size
    for y_pos in range(height):
        for x_pos in range(width):
            func(x_pos, y_pos)

def get_col_reduced_palette_image(img):
    """Returns image reduced to in Excel allowed number of colors."""
    cust_col_num_range = (8, 64)
    col_cnt = cust_col_num_range[1] - cust_col_num_range[0]
    pal_img = img.convert('P', palette=Image.ADAPTIVE, colors=col_cnt)
    pal_pixels = pal_img.load()
    def add_col_offset(x_pos, y_pos):
        """Add minimum color number to a pixel in palette image."""
        pal_pixels[x_pos, y_pos] += cust_col_num_range[0]
    map2d(pal_img.size, add_col_offset)
    return pal_img

def scale_table_cells(sheet1, img_size, c_size):
    """Adjust cell size to image resolution."""
    width, height = img_size
    c_width, c_height = c_size
    max_edge = max(width, height)
    col_width = int(c_width / max_edge)
    row_height = int(c_height / max_edge)
    for x_pos in range(width):
        sheet1.col(x_pos).width = col_width
    for y_pos in range(height):
        sheet1.row(y_pos).height = row_height

def create_workbook_with_sheet(name):
    """Removes non-alpha-numerical values in name."""
    book = xlwt.Workbook()
    valid_name = re.sub(r'[^\.0-9a-zA-Z]+', '', os.path.basename(name))
    sheet1 = book.add_sheet(valid_name)
    return book, sheet1

def gen_style_lookup(img, pal_img, book):
    """Create lookup dict for accessing spreadsheet styles by image color."""
    img_pixels = img.load()
    pal_pixels = pal_img.load()
    assert img.size == pal_img.size
    already_used_colors = set()
    style_lookup = {}

    def add_style_lookup(x_pos, y_pos):
        """Add a new style to lookup table for one pixel if needed."""
        palcolnum = pal_pixels[x_pos, y_pos]
        if palcolnum in already_used_colors:
            return
        # We need to define a style for every possible style-combination
        # In our case: 1) cell without border 2) cell with top border
        # 3) cell with left border 4) cell with top & left border
        already_used_colors.add(palcolnum)
        col_name = "custom_colour_" + str(palcolnum)
        xlwt.add_palette_colour(col_name, palcolnum)
        book.set_colour_RGB(palcolnum, *img_pixels[x_pos, y_pos])

        def add_new_style(top, left):
            """Add a new style to lookup table with potential cell border."""
            style = xlwt.easyxf('pattern: pattern solid, fore_colour ' +
                                col_name)
            style.pattern.pattern_fore_colour = palcolnum
            style.borders.top = top
            style.borders.left = left
            offset_dict = {(0, 0): STYLE_OFFSET_BORDER_NONE,
                           (1, 0): STYLE_OFFSET_BORDER_TOP,
                           (0, 1): STYLE_OFFSET_BORDER_LEFT,
                           (1, 1): STYLE_OFFSET_BORDER_TOPLEFT}
            style_lookup[palcolnum * STYLE_CNT +
                         offset_dict[(top, left)]] = style
        add_new_style(0, 0)
        add_new_style(1, 0)
        add_new_style(0, 1)
        add_new_style(1, 1)

    map2d(img.size, add_style_lookup)

    return style_lookup

def set_cell_colors(pal_img, style_lookup, sheet, grid_gap_vert=0,
                    grid_gap_horiz=0):
    """Pixelwise copies colors from image into table."""
    pal_pixels = pal_img.load()

    def write_sheet_cell(x_pos, y_pos):
        """Set a single pixel, i.e. cell, in table."""
        palcolnum = pal_pixels[x_pos, y_pos]
        style_idx_offset = STYLE_OFFSET_BORDER_NONE
        if grid_gap_vert > 0 or grid_gap_horiz > 0:
            # Testing for a positive value because x modulo a negative number
            # is always zero
            if grid_gap_vert > 0 and x_pos % grid_gap_vert == 0:
                if grid_gap_horiz > 0 and y_pos % grid_gap_horiz == 0:
                    style_idx_offset = STYLE_OFFSET_BORDER_TOPLEFT
                else:
                    style_idx_offset = STYLE_OFFSET_BORDER_LEFT
            elif grid_gap_horiz > 0 and y_pos % grid_gap_horiz == 0:
                style_idx_offset = STYLE_OFFSET_BORDER_TOP
        sheet.write(y_pos, x_pos, ' ',
                    style_lookup[palcolnum * STYLE_CNT + style_idx_offset])

    map2d(pal_img.size, write_sheet_cell)

def img2xls(c_width, img_path, xls_path, grid_gap_vert, grid_gap_horiz):
    """Convert image to spreadsheet."""
    img = load_image_rgb(img_path)
    img = prepare_image(img)
    pal_img = get_col_reduced_palette_image(img)

    book, sheet1 = create_workbook_with_sheet(img_path)

    style_lookup = gen_style_lookup(img, pal_img, book)

    set_cell_colors(pal_img, style_lookup, sheet1, grid_gap_vert,
                    grid_gap_horiz)

    scale_table_cells(sheet1, img.size, c_width)

    book.save(xls_path)
    print('saved', xls_path)

def main():
    """Parse command line and run."""
    parser = argparse.ArgumentParser()
    parser.add_argument('format', choices=['libre', 'ms', 'mac'],
help = 'Choose the office package to use. | libre -> LibreOffice xls | \
ms    -> Microsoft Office xls | mac   -> Mac Office xls')
    parser.add_argument('image', help = 'Image which should be converted into\
a spreadsheet')
    parser.add_argument('--grid', nargs=2, type=int, default=[0,0], help='Adds \
a grid to the resulting spreadsheet. Numbers represent the number of pixels \
between vertical and horizontal grid lines, respectively. 0 (zero) means no \
grid on this axis.')
    args = parser.parse_args()

    switch = args.format
    size_dict = {"libre": (25000, 10000),
                 "ms": (50000, 10000),
                 "mac": (135000, 10000)}

    grid_gap_vert = args.grid[0]
    grid_gap_horiz = args.grid[1]

    img_path = args.image
    xls_path = img_path + "." + switch + ".xls"

    img2xls(size_dict[switch], img_path, xls_path, grid_gap_vert,
            grid_gap_horiz)

if __name__ == "__main__":
    sys.exit(main())
