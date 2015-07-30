# img2xls
Convert images to colored cells in an Excel spreadsheet.

![screenshot](screenshot.png "screenshot")

## Install dependencies

    pip3 install Pillow
    pip3 install xlwt3

## Usage

    python3 img2xls.py libre image.png

image.png.xls will be created.

Since different spredsheet programs processes the cell sizes in different ways, you can use `mac` or `ms` instead of `libre` for Mac Excel or MS Excel output format respectively.

If you want a grid in the resulting spreadsheet, you can pass the --grid option, as follows:

    python3 img2xls.py libre --grid vertical_gap_in_px horizontal_gap_in_px image.png

Both values have to be specified. If you don't want grid lines on an axis just set this value to 0. Negative values are ignored.
![screenshot2](screenshot2.png "screenshot2")
