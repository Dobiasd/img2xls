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