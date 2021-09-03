#!/usr/bin/env python3
# -*- encoding: utf-8 -*-

from utils import XlsxImages, Path
from multiprocessing.pool import Pool
import argparse


def extract_one(filepath, sheet_index=1, image_field='A', name_field='B'):
    xlsx_image = XlsxImages(filepath)
    xlsx_image.unzip()
    xlsx_image.get_images(sheet_index, image_field, name_field)
    return


def extract_many(file_directory, sheet_index=1, image_field='A', name_field='B'):
    extract_pool = Pool(4)
    for xlsx in Path(file_directory).rglob(pattern='*.xlsx'):
        print(xlsx)
        extract_pool.apply(extract_one,
                           args=(xlsx, sheet_index, image_field, name_field))

    extract_pool.close()
    extract_pool.join()


def main():
    parser = argparse.ArgumentParser(
        description="Help extract the images in the excel table and"
                    "name them according to the name corresponding to a column of the table "
    )

    parser.add_argument("-a", help="Specify a *.xlsx file name")
    parser.add_argument("-d", help="Specify a directory which include all *.xlsx files")
    parser.add_argument("-s", "--sheet", default=1, type=int, help="Specify the sheet used")
    parser.add_argument("-i", "--image", default="A", type=str, help="Specify which field contains images")
    parser.add_argument("-n", "--name", default="B", type=str, help="Specify which field contains filename")

    args = parser.parse_args()

    if args.a:
        extract_one(args.a, args.sheet, args.image, args.name)

    if args.d:
        extract_many(args.d, args.sheet, args.image, args.name)


if __name__ == "__main__":
    main()
