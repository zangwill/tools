#!/usr/bin/env python
# -*- coding: utf-8 -*-


__author__ = "gbmumumu"

from pathlib import Path
from collections import OrderedDict
from tqdm import tqdm
import re
import shutil
import zipfile
from xml.dom.minidom import parse


class XlsxImages:
    def __init__(self, filepath, image_idx=1, symbol_idx=2,
                 work_space=Path("./data"), images_output_path=Path("./images")):
        if not isinstance(filepath, Path):
            filepath = Path(filepath)
        self._xlsx = filepath
        self._work_space = work_space / filepath.stem
        self._zip = self._work_space / self._xlsx.with_suffix(".zip").name
        self._output_path = images_output_path / filepath.stem
        self._iid = image_idx
        self._sid = symbol_idx
        try:
            self._work_space.mkdir(exist_ok=True, parents=True)
            self._output_path.mkdir(exist_ok=True, parents=True)
            shutil.copy(str(self._xlsx), str(self._zip))
        except Exception as e:
            print("Failed to initialize the file directory,"
                  "please check the file system or the permissions of this script"
                  f"error type: {e.__class__.__name__}")
            exit(1)
        else:
            print(f"{filepath.name} initialize successfully!")

    def unzip(self):
        try:
            print(f"Extracting files from {self._xlsx} to {self._work_space.absolute()}...")
            zipfile.ZipFile(self._zip).extractall(str(self._work_space))
        except Exception as e:
            print(f"File decompression failed!: {self._xlsx} "
                  f"error type: {e.__class__.__name__}")
        else:
            print("Decompression done!")
        return

    def get_shared_string_data(self):
        print("reading sharedStrings.xml...")
        shared = self._work_space / "xl" / "sharedStrings.xml"
        string_data = OrderedDict()
        tree = parse(str(shared))
        shared_data = tree.documentElement.getElementsByTagName("si")
        for idx, node in enumerate(shared_data):
            for node_i in node.childNodes:
                if node_i.tagName == "t":
                    string_data[str(idx)] = node_i.childNodes[0].nodeValue

        return string_data

    def get_sheet_data(self, index=1):
        image_rgx = re.compile(r".*DISPIMG\(\"(ID_.*)\",\d+\).*")
        print(f"reading sheet{index}")
        sheet = self._work_space / "xl" / "worksheets" / f"sheet{index}.xml"
        tree = parse(str(sheet))
        sheet_data = tree.documentElement.getElementsByTagName("sheetData")
        image_data, symbol_data = OrderedDict(), OrderedDict()
        for cell in sheet_data:
            for row in cell.getElementsByTagName("row"):
                image = row.getElementsByTagName("c")[self._iid - 1]
                symbol = row.getElementsByTagName("c")[self._sid - 1]
                image_cell = image.getAttribute("r")
                symbol_cell = symbol.getAttribute("r")
                inv, jnv = None, None
                try:
                    for node_i in image.childNodes:
                        if node_i.tagName == "v":
                            inv = node_i.childNodes[0].nodeValue
                    for node_j in symbol.childNodes:
                        if node_j.tagName == "v":
                            jnv = node_j.childNodes[0].nodeValue
                except ValueError:
                    continue
                else:
                    if jnv is not None and inv is not None:
                        image_data[image_cell] = image_rgx.findall(inv)[0]
                        symbol_data[symbol_cell] = jnv

        return image_data, symbol_data

    def get_target_data(self):
        print("reading cellimages.xml.rels")
        cell_images = self._work_space / "xl" / "_rels" / "cellimages.xml.rels"
        tree = parse(str(cell_images))
        target_root = tree.documentElement
        target_data = OrderedDict()
        for image in target_root.getElementsByTagName("Relationship"):
            target_data[image.getAttribute("Id")] = image.getAttribute("Target")
        return target_data

    def get_image_rids(self):
        r_id_with_name = self._work_space / "xl" / "cellimages.xml"
        r_id_name_tree = parse(str(r_id_with_name))
        r_id_name_root = r_id_name_tree.documentElement
        r_id_names = OrderedDict()
        r_i_ds = []
        for cell_image in r_id_name_root.getElementsByTagName("etc:cellImage"):
            a_blip_list = cell_image.getElementsByTagName("a:blip")
            xdr_cnvpr_list = cell_image.getElementsByTagName("xdr:cNvPr")
            if a_blip_list and xdr_cnvpr_list:
                a_blip = a_blip_list[0]
                xdr_cNvPr = xdr_cnvpr_list[0]
                r_id_names[xdr_cNvPr.getAttribute("name")] = a_blip.getAttribute("r:embed")
        return r_id_names

    def get_images(self, sheet_index=1, image_field='A', name_field='B'):
        image_field = image_field.upper()
        name_field = name_field.upper()
        image_data, symbol_data = self.get_sheet_data(sheet_index)
        symbols = self.get_shared_string_data()
        for item_cell, item_symbol_index in symbol_data.items():
            symbol_data[item_cell] = symbols[item_symbol_index]
        image_target = self.get_target_data()
        image_rels = self.get_image_rids()

        for cell, filename in tqdm(symbol_data.items(), desc='copying:'):
            if cell.startswith(name_field):
                src_name = Path(image_target.get(
                    image_rels.get(
                        image_data.get(image_field + re.findall(r"\d+", cell)[0])
                        )
                )).name
                src = self._work_space / "xl" / "media" / src_name
                des = self._output_path / Path(filename).with_suffix(src.suffix)
                shutil.copy(str(src), str(des))
        print(f"{self._xlsx} done!")


if __name__ == "__main__":
    pass
