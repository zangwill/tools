#!/home/10321887@zte.intra/work/env/bin/python3/bin/python

__author__ = "gbmumumu"

from pathlib import Path

import pytesseract
from PIL import Image


class Walker:
    def __init__(self, directory: Path):
        self._d = directory
        # clean directory name
        try:
            self._clean()
        except:
            raise FileNotFoundError

    def _clean(self):
        name = list(str(self._d.name))
        for idx, char in enumerate(name):
            if char == "（":
                name[idx] = "-"
            elif char == "）":
                name[idx] = " "
                self._d.rename(self._d.parent / Path("".join(name[:idx])))
                break

    @property
    def images(self):
        flst = []
        for file in self._d.rglob("CjYTPW*"):
            flst.append(
                file.absolute()
            )
        return flst

    @property
    def label(self):
        tmp = self._d.name.split("-")
        name, desc = tmp
        if "&" in desc:
            desc = desc.split("&")

        return {name: desc}


class XImage:
    def __init__(self, image: Path):
        self.image = image

    def parse(self):
        image = Image.open(str(self.image))
        data = pytesseract.image_to_string(
            image,
            lang="ch"
        )
        print(data)

    def extract(self):
        pass


if __name__ == "__main__":
    t = Path(r"./data/xyz)
    tx = Walker(t)
    x = XImage(tx.images[0])
    x.parse()
