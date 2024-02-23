from dataclasses import dataclass

from docx.shared import RGBColor


@dataclass
class Color:
    red = RGBColor(255, 0, 0)
    orange = RGBColor(255, 255, 0)
    yellow = RGBColor(0, 255, 255)
    green = RGBColor(0, 255, 0)
    blue = RGBColor(0, 0, 255)
    purple = RGBColor(255, 0, 255)

    def __getitem__(self, item):
        try:
            return getattr(self, item)
        except AttributeError:
            return self.red


class Index:
    first = 0
    last = -1
