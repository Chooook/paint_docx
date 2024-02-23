from dataclasses import dataclass

from docx.shared import RGBColor


@dataclass
class Color:  # Нет готовых реализаций на PyPi?
    # Basic:
    red     = RGBColor(255, 0, 0)
    maroon  = RGBColor(128, 0, 0)
    yellow  = RGBColor(255, 255, 0)
    olive   = RGBColor(128, 128, 0)
    lime    = RGBColor(0, 255, 0)
    green   = RGBColor(0, 128, 0)
    aqua    = RGBColor(0, 255, 255)
    teal    = RGBColor(0, 128, 128)
    blue    = RGBColor(0, 0, 255)
    navy    = RGBColor(0, 0, 128)
    fuchsia = RGBColor(255, 0, 255)
    purple  = RGBColor(128, 0, 128)
    black   = RGBColor(0, 0, 0)
    gray    = RGBColor(128, 128, 128)
    white   = RGBColor(255, 255, 255)
    # Alt:
    darkblue = navy
    magenta  = fuchsia
    cyan     = aqua
    # Extended:
    orange     = RGBColor(255, 165, 0)
    pink       = RGBColor(255, 20, 147)
    coral      = RGBColor(240, 128, 128)
    violet     = RGBColor(138, 43, 226)
    aquamarine = RGBColor(127, 255, 212)

    def __getitem__(self, item: str) -> RGBColor:
        try:
            return getattr(self, item)
        except AttributeError:
            return self.red


class Index:  # Зачем?
    first = 0
    last = -1
