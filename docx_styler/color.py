from docx.shared import RGBColor
from docx.text.run import Run


def color_run(run: Run, color: str) -> None:
    run.font.color.rgb = Color.get(color)


class Color:
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

    @classmethod
    def __getitem__(cls, item: str) -> RGBColor:
        try:
            return getattr(cls, item)
        except AttributeError:
            return cls.red

    @classmethod
    def get(cls, item):
        return cls.__getitem__(item)
