"""Модуль с функциями для покраски элементов объекта Document (Run)."""

from types import MappingProxyType
from typing import NewType

from docx.shared import RGBColor
from docx.text.run import Run

# Для цвета заливки в python-docx используется цифровой идентификатор,
#  чтобы отличить его от обычных int и для читабельности, создаём новый тип.
HighlightColor = NewType('HighlightColor', int)


class Color:
    """Класс, содержащий цвета для покраски и заливки текста.

    Позволяет извлекать цвета таким образом:
    >>> Color.get_rgb_color('color_name')
    В случае отсутствия цвета, соответствующего переданному имени
    возвращает цвет по умолчанию в соответствии с переменными класса
    """

    def __new__(cls, *args, **kwargs):
        raise TypeError('Создание экземпляров этого класса запрещено.')

    DEFAULT_COLOR = 'red'
    DEFAULT_HIGHLIGHT_COLOR = 'yellow'

    RGB = MappingProxyType({
        # Basic:
        'red':         RGBColor(255, 0, 0),
        'darkred':     RGBColor(128, 0, 0),
        'yellow':      RGBColor(255, 255, 0),
        'darkyellow':  RGBColor(128, 128, 0),
        'lime':        RGBColor(0, 255, 0),
        'darkgreen':   RGBColor(0, 128, 0),
        'aqua':        RGBColor(0, 255, 255),
        'teal':        RGBColor(0, 128, 128),
        'blue':        RGBColor(0, 0, 255),
        'navy':        RGBColor(0, 0, 128),
        'magenta':     RGBColor(255, 0, 255),
        'violet':      RGBColor(128, 0, 128),
        'black':       RGBColor(0, 0, 0),
        'gray':        RGBColor(128, 128, 128),
        'lightgray':   RGBColor(192, 192, 192),
        'white':       RGBColor(255, 255, 255),
        # Alt:
        'maroon':      RGBColor(128, 0, 0),      # darkred
        'olive':       RGBColor(128, 128, 0),    # darkyellow
        'brightgreen': RGBColor(0, 255, 0),      # lime
        'green':       RGBColor(0, 255, 0),      # lime
        'turquoise':   RGBColor(0, 255, 255),    # aqua
        'cyan':        RGBColor(0, 255, 255),    # aqua
        'darkteal':    RGBColor(0, 128, 128),    # teal
        'darkcyan':    RGBColor(0, 128, 128),    # teal
        'darkblue':    RGBColor(0, 0, 128),      # navy
        'pink':        RGBColor(255, 0, 255),    # magenta
        'fuchsia':     RGBColor(255, 0, 255),    # magenta
        'purple':      RGBColor(255, 0, 255),    # magenta
        'darkmagenta': RGBColor(128, 0, 128),    # violet
        'darkpurple':  RGBColor(128, 0, 128),    # violet
        'darkgray':    RGBColor(128, 128, 128),  # gray
        'gray50':      RGBColor(128, 128, 128),  # gray
        'gray25':      RGBColor(192, 192, 192),  # lightgray
        # Extended:
        # 'orange' = RGBColor(255, 165, 0),
        # 'pink' = RGBColor(255, 20, 147),
        # 'coral' = RGBColor(240, 128, 128),
        # 'violet' = RGBColor(138, 43, 226),
        # 'aquamarine' = RGBColor(127, 255, 212),
    })

    HIGHLIGHT_MS = MappingProxyType({
        # Basic:
        'red':         6,
        'darkred':     13,
        'yellow':      7,
        'darkyellow':  14,
        'lime':        4,
        'darkgreen':   11,  # GREEN in WD_COLOR_INDEX
        'aqua':        3,
        'teal':        10,
        'blue':        2,
        'navy':        9,
        'magenta':     5,
        'violet':      12,
        'black':       1,
        'gray':        15,
        'lightgray':   16,
        'white':       8,
        'auto':        0,
        # Alt:
        'maroon':      13,  # darkred
        'olive':       14,  # darkyellow
        'brightgreen': 4,   # lime
        'green':       4,   # lime
        'turquoise':   3,   # aqua
        'cyan':        3,   # aqua
        'darkteal':    10,  # teal
        'darkcyan':    10,  # teal
        'darkblue':    9,   # navy
        'pink':        5,   # magenta
        'fuchsia':     5,   # magenta
        'purple':      5,   # magenta
        'darkmagenta': 12,  # violet
        'darkpurple':  12,  # violet
        'darkgray':    15,  # gray
        'gray50':      15,  # gray
        'gray25':      16,  # lightgray
    })

    @classmethod
    def __getitem__(
            cls, color_format: str, color_name: str
    ) -> RGBColor | HighlightColor:
        """Метод для извлечения цвета по его наименованию."""
        try:
            return getattr(cls, color_format)[color_name.lower()]
        except KeyError:
            # в случае отсутствия цвета, соответствующего переданному имени
            if color_format == 'RGB':
                return getattr(cls, color_format)[cls.DEFAULT_COLOR]
            elif color_format == 'HIGHLIGHT_MS':
                return getattr(cls, color_format)[cls.DEFAULT_HIGHLIGHT_COLOR]

    @classmethod
    def get_rgb_color(cls, color_name: str) -> RGBColor:
        """
        Метод для доступа к __getitem__ без инициализации объекта.
        Для извлечения RGBColor цвета по его наименованию.
        """
        return cls.__getitem__(color_format='RGB',
                               color_name=color_name)

    @classmethod
    def get_highlight_color(cls, color_name: str) -> HighlightColor:
        """
        Метод для доступа к __getitem__ без инициализации объекта.
        Для извлечения FillColorMS цвета по его наименованию.
        """
        return cls.__getitem__(color_format='HIGHLIGHT_MS',
                               color_name=color_name)

    @classmethod
    def color_run(cls, run: Run, color: str = '') -> None:
        """Функция для покраски текста объекта Run.

        :param run: Run, который нужно покрасить.
        :param color: Цвет, в который нужно покрасить.
        """
        run.font.color.rgb = cls.get_rgb_color(color)

    @classmethod
    def highlight_run(cls, run: Run, color: str = '') -> None:
        """Функция для покраски текста объекта Run.

        :param run: Run, который нужно покрасить.
        :param color: Цвет, в который нужно покрасить.
        """
        run.font.highlight_color = cls.get_highlight_color(color)
