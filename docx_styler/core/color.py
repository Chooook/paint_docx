"""Модуль с функциями для покраски Run в объекте Document."""
import warnings
from functools import singledispatch
from types import MappingProxyType
from typing import NewType, Tuple

from docx.shared import RGBColor
from docx.text.run import Run

# Для цвета заливки в python-docx используется цифровой идентификатор,
#  чтобы отличить его от обычных int и для читабельности, создаём новый тип.
HighlightColor = NewType('HighlightColor', int)

# Цвета по умолчанию:
DEFAULT_RGB_COLOR = 'red'
DEFAULT_HIGHLIGHT_COLOR = 'yellow'


def color_run(run: Run, color: str | Tuple[int, int, int] = '') -> None:
    """Изменяет цвет текста объекта Run.

    :param run: Run, который нужно покрасить.
    :param color: Цвет, в который нужно покрасить.
    """
    run.font.color.rgb = __get_rgb_color(color)


def highlight_run(run: Run, color: str = '') -> None:
    """Изменяет цвет заливки текста объекта Run.

    :param run: Run, который нужно выделить.
    :param color: Цвет, в который нужно выделить.
    """
    run.font.highlight_color = __get_highlight_color(color)


@singledispatch
def __get_rgb_color(color) -> RGBColor:
    """Получает объект RGBColor для покраски текста объекта Run.

    :param color: Наименование цвета или кортеж параметров RGB.
    :return: Объект RGBColor.
    """
    raise NotImplementedError(
        f"Неподдерживаемый тип аргумента: {type(color)}")


@__get_rgb_color.register
def _(color: str) -> RGBColor:
    """Перегрузка функции __get_rgb_color.
    Извлекает RGBColor объект по наименованию цвета из словаря.

    :param color: Наименование цвета.
    :return: Объект RGBColor, соответствующий наименованию цвета.
    """
    try:
        return RGB_COLORS[color]
    except KeyError:
        # TODO: Выводить варнинг если имя не найдено при его наличии (не '')
        return RGB_COLORS[DEFAULT_RGB_COLOR]


@__get_rgb_color.register
def _(color: tuple) -> RGBColor:
    """Перегрузка функции __get_rgb_color.
    Создаёт объект RGBColor по параметрам RGB.

    :param color: Кортеж параметров RGB, например: (255, 0, 0).
    :return: Объект RGBColor, соответствующий параметрам RGB.
    """
    try:
        return RGBColor(*color)
    except TypeError:
        warnings.warn(
            'Параметров цвета должно быть 3! '
            f'Использован цвет по умолчанию ({DEFAULT_RGB_COLOR}).',
            UserWarning,
            stacklevel=4
        )
        return RGB_COLORS[DEFAULT_RGB_COLOR]
    except ValueError:
        warnings.warn(
            'Параметры RGB цвета могут принимать значения от 0 до 255! '
            f'Использован цвет по умолчанию ({DEFAULT_RGB_COLOR}).',
            UserWarning,
            stacklevel=4
        )
        return RGB_COLORS[DEFAULT_RGB_COLOR]


def __get_highlight_color(color_name: str) -> HighlightColor:
    """Извлекает объект цвета HighlightColor по его наименованию.

    :param color_name: Наименование цвета.
    :return: Объект HighlightColor, соответствующий наименованию цвета.
    """
    try:
        return HIGHLIGHT_COLORS_DOCX[color_name]
    except KeyError:
        if color_name:
            warnings.warn(
                'Среди доступных наименований цветов '
                f'отсутствует "{color_name}"! '
                f'Использован цвет по умолчанию ({DEFAULT_RGB_COLOR}).',
                UserWarning,
                stacklevel=3
            )
        return HIGHLIGHT_COLORS_DOCX[DEFAULT_HIGHLIGHT_COLOR]


# Словарь цветов RGB:
RGB_COLORS = MappingProxyType({
    # Basic:
    'red': RGBColor(255, 0, 0),
    'darkred': RGBColor(128, 0, 0),
    'yellow': RGBColor(255, 255, 0),
    'darkyellow': RGBColor(128, 128, 0),
    'lime': RGBColor(0, 255, 0),
    'darkgreen': RGBColor(0, 128, 0),
    'aqua': RGBColor(0, 255, 255),
    'teal': RGBColor(0, 128, 128),
    'blue': RGBColor(0, 0, 255),
    'navy': RGBColor(0, 0, 128),
    'magenta': RGBColor(255, 0, 255),
    'violet': RGBColor(128, 0, 128),
    'black': RGBColor(0, 0, 0),
    'gray': RGBColor(128, 128, 128),
    'lightgray': RGBColor(192, 192, 192),
    'white': RGBColor(255, 255, 255),
    # Alt:
    'maroon': RGBColor(128, 0, 0),  # darkred
    'olive': RGBColor(128, 128, 0),  # darkyellow
    'brightgreen': RGBColor(0, 255, 0),  # lime
    'green': RGBColor(0, 255, 0),  # lime
    'turquoise': RGBColor(0, 255, 255),  # aqua
    'cyan': RGBColor(0, 255, 255),  # aqua
    'darkteal': RGBColor(0, 128, 128),  # teal
    'darkcyan': RGBColor(0, 128, 128),  # teal
    'darkblue': RGBColor(0, 0, 128),  # navy
    'pink': RGBColor(255, 0, 255),  # magenta
    'fuchsia': RGBColor(255, 0, 255),  # magenta
    'purple': RGBColor(255, 0, 255),  # magenta
    'darkmagenta': RGBColor(128, 0, 128),  # violet
    'darkpurple': RGBColor(128, 0, 128),  # violet
    'darkgray': RGBColor(128, 128, 128),  # gray
    'gray50': RGBColor(128, 128, 128),  # gray
    'gray25': RGBColor(192, 192, 192),  # lightgray
    # Extended:
    # 'orange' = RGBColor(255, 165, 0),
    # 'pink' = RGBColor(255, 20, 147),
    # 'coral' = RGBColor(240, 128, 128),
    # 'violet' = RGBColor(138, 43, 226),
    # 'aquamarine' = RGBColor(127, 255, 212),
})
# Словарь цветов для заливки в python-docx:
HIGHLIGHT_COLORS_DOCX = MappingProxyType({
    # Basic:
    'red': HighlightColor(6),
    'darkred': HighlightColor(13),
    'yellow': HighlightColor(7),
    'darkyellow': HighlightColor(14),
    'lime': HighlightColor(4),
    'darkgreen': HighlightColor(11),  # GREEN in WD_COLOR_INDEX
    'aqua': HighlightColor(3),
    'teal': HighlightColor(10),
    'blue': HighlightColor(2),
    'navy': HighlightColor(9),
    'magenta': HighlightColor(5),
    'violet': HighlightColor(12),
    'black': HighlightColor(1),
    'gray': HighlightColor(15),
    'lightgray': HighlightColor(16),
    'white': HighlightColor(8),
    'auto': HighlightColor(0),
    # Alt:
    'maroon': HighlightColor(13),  # darkred
    'olive': HighlightColor(14),  # darkyellow
    'brightgreen': HighlightColor(4),  # lime
    'green': HighlightColor(4),  # lime
    'turquoise': HighlightColor(3),  # aqua
    'cyan': HighlightColor(3),  # aqua
    'darkteal': HighlightColor(10),  # teal
    'darkcyan': HighlightColor(10),  # teal
    'darkblue': HighlightColor(9),  # navy
    'pink': HighlightColor(5),  # magenta
    'fuchsia': HighlightColor(5),  # magenta
    'purple': HighlightColor(5),  # magenta
    'darkmagenta': HighlightColor(12),  # violet
    'darkpurple': HighlightColor(12),  # violet
    'darkgray': HighlightColor(15),  # gray
    'gray50': HighlightColor(15),  # gray
    'gray25': HighlightColor(16),  # lightgray
})
