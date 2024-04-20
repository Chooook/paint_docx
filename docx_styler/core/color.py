"""Модуль с функциями для покраски элементов объекта Document."""
# TODO сделать одну общую сущность, от которой можно брать цвет по типу?
#  Например, Color.get('red').fill_ms или Color.get('red').rgb_ms
#  или Color.get('red') -> RGBColor, Color.get('red', 'fill') -> FillColor
#  ну или Color.get('red', fill=True) -> FillColor, иначе RGBColor
#  Или наоборот сделать так, чтобы функция сама брала из кортежа нужный тип:
#  сейчас так и происходит, классы берутся разные, а так интерфейс будет единым
#  Также можно сделать конструктор RGBColor из кортежа чисел от пользователя
from docx.shared import RGBColor
from docx.text.run import Run


def color_run(run: Run, color: str) -> None:
    """Функция для покраски текста объекта Run.

    :param run: Run, который нужно покрасить.
    :param color: Цвет, в который нужно покрасить
    """
    run.font.color.rgb = RgbColor.get(color)


def fill_run_with_color(run: Run, color: str) -> None:
    """Функция для покраски текста объекта Run.

    :param run: Run, который нужно покрасить.
    :param color: Цвет, в который нужно покрасить
    """
    run.font.highlight_color = FillColor.get(color)


# TODO выделить каждый тип цвета в отдельный модуль?
class RgbColor:
    """Класс, содержащий цвета для покраски текста.

    Позволяет извлекать цвета таким образом:
    >>> RgbColor.get('color_name')
    В случае отсутствия цвета, соответствующего переданному имени
    возвращает красный цвет
    """

    # Basic:
    red = RGBColor(255, 0, 0)
    darkred = RGBColor(128, 0, 0)
    yellow = RGBColor(255, 255, 0)
    darkyellow = RGBColor(128, 128, 0)
    lime = RGBColor(0, 255, 0)
    darkgreen = RGBColor(0, 128, 0)
    aqua = RGBColor(0, 255, 255)
    teal = RGBColor(0, 128, 128)
    blue = RGBColor(0, 0, 255)
    navy = RGBColor(0, 0, 128)
    magenta = RGBColor(255, 0, 255)
    violet = RGBColor(128, 0, 128)
    black = RGBColor(0, 0, 0)
    gray = RGBColor(128, 128, 128)
    lightgray = RGBColor(192, 192, 192)
    white = RGBColor(255, 255, 255)
    # Alt:
    maroon = darkred
    olive = darkyellow
    brightgreen = green = lime
    turquoise = cyan = aqua
    darkteal = darkcyan = teal
    darkblue = navy
    pink = fuchsia = purple = magenta
    darkmagenta = darkpurple = violet
    darkgray = gray50 = gray
    gray25 = lightgray
    # Extended:
    # orange = RGBColor(255, 165, 0)
    # pink = RGBColor(255, 20, 147)
    # coral = RGBColor(240, 128, 128)
    # violet = RGBColor(138, 43, 226)
    # aquamarine = RGBColor(127, 255, 212)

    @classmethod
    def __getitem__(cls, item: str) -> RGBColor:
        """Метод для извлечения цвета по его наименованию."""
        try:
            item = ''.join(filter(str.isalpha, item))  # только буквы
            return getattr(cls, item.lower())
        except AttributeError:
            return cls.red

    @classmethod
    def get(cls, item: str):
        """Метод для доступа к __getitem__ без инициализации объекта."""
        return cls.__getitem__(item)


# TODO выделить каждый тип цвета в отдельный модуль?
class FillColor:
    """Класс, содержащий цвета для заливки текста.

    Позволяет извлекать цвета таким образом:
    >>> FillColor.get('color_name')
    В случае отсутствия цвета, соответствующего переданному имени
    возвращает красный цвет
    """
    # Basic:
    red = 6
    darkred = 13
    yellow = 7
    darkyellow = 14
    lime = 4
    darkgreen = 11  # GREEN in WD_COLOR_INDEX
    aqua = 3
    teal = 10
    blue = 2
    navy = 9
    magenta = 5
    violet = 12
    black = 1
    gray = 15
    lightgray = 16
    white = 8
    auto = 0
    # Alt:
    maroon = darkred
    olive = darkyellow
    brightgreen = green = lime
    turquoise = cyan = aqua
    darkteal = darkcyan = teal
    darkblue = navy
    pink = fuchsia = purple = magenta
    darkmagenta = darkpurple = violet
    darkgray = gray50 = gray
    gray25 = lightgray

    @classmethod
    def __getitem__(cls, item: str):
        """Метод для извлечения цвета по его наименованию."""
        try:
            item = ''.join(filter(str.isalpha, item))  # только буквы
            return getattr(cls, item.lower())
        except AttributeError:
            return cls.red

    @classmethod
    def get(cls, item: str):
        """Метод для доступа к __getitem__ без инициализации объекта."""
        return cls.__getitem__(item)
