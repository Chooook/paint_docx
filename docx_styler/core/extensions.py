from dataclasses import dataclass
from typing import Iterable


@dataclass(frozen=True)
class TextToHighlight:
    """
    Класс TextHighlight предназначен для управления и валидации строк,
    которые необходимо выделить в заданном тексте.
    Этот класс обеспечивает наличие указанных строк в тексте и гарантирует,
    что все поля класса не пустые и имеют корректные типы.

    :param text: Текст, в котором будут выделены указанные строки.
    :param words_to_highlight: Список слов, которые необходимо выделить.
    """
    __slots__ = ['text', 'words_to_highlight']

    text: str
    words_to_highlight: Iterable[str]

    def __post_init__(self):
        """Validate that all fields are non-empty and of the correct types."""
        if not all(getattr(self, slot) for slot in self.__slots__):
            raise ValueError(
                f'Expected all fields of {self.__class__.__name__} '
                'to be non-empty')

        if not all(self.words_to_highlight):
            raise ValueError(
                'Expected all elements of "words_to_highlight" '
                'to be non-empty')

        if not isinstance(self.text, str):
            raise TypeError(
                'Expected "text" to be of type "str", '
                f'got {type(self.text).__name__}')

        if not isinstance(self.words_to_highlight, Iterable):
            raise TypeError(
                'Expected "words_to_highlight" to be of type "Iterable", '
                f'got {type(self.words_to_highlight).__name__}')

        for string in self.words_to_highlight:
            if not string:
                raise ValueError(
                    'Expected all elements of "words_to_highlight"'
                    'to be non-empty')

            if not isinstance(string, str):
                raise TypeError(
                    'Expected all elements of "words_to_highlight" '
                    f'to be of type "str", got {type(string).__name__}')

            if string not in self.text:
                raise ValueError(
                    'Expected all elements of "words_to_highlight" to be '
                    f'in "text" field. Not found "{string}" in "text" field')
