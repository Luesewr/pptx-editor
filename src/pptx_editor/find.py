from enum import Enum
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from pptx_editor.xml_elements.text import ParagraphContent, Paragraph

class StyleInheritMode(Enum):
    FROM_LEFT = 1
    FROM_RIGHT = 2
    FROM_NONE = 3

class MergeMode(Enum):
    LEFT_MERGE = 1
    RIGHT_MERGE = 2
    ISOLATE = 3
    ISOLATE_LEFT = 4
    ISOLATE_RIGHT = 5
    DIVIDE = 6

class CleanupMode(Enum):
    NONE = 1
    DELETE_EMPTY = 2

class ReplaceOptions:
    def __init__(self, style_inherit_mode: 'StyleInheritMode' = StyleInheritMode.FROM_LEFT, merge_mode: 'MergeMode' = MergeMode.LEFT_MERGE, cleanup_mode: 'CleanupMode' = CleanupMode.DELETE_EMPTY):
        self.style_inherit_mode = style_inherit_mode
        self.merge_mode = merge_mode
        self.cleanup_mode = cleanup_mode


class FindResult:
    def __init__(self, result: str, groups: tuple[str, ...], paragraph: 'Paragraph', elements: list['ParagraphContent'], start_offset: int, end_offset: int):
        self.result = result
        self.groups = groups
        self.paragraph = paragraph
        self.elements = elements
        self.start_offset = start_offset
        self.end_offset = end_offset
        self.dependent_results: list['FindResult'] = []

    def replace_with(self, new_text: str, replace_options: 'ReplaceOptions' = ReplaceOptions()) -> None:
        self.replace_with_format(new_text.replace('{', '{{').replace('}', '}}'), replace_options)

    def replace_with_format(self, new_text: str, replace_options: 'ReplaceOptions' = ReplaceOptions()) -> None:
        format_text = new_text.format(*self.groups)

        if not self.elements:
            return

        last_element = self.elements[-1]
        last_element_length = len(last_element.content_text)

        if replace_options.merge_mode == MergeMode.LEFT_MERGE:
            self._replace_with_format_left_merge(format_text)
        elif replace_options.merge_mode == MergeMode.RIGHT_MERGE:
            self._replace_with_format_right_merge(format_text)
        elif replace_options.merge_mode in (MergeMode.ISOLATE, MergeMode.ISOLATE_LEFT, MergeMode.ISOLATE_RIGHT):
            self._replace_with_format_isolate(format_text, replace_options)
        elif replace_options.merge_mode == MergeMode.DIVIDE:
            self._replace_with_format_divide(format_text)

        self._recalculate_elements()

        offset_change = len(last_element.content_text) - last_element_length
        self._shift_dependent_results(offset_change)

        if replace_options.cleanup_mode == CleanupMode.DELETE_EMPTY:
            self._cleanup_empty_elements()

    def _replace_with_format_left_merge(self, new_text: str) -> None:
        if not self.elements:
            return

        first_element = self.elements[0]
        last_element = self.elements[-1]

        if len(self.elements) == 1:
            first_element.content_text = first_element.content_text[:self.start_offset] + new_text + last_element.content_text[self.end_offset:]
            return

        first_element.content_text = first_element.content_text[:self.start_offset] + new_text
        last_element.content_text = last_element.content_text[self.end_offset:]

        for element in self.elements[1:-1]:
            element.content_text = ''

    def _replace_with_format_right_merge(self, new_text: str) -> None:
        if not self.elements:
            return

        first_element = self.elements[0]
        last_element = self.elements[-1]

        if len(self.elements) == 1:
            first_element.content_text = first_element.content_text[:self.start_offset] + new_text + last_element.content_text[self.end_offset:]
            return

        first_element.content_text = first_element.content_text[:self.start_offset]
        last_element.content_text = new_text + last_element.content_text[self.end_offset:]

        for element in self.elements[1:-1]:
            element.content_text = ''

    def _replace_with_format_isolate(self, new_text: str, replace_options: 'ReplaceOptions') -> None:
        if not self.elements:
            return

        first_element = self.elements[0]
        last_element = self.elements[-1]

        if len(self.elements) == 1:
            first_element = first_element.copy()
            self.paragraph.insert_element_before(first_element, last_element)
            self.elements = [first_element, last_element]

        first_element.content_text = first_element.content_text[:self.start_offset]
        last_element.content_text = last_element.content_text[self.end_offset:]

        for element in self.elements[1:-1]:
            element.content_text = ''

        if replace_options.style_inherit_mode == StyleInheritMode.FROM_LEFT:
            new_element = first_element.copy()
        elif replace_options.style_inherit_mode == StyleInheritMode.FROM_RIGHT:
            new_element = last_element.copy()
        elif replace_options.style_inherit_mode == StyleInheritMode.FROM_NONE:
            from pptx_editor.xml_elements.text import Run, Text
            new_text_element = Text('t', 'a', (), (), '')
            new_element = Run('r', 'a', (), (new_text_element,), None)
        else:
            raise ValueError('Invalid style inherit mode.')

        new_element.content_text = new_text

        if replace_options.merge_mode in (MergeMode.ISOLATE, MergeMode.ISOLATE_LEFT):
            self.paragraph.insert_element_after(new_element, first_element)
        if replace_options.merge_mode in (MergeMode.ISOLATE_RIGHT,):
            self.paragraph.insert_element_before(new_element, last_element)

    def _replace_with_format_divide(self, new_text: str) -> None:
        if not self.elements:
            return

        first_element = self.elements[0]
        last_element = self.elements[-1]

        if len(self.elements) == 1:
            first_element.content_text = first_element.content_text[:self.start_offset] + new_text + last_element.content_text[self.end_offset:]
            return

        texts = _split_smooth(new_text, len(self.elements))

        first_element.content_text = first_element.content_text[:self.start_offset] + texts[0]
        last_element.content_text = texts[-1] + last_element.content_text[self.end_offset:]

        for element, text in zip(self.elements[1:-1], texts[1:-1]):
            element.content_text = text

    def _recalculate_elements(self) -> None:
        first_element = self.elements[0]
        last_element = self.elements[-1]

        first_element_index = self.paragraph.paragraph_elements.index(first_element)
        last_element_index = self.paragraph.paragraph_elements.index(last_element)

        self.elements = self.paragraph.paragraph_elements[first_element_index:last_element_index + 1]

    def _shift_dependent_results(self, offset_change: int) -> None:
        for dependent_result in self.dependent_results:
            dependent_result._shift_offsets(offset_change)

    def _cleanup_empty_elements(self) -> None:
        empty_elements = [element for element in self.elements if element.content_text == '']
        self.paragraph.children = [child for child in self.paragraph.children if child not in empty_elements]


    def _shift_offsets(self, offset_change: int) -> None:
        self.start_offset += offset_change
        self.end_offset += offset_change


def _split_smooth(text, n):
    k, m = divmod(len(text), n)
    # Distribute the remainder 'm' by adding 1 extra character to the first 'm' chunks
    return [
        text[i * k + min(i, m) : (i + 1) * k + min(i + 1, m)] for i in range(n)
    ]
