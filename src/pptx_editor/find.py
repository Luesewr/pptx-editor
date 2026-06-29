from enum import Enum
from typing import cast

import pptx_editor.xml_elements.text as text

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


class NewlineMode(Enum):
    REPLACE = 1
    PRESERVE = 2


class CleanupMode(Enum):
    NONE = 1
    DELETE_EMPTY = 2
    DELETE_FULL_EMPTY = 3


class ReplaceOptions:
    def __init__(self, style_inherit_mode: 'StyleInheritMode' = StyleInheritMode.FROM_LEFT, merge_mode: 'MergeMode' = MergeMode.LEFT_MERGE, newline_mode: 'NewlineMode' = NewlineMode.REPLACE, cleanup_mode: 'CleanupMode' = CleanupMode.DELETE_EMPTY):
        self.style_inherit_mode = style_inherit_mode
        self.merge_mode = merge_mode
        self.newline_mode = newline_mode
        self.cleanup_mode = cleanup_mode


class FindResult:
    def __init__(self, result: str, groups: tuple[str, ...], paragraph: 'text.Paragraph', elements: list['text.ParagraphContent'], start_offset: int, end_offset: int):
        self.result = result
        self.groups = groups
        self.paragraph = paragraph
        self.elements = elements
        self._full_elements = elements.copy()
        self.start_offset = start_offset
        self.end_offset = end_offset
        self.dependent_results: list['FindResult'] = []
        self.is_dependent = False
        self._replaced = False

    def replace_with(self, new_text: str, replace_options: 'ReplaceOptions' = ReplaceOptions()) -> None:
        self.replace_with_format(new_text.replace('{', '{{').replace('}', '}}'), replace_options)

    def replace_with_format(self, new_text: str, replace_options: 'ReplaceOptions' = ReplaceOptions()) -> None:
        if self._replaced:
            raise ValueError('This FindResult has already been replaced. You cannot replace it again.')

        format_text = new_text.format(*self.groups)

        if not self.elements:
            return

        last_element_length = self._last_element_length()

        if replace_options.merge_mode == MergeMode.LEFT_MERGE:
            self._replace_with_format_left_merge(format_text)
        elif replace_options.merge_mode == MergeMode.RIGHT_MERGE:
            self._replace_with_format_right_merge(format_text)
        elif replace_options.merge_mode in (MergeMode.ISOLATE, MergeMode.ISOLATE_LEFT, MergeMode.ISOLATE_RIGHT):
            self._replace_with_format_isolate(format_text, replace_options)
        elif replace_options.merge_mode == MergeMode.DIVIDE:
            self._replace_with_format_divide(format_text)

        if replace_options.newline_mode == NewlineMode.REPLACE:
            self._process_newlines()

        offset_change = self._last_element_length() - last_element_length
        self._shift_dependent_results(offset_change)

        if replace_options.cleanup_mode in (CleanupMode.DELETE_EMPTY, CleanupMode.DELETE_FULL_EMPTY):
            self._cleanup_empty_elements(replace_options.cleanup_mode)

        self._replaced = True

    def is_replaced(self) -> bool:
        return self._replaced

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

        self.elements = [first_element]

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

        self.elements = [last_element]

    def _replace_with_format_isolate(self, new_text: str, replace_options: 'ReplaceOptions') -> None:
        if not self.elements:
            return

        first_element = self.elements[0]
        last_element = self.elements[-1]

        if len(self.elements) == 1:
            first_element = first_element.copy()
            self.paragraph.insert_element_before(first_element, last_element)
            self.elements = [first_element, last_element]
            self._full_elements = [first_element, last_element]

        first_element.content_text = first_element.content_text[:self.start_offset]
        last_element.content_text = last_element.content_text[self.end_offset:]

        for element in self.elements[1:-1]:
            element.content_text = ''

        if replace_options.style_inherit_mode == StyleInheritMode.FROM_LEFT:
            new_element = first_element.copy()
        elif replace_options.style_inherit_mode == StyleInheritMode.FROM_RIGHT:
            new_element = last_element.copy()
        elif replace_options.style_inherit_mode == StyleInheritMode.FROM_NONE:
            new_text_element = text.Text(part=self.paragraph.part)
            new_element = cast(text.ParagraphContent, text.Run(children=(new_text_element,), part=self.paragraph.part))
        else:
            raise ValueError('Invalid style inherit mode.')

        new_element.content_text = new_text

        if replace_options.merge_mode in (MergeMode.ISOLATE, MergeMode.ISOLATE_LEFT):
            self.paragraph.insert_element_after(new_element, first_element)
        if replace_options.merge_mode in (MergeMode.ISOLATE_RIGHT,):
            self.paragraph.insert_element_before(new_element, last_element)

        self._full_elements = self._recalculate_elements(self._full_elements[0], self._full_elements[-1])
        self.elements = [new_element]

    def _replace_with_format_divide(self, new_text: str) -> None:
        if not self.elements:
            return

        first_element = self.elements[0]
        last_element = self.elements[-1]

        if len(self.elements) == 1:
            first_element.content_text = first_element.content_text[:self.start_offset] + new_text + last_element.content_text[self.end_offset:]
            return

        smooth_texts = _split_smooth(new_text, len(self.elements))

        first_element.content_text = first_element.content_text[:self.start_offset] + smooth_texts[0]
        last_element.content_text = smooth_texts[-1] + last_element.content_text[self.end_offset:]

        for element, smooth_text in zip(self.elements[1:-1], smooth_texts[1:-1]):
            element.content_text = smooth_text

    def _recalculate_elements(self, first_element: 'text.ParagraphContent', last_element: 'text.ParagraphContent') -> list['text.ParagraphContent']:
        first_element_index = next((i for i, obj in enumerate(self.paragraph.paragraph_elements) if obj is first_element), None)
        last_element_index = next((i for i, obj in enumerate(self.paragraph.paragraph_elements) if obj is last_element), None)

        if first_element_index is None or last_element_index is None:
            raise ValueError('First or last element is not a child of the paragraph.')

        return self.paragraph.paragraph_elements[first_element_index:last_element_index + 1]

    def _process_newlines(self) -> None:
        element_count = len(self.elements)

        first_element = self.elements[0]
        last_element = self.elements[-1]

        shares_last_element = last_element is self._full_elements[-1]

        for element_index, element in enumerate(self.elements):
            if isinstance(element, text.Run) and '\n' in element.content_text:
                parts = element.content_text.split('\n')
                element.content_text = parts[0]

                for part in parts[1:]:
                    break_element = text.Break(part=self.paragraph.part)
                    self.paragraph.insert_element_after(break_element, element)
                    new_element = element.copy()
                    new_element.content_text = part
                    self.paragraph.insert_element_after(new_element, break_element)
                    element = new_element

                if element_index == element_count - 1:
                    last_element = element

        if shares_last_element:
            self._full_elements = self._recalculate_elements(self._full_elements[0], last_element)
        else:
            self._full_elements = self._recalculate_elements(self._full_elements[0], self._full_elements[-1])

        self.elements = self._recalculate_elements(first_element, last_element)


    def _shift_dependent_results(self, offset_change: int) -> None:
        for dependent_result in self.dependent_results:
            dependent_result.shift_offsets(offset_change)

    def _cleanup_empty_elements(self, cleanup_mode: CleanupMode) -> None:
        elements = self.elements if cleanup_mode == CleanupMode.DELETE_EMPTY else self._full_elements

        start_index = 1 if self.is_dependent and elements[0] is self._full_elements[0] else 0
        active_dependent_results = [result for result in self.dependent_results if not result.is_replaced()]
        end_index = len(elements) - 1 if len(active_dependent_results) > 0 and elements[-1] is self._full_elements[-1] else len(elements)
        empty_elements_ids = [id(element) for element in elements[start_index:end_index] if element.content_text == '']

        self.paragraph.children = [child for child in self.paragraph.children if id(child) not in empty_elements_ids]
        self.elements = [element for element in self.elements if id(element) not in empty_elements_ids]
        self._full_elements = [element for element in self._full_elements if id(element) not in empty_elements_ids]

    def shift_offsets(self, offset_change: int) -> None:
        if self._replaced:
            return

        self.start_offset += offset_change

        if len(self._full_elements) == 1:
            self.end_offset += offset_change

    def _last_element_length(self) -> int:
        if not self._full_elements:
            return 0

        last_element = self._full_elements[-1]
        return len(last_element.content_text)


def _split_smooth(text, n):
    k, m = divmod(len(text), n)
    # Distribute the remainder 'm' by adding 1 extra character to the first 'm' chunks
    return [
        text[i * k + min(i, m) : (i + 1) * k + min(i + 1, m)] for i in range(n)
    ]
