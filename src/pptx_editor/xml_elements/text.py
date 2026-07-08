import re

from abc import ABC, abstractmethod
from typing import TypeGuard

from pptx_editor.attribute import BooleanAttributeProperty, IntegerAttributeProperty, StringAttributeProperty
from pptx_editor.attributes.text import Bold, Italic, Language, Underline, Strikethrough, FontSize, Dirty, Error
from pptx_editor.find import FindResult
from pptx_editor.exceptions import PowerpointIntegrityError
from pptx_editor.xml_element import XmlElement, XmlElementProperty
from pptx_editor.xml_elements.color import AbstractColor
from pptx_editor.xml_elements.fill import AbstractFill
from pptx_editor.xml_elements.font import LatinFont, ComplexScriptFont, EastAsianFont, SymbolFont

class TextBody(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'txBody'

    @property
    def paragraphs(self) -> list['Paragraph']:
        attributes = self.get_elements_by_type(Paragraph)

        if not self._is_paragraphs_valid(attributes):
            raise PowerpointIntegrityError('All paragraph attributes must be of type Paragraph.')

        return attributes

    @property
    def paragraph_texts(self) -> str:
        return '\n\n'.join(paragraph.paragraph_text for paragraph in self.paragraphs)

    def _is_paragraphs_valid(self, paragraphs: list[XmlElement]) -> TypeGuard[list['Paragraph']]:
        return all(isinstance(paragraph, Paragraph) for paragraph in paragraphs)


class Paragraph(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'p'

    @property
    def runs(self) -> list['Run']:
        return [element for element in self.children if isinstance(element, Run)]

    @property
    def paragraph_elements(self) -> list['AbstractParagraphContent']:
        return [element for element in self.children if isinstance(element, AbstractParagraphContent)]

    @property
    def paragraph_text(self) -> str:
        return ''.join(element.content_text for element in self.paragraph_elements)

    def find_in_text(self, text: str) -> list['FindResult']:
        return self.find_regex_in_text(re.escape(text))

    def find_regex_in_text(self, pattern: str | re.Pattern) -> list['FindResult']:
        if isinstance(pattern, str):
            pattern = re.compile(pattern)

        paragraph_text = self.paragraph_text
        paragraph_elements = self.paragraph_elements
        element_index = 0
        current_offset = 0

        dependent_matches = []
        results = []

        for match in pattern.finditer(paragraph_text):
            start_offset = match.start()
            end_offset = match.end()
            match_elements = []

            while element_index < len(paragraph_elements) and current_offset + len(paragraph_elements[element_index].content_text) <= start_offset:
                current_offset += len(paragraph_elements[element_index].content_text)
                dependent_matches = []
                element_index += 1

            match_start_offset = start_offset - current_offset

            match_elements.append(paragraph_elements[element_index])

            while element_index < len(paragraph_elements) and current_offset + len(paragraph_elements[element_index].content_text) < end_offset:
                current_offset += len(paragraph_elements[element_index].content_text)
                dependent_matches = []
                element_index += 1

                if element_index < len(paragraph_elements):
                    match_elements.append(paragraph_elements[element_index])

            match_end_offset = end_offset - current_offset

            find_result = FindResult(match.group(0), (match.group(0), *match.groups()), self, match_elements, match_start_offset, match_end_offset)
            results.append(find_result)

            for dependent_match in dependent_matches:
                dependent_match.dependent_results.append(find_result)
                find_result.is_dependent = True

            dependent_matches.append(find_result)

        return results

class Highlight(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'highlight'

    color: AbstractColor = XmlElementProperty(AbstractColor, nullable=False)

class RunProperties(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'rPr'

    fill: AbstractFill | None = XmlElementProperty(AbstractFill)
    bold: bool | None = BooleanAttributeProperty(Bold)
    italic: bool | None = BooleanAttributeProperty(Italic)
    underline: bool | None = BooleanAttributeProperty(Underline)
    strikethrough: bool | None = BooleanAttributeProperty(Strikethrough)
    font_size: int | None = IntegerAttributeProperty(FontSize, scalar=100)
    language: str | None = StringAttributeProperty(Language)
    dirty: bool | None = BooleanAttributeProperty(Dirty)
    error: bool | None = BooleanAttributeProperty(Error)
    highlight: Highlight | None = XmlElementProperty(Highlight)
    latin_font: LatinFont | None = XmlElementProperty(LatinFont)
    complex_script_font: ComplexScriptFont | None = XmlElementProperty(ComplexScriptFont)
    east_asian_font: EastAsianFont | None = XmlElementProperty(EastAsianFont)
    symbol_font: SymbolFont | None = XmlElementProperty(SymbolFont)


class Text(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 't'


class AbstractParagraphContent(XmlElement, ABC):
    is_abstract = True

    properties: RunProperties | None = XmlElementProperty(RunProperties)

    @property
    @abstractmethod
    def content_text(self) -> str:
        pass

    def __init_subclass__(cls, **kwargs):
        cls.is_abstract = False
        super().__init_subclass__(**kwargs)


class Run(AbstractParagraphContent):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'r'

    _content_text: Text = XmlElementProperty(Text, nullable=False)

    @property
    def content_text(self) -> str:
        return self._content_text.text if self._content_text is not None else ''

    @content_text.setter
    def content_text(self, value: str) -> None:
        self._content_text.text = value


class Break(AbstractParagraphContent):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'br'

    @property
    def content_text(self) -> str:
        return '\n'

class TextField(AbstractParagraphContent):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'fld'

    _content_text: Text = XmlElementProperty(Text, nullable=False)

    @property
    def content_text(self) -> str:
        return self._content_text.text if self._content_text is not None else ''

    @content_text.setter
    def content_text(self, value: str) -> None:
        self._content_text.text = value
