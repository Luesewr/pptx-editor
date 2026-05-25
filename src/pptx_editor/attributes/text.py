from typing import TYPE_CHECKING, TypeGuard

from pptx_editor.attribute import Attribute
from pptx_editor.exceptions import PowerpointIntegrityError

class TextBody(Attribute):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_name = 'txBody'

    @property
    def paragraphs(self) -> list['Paragraph']:
        attributes = self.get_attributes('p', 'a')

        if not self._is_paragraphs_valid(attributes):
            raise PowerpointIntegrityError('All paragraph attributes must be of type Paragraph.')

        return attributes

    def _is_paragraphs_valid(self, paragraphs: list[Attribute]) -> TypeGuard[list['Paragraph']]:
        return all(isinstance(paragraph, Paragraph) for paragraph in paragraphs)


class Paragraph(Attribute):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_name = 'p'

    @property
    def runs(self) -> list['Run']:
        attributes = self.get_attributes('r', 'a')

        if not self._is_runs_valid(attributes):
            raise PowerpointIntegrityError('All run attributes must be of type Run.')

        return attributes

    def _is_runs_valid(self, runs: list[Attribute]) -> TypeGuard[list['Run']]:
        return all(isinstance(run, Run) for run in runs)

class Run(Attribute):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_name = 'r'

    @property
    def run_text(self) -> str:
        text_attribute = self.get_attribute('t', 'a')

        if text_attribute is None:
            raise PowerpointIntegrityError('The run element is missing the required t element.')

        if not isinstance(text_attribute, Text):
            raise PowerpointIntegrityError('The t element in the run element is not of the expected type.')
        print(text_attribute)
        return text_attribute.text

    @run_text.setter
    def run_text(self, value: str) -> None:
        text_attribute = self.get_attribute('t', 'a')

        if text_attribute is None:
            raise PowerpointIntegrityError('The run element is missing the required t element.')

        if not isinstance(text_attribute, Text):
            raise PowerpointIntegrityError('The t element in the run element is not of the expected type.')

        text_attribute.text = value

class Text(Attribute):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_name = 't'
