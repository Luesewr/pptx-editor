from pptx_editor.xml_element import XmlElement
from pptx_editor.exceptions import PowerpointIntegrityError
from pptx_editor.xml_elements.shape import Shape


class CommonSlideData(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'cSld'

    @property
    def shapes(self) -> list['Shape']:
        return self._shape_tree.shapes

    @property
    def _shape_tree(self) -> 'ShapeTree':
        shape_tree = self.get_element('spTree', 'p')

        if shape_tree is None:
            raise PowerpointIntegrityError('The cSld element is missing the required spTree element.')

        if shape_tree is not None and not isinstance(shape_tree, ShapeTree):
            raise PowerpointIntegrityError('The spTree element in the cSld element is not of the expected type.')

        return shape_tree


class ShapeTree(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'spTree'

    @property
    def shapes(self) -> list['Shape']:
        return [attribute for attribute in self.children if isinstance(attribute, Shape)]
