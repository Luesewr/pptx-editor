from typing import TYPE_CHECKING, TypeGuard

from pptx_editor.attribute import Attribute
from pptx_editor.exceptions import PowerpointIntegrityError

if TYPE_CHECKING:
    from pptx_editor.attributes.shape import Shape

class CommonSlideData(Attribute):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_name = 'cSld'

    def shapes(self) -> list['Shape']:
        return self._shape_tree().shapes()

    def _shape_tree(self) -> 'ShapeTree':
        shape_tree = self.get_attribute('spTree', 'p')

        if shape_tree is None:
            raise PowerpointIntegrityError('The cSld element is missing the required spTree element.')

        if shape_tree is not None and not isinstance(shape_tree, ShapeTree):
            raise PowerpointIntegrityError('The spTree element in the cSld element is not of the expected type.')

        return shape_tree

class ShapeTree(Attribute):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_name = 'spTree'

    def shapes(self) -> list['Shape']:
        from pptx_editor.attributes.shape import Shape

        return [attribute for attribute in self.attributes if isinstance(attribute, Shape)]
