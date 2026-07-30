import random

from typing import TYPE_CHECKING, Generic, MutableSequence, TypeVar, overload

from pptx_editor.attribute import Attribute
from pptx_editor.attributes.relation_attribute import RelationshipAttribute
from pptx_editor.exceptions import PowerpointIntegrityError
from pptx_editor.relationship import Relationship

if TYPE_CHECKING:
    from pptx_editor.xml_elements.id_list import AbstractIdList

T = TypeVar('T')

class PartIdList(MutableSequence[T]):
    def __init__(self, part_type: type[T], linked_element: 'AbstractIdList'):
        self.element = linked_element
        self.part_type = part_type

    def __len__(self) -> int:
        return len(self.element.children)

    @overload
    def __getitem__(self, index: int) -> T: ...

    @overload
    def __getitem__(self, index: slice) -> list[T]: ...

    def __getitem__(self, index: int | slice) -> T | list[T]:
        if isinstance(index, int):
            id_element = self.element.children[index]
            relationship_attribute = id_element.get_attribute_by_type(RelationshipAttribute)
            return relationship_attribute.value.target

        if isinstance(index, slice):
            return [self[i] for i in range(*index.indices(len(self)))]

        raise TypeError(f"Invalid index type: {type(index).__name__}. Expected int or slice.")

    @overload
    def __setitem__(self, index: int, value: T) -> None: ...

    @overload
    def __setitem__(self, index: slice, value: list[T]) -> None: ...

    def __setitem__(self, index: int | slice, value: T | list[T]) -> None:
        if isinstance(index, int):
            id_element = self.element.children[index]
            relationship_attribute = id_element.get_attribute_by_type(RelationshipAttribute)
            relationship_attribute.value.target = value
        elif isinstance(index, slice):
            if not isinstance(value, list):
                raise TypeError(f"Expected a list for slice assignment, got {type(value).__name__}.")
            for i, v in zip(range(*index.indices(len(self))), value):
                self[i] = v
        else:
            raise TypeError(f"Invalid index type: {type(index).__name__}. Expected int or slice.")

    def __delitem__(self, index: int) -> None:
        element = self.element.children[index]
        relationship = element.get_attribute_by_type(RelationshipAttribute).value
        self.element.remove_element(element)
        self.element.part.remove_relationship(relationship)

    def insert(self, index: int, value: T) -> None:
        relationship_type = self.element.relationship_type
        id_class = self.element.id_class

        if relationship_type is None or id_class is None:
            raise PowerpointIntegrityError(f"Cannot insert into {self.element.__class__.__name__} because it does not have a defined relationship_type or id_class.")

        relationship = Relationship(relationship_type, value, self.element.part)
        relationship_attribute = RelationshipAttribute(relationship, name='id')
        id_attribute = Attribute(str(random.randint(256, 2147483647)), name='id')

        id_element = id_class(attributes=(id_attribute, relationship_attribute,))
        self.element.add_element(id_element, index=index)
        self.element.part.add_relationship(relationship)

    def __repr__(self) -> str:
        return f"PartIdList<{self.part_type.__name__}>{list(self)}"
