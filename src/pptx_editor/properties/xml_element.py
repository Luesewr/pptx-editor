from typing import TypeVar, Generic

from pptx_editor.xml_element import XmlElement
from pptx_editor.modes.xml_element import AddMode
from pptx_editor.xml_elements.null import NullElement
from pptx_editor.exceptions import PowerpointIntegrityError

T = TypeVar('T', bound='XmlElement')

class XmlElementProperty(Generic[T]):
    def __init__(self, element_type: type[T], nullable: bool = True):
        self.element_type = element_type
        self.nullable = nullable

    def get(self, instance: 'XmlElement', *, nullable_override: bool | None = None) -> T | NullElement[T]:
        child_element = instance.get_element_by_type(self.element_type)

        if child_element is not None:
            return child_element

        if not self.nullable and not nullable_override:
            raise PowerpointIntegrityError(f"Expected a child of type {self.element_type.__name__} in {instance.__class__.__name__}, but none was found.")

        return NullElement(self.element_type, parent=instance)

    def set(self, instance: 'XmlElement', value: T | None, add_mode: AddMode = AddMode.SORT) -> None:
        existing_element = self.get(instance, nullable_override=True)

        if value is None and not self.nullable:
            raise PowerpointIntegrityError(f"Cannot set a non-nullable XmlElementProperty to None in {instance.__class__.__name__}.")

        if value is not None and value.part is not instance.part:
            value = value.copy()
            value.update_part_recursive(instance.part)

        if not isinstance(existing_element, NullElement):
            if value is not None:
                instance.auto_replace_element(existing_element, value, add_mode=add_mode)
            else:
                instance.remove_element(existing_element)
        elif value is not None:
            instance.auto_add_element(value, add_mode=add_mode)

    def __get__(self, instance: 'XmlElement | None', owner: type['XmlElement'], nullable_override: bool | None = None) -> T | NullElement[T]:
        return self.get(instance, nullable_override=nullable_override)

    def __set__(self, instance: 'XmlElement', value: T | None, add_mode: AddMode = AddMode.SORT) -> None:
        self.set(instance, value, add_mode=add_mode)

class RequiredXmlElementProperty(XmlElementProperty[T]):
    def __init__(self, element_type: type[T]):
        super().__init__(element_type, nullable=False)

    def __get__(self, instance: 'XmlElement | None', owner: type['XmlElement'], nullable_override: bool | None = None) -> T:
        if instance is None:
            return self

        return self.get(instance, nullable_override=bool(nullable_override))

    def __set__(self, instance: 'XmlElement', value: T | None, add_mode: AddMode = AddMode.SORT) -> None:
        if value is None:
            raise PowerpointIntegrityError(f"Cannot set a non-nullable RequiredXmlElementProperty to None in {instance.__class__.__name__}.")

        self.set(instance, value, add_mode=add_mode)
