from typing import TypeVar, Generic

from pptx_editor.attribute import Attribute
from pptx_editor.xml_element import XmlElement
from pptx_editor.exceptions import PowerpointIntegrityError

T = TypeVar('T', bound='Attribute')

class AttributeProperty(Generic[T]):
    def __init__(self, attribute_type: type[T], nullable: bool = True):
        self.attribute_type = attribute_type
        self.nullable = nullable

    def __get__(self, instance: 'XmlElement | None', owner: type['XmlElement'], nullable_override: bool | None = None) -> T | None:
        if instance is None:
            raise AttributeError("AttributeProperty can only be accessed from an instance.")

        attribute = instance.get_attribute_by_type(self.attribute_type)

        if attribute is not None:
            return attribute

        if not self.nullable and not nullable_override:
            raise PowerpointIntegrityError(f"Expected an attribute of type {self.attribute_type.__name__} in {instance.__class__.__name__}, but none was found.")

        return None

    def __set__(self, instance: 'XmlElement', value: T | None) -> None:
        existing_attribute = self.__get__(instance, type(instance), nullable_override=True)

        if value is None and not self.nullable:
            raise PowerpointIntegrityError(f"Cannot set a non-nullable AttributeProperty to None in {instance.__class__.__name__}.")

        if value is not None and value.part is not instance.part:
            value = value.copy()

        if existing_attribute is not None:
            if value is not None:
                instance.replace_attribute(existing_attribute, value)
            else:
                instance.remove_attribute(existing_attribute)
        elif value is not None:
            instance.add_attribute(value)

class StringAttributeProperty(AttributeProperty[T]):
    def __get__(self, instance: 'XmlElement | None', owner: type['XmlElement'], nullable_override: bool | None = None) -> str | None:
        attribute = super().__get__(instance, owner, nullable_override)
        return attribute.value if attribute is not None else None

    def __set__(self, instance: 'XmlElement', value: str | None) -> None:
        existing_attribute = super().__get__(instance, type(instance), nullable_override=True)

        if value is None and not self.nullable:
            raise PowerpointIntegrityError(f"Cannot set a non-nullable StringAttributeProperty to None in {instance.__class__.__name__}.")

        if existing_attribute is not None:
            if value is not None:
                existing_attribute.value = value
            else:
                instance.remove_attribute(existing_attribute)
        elif value is not None:
            attribute_instance = self.attribute_type(value)
            instance.add_attribute(attribute_instance)

class RequiredStringAttributeProperty(StringAttributeProperty[T]):
    def __init__(self, attribute_type: type[T]):
        super().__init__(attribute_type, nullable=False)

    def __get__(self, instance: 'XmlElement | None', owner: type['XmlElement'], nullable_override: bool | None = None) -> str:
        return super().__get__(instance, owner, nullable_override=bool(nullable_override))

class IntegerAttributeProperty(AttributeProperty[T]):
    def __init__(self, attribute_type: type[T], nullable: bool = True, scalar: int | float = 1):
        super().__init__(attribute_type, nullable)
        self.scalar = scalar

    def __get__(self, instance: 'XmlElement | None', owner: type['XmlElement'], nullable_override: bool | None = None) -> int | None:
        attribute = super().__get__(instance, owner, nullable_override)
        return int(int(attribute.value) / self.scalar) if attribute is not None else None

    def __set__(self, instance: 'XmlElement', value: int | None) -> None:
        existing_attribute = super().__get__(instance, type(instance), nullable_override=True)

        if value is None and not self.nullable:
            raise PowerpointIntegrityError(f"Cannot set a non-nullable IntegerAttributeProperty to None in {instance.__class__.__name__}.")

        if existing_attribute is not None:
            if value is not None:
                existing_attribute.value = str(int(value * self.scalar))
            else:
                instance.remove_attribute(existing_attribute)
        elif value is not None:
            attribute_instance = self.attribute_type(str(int(value * self.scalar)))
            instance.add_attribute(attribute_instance)

class BooleanAttributeProperty(AttributeProperty[T]):
    def __get__(self, instance: 'XmlElement | None', owner: type['XmlElement'], nullable_override: bool | None = None) -> bool | None:
        attribute = super().__get__(instance, owner, nullable_override)
        if attribute is not None:
            return attribute.value == '1'
        return None

    def __set__(self, instance: 'XmlElement', value: bool | None) -> None:
        existing_attribute = super().__get__(instance, type(instance), nullable_override=True)

        if value is None and not self.nullable:
            raise PowerpointIntegrityError(f"Cannot set a non-nullable BooleanAttributeProperty to None in {instance.__class__.__name__}.")

        if existing_attribute is not None:
            if value is not None:
                existing_attribute.value = '1' if value else '0'
            else:
                instance.remove_attribute(existing_attribute)
        elif value is not None:
            attribute_instance = self.attribute_type('1' if value else '0')
            instance.add_attribute(attribute_instance)
