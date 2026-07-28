from typing import Generic, TypeVar

from pptx_editor.exceptions import PowerpointIntegrityError
from pptx_editor.part import Part

T = TypeVar('T', bound='Part')

class RelatedPartProperty(Generic[T]):
    def __init__(self, part_type: type[T], target_type: str | None = None, nullable: bool = True):
        self.part_type = part_type
        self.target_type = target_type
        self.nullable = nullable

    def get(self, instance: 'Part', *, nullable_override: bool | None = None) -> T | None:
        related_part = instance.get_related_part_by_type(self.part_type)

        if related_part is not None:
            return related_part

        if not self.nullable and not nullable_override:
            raise PowerpointIntegrityError(f"Expected a related part of type {self.part_type.__name__} in {instance.__class__.__name__}, but none was found.")

        return None

    def set(self, instance: 'Part', value: T | None) -> None:
        existing_related_part = self.get(instance, nullable_override=True)

        if value is None and not self.nullable:
            raise PowerpointIntegrityError(f"Cannot set a non-nullable RelatedPartProperty to None in {instance.__class__.__name__}.")

        if value is not None and value.main_part is not instance.main_part:
            value = value.copy()
            value.update_main_part_recursive(instance.main_part)

        if existing_related_part is not None:
            if value is not None:
                instance.replace_related_part(existing_related_part, value)
            else:
                instance.remove_related_part(existing_related_part)
        elif value is not None:
            instance.add_related_part(value, self.target_type)

    def __get__(self, instance: 'Part | None', owner: type['Part'], nullable_override: bool | None = None) -> T | None:
        if instance is None:
            return self

        return self.get(instance, nullable_override=bool(nullable_override))

    def __set__(self, instance: 'Part', value: T | None) -> None:
        self.set(instance, value)

class RequiredRelatedPartProperty(RelatedPartProperty[T]):
    def __init__(self, part_type: type[T] = None, target_type: str | None = None):
        super().__init__(part_type, target_type=target_type, nullable=False)

    def __get__(self, instance: 'Part | None', owner: type['Part'], nullable_override: bool | None = None) -> T:
        if instance is None:
            return self

        return self.get(instance, nullable_override=bool(nullable_override))

    def __set__(self, instance: 'Part', value: T | None) -> None:
        if value is None:
            raise PowerpointIntegrityError(f"Cannot set a non-nullable RequiredRelatedPartProperty to None in {instance.__class__.__name__}.")

        super().set(instance, value)
