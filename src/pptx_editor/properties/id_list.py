from typing import TypeVar, TYPE_CHECKING, Generic

from pptx_editor.part_id_list import PartIdList

if TYPE_CHECKING:
    from pptx_editor.xml_elements.id_list import AbstractIdList

T = TypeVar('T', bound='AbstractIdList')

class PartIdListProperty(Generic[T]):
    def __init__(self, part_type: type[T]):
        self.part_type = part_type

    def __get__(self, instance: 'AbstractIdList | None', owner: type['AbstractIdList']) -> 'PartIdList[T]':
        if instance is None:
            return self

        return PartIdList(self.part_type, instance)

    def __set__(self, instance: 'AbstractIdList', value: 'PartIdList[T] | list[T]') -> None:
        part_id_list = PartIdList(self.part_type, instance)
        part_id_list.clear()
        for item in value:
            part_id_list.append(item)
