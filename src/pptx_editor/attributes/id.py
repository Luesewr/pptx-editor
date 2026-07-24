import uuid

from pptx_editor.attribute import Attribute
from pptx_editor.parts.xml_part import XmlPart


class CreationId(Attribute):
    default_namespace = None
    default_prefix = None
    default_name = 'id'
    default_element_names = ['creationId']

    def copy(self, part: 'XmlPart | None' = None, relationship_map: dict | None = None) -> 'CreationId':
        new_id = f'{{{str(uuid.uuid4()).upper()}}}'
        return self.__class__(new_id, self.prefix, self.name, overwrite_prefix=True)
