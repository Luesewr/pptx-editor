from pptx_editor.parts.masters import NotesMaster, SlideMaster
from pptx_editor.parts.slide import Slide
from pptx_editor.properties.id_list import PartIdListProperty
from pptx_editor.xml_elements.id_list import AbstractId, AbstractIdList


class SlideId(AbstractId):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'sldId'


class SlideIdList(AbstractIdList):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'sldIdLst'
    id_class = SlideId
    relationship_type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'

    slides = PartIdListProperty(Slide)


class SlideMasterId(AbstractId):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'sldMasterId'


class SlideMasterIdList(AbstractIdList):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'sldMasterIdLst'
    id_class = SlideMasterId
    relationship_type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster'

    slide_masters = PartIdListProperty(SlideMaster)


class NotesMasterId(AbstractId):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'notesMasterId'


class NotesMasterIdList(AbstractIdList):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'notesMasterIdLst'
    id_class = NotesMasterId
    relationship_type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster'

    notes_masters = PartIdListProperty(NotesMaster)
