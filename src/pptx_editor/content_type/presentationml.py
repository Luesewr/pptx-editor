from enum import StrEnum

office_prefix = "application/vnd.openxmlformats-officedocument"
presentationml_prefix = f"{office_prefix}.presentationml"

class PresentationML(StrEnum):
    """PresentationML content types."""
    PRESENTATION = f"{presentationml_prefix}.presentation.main+xml"
    SLIDE = f"{presentationml_prefix}.slide+xml"
    SLIDE_MASTER = f"{presentationml_prefix}.slideMaster+xml"
    SLIDE_LAYOUT = f"{presentationml_prefix}.slideLayout+xml"
    HANDOUT_MASTER = f"{presentationml_prefix}.handoutMaster+xml"
    NOTES_MASTER = f"{presentationml_prefix}.notesMaster+xml"
    NOTES_SLIDE = f"{presentationml_prefix}.notesSlide+xml"
    PRESENTATION_PROPS = f"{presentationml_prefix}.presProps+xml"
    TABLE_STYLES = f"{presentationml_prefix}.tableStyles+xml"
    TAGS = f"{presentationml_prefix}.tags+xml"
    VIEW_PROPS = f"{presentationml_prefix}.viewProps+xml"
    THEME = f"{office_prefix}.theme+xml"
