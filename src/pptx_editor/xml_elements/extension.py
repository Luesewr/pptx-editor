from pptx_editor.xml_element import XmlElement


class ExtensionList(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'extLst'


class Extension(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'ext'
