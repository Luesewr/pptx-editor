from pptx_editor.xml_element import XmlElement
from pptx_editor.properties.xml_element import XmlElementProperty
from pptx_editor.xml_elements.text import TextBody

class TableCell(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'tc'
    default_order = ('txBody', 'tcPr', 'extLst',)

    text_body = XmlElementProperty(TextBody)


class TableRow(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'tr'
    default_order = ('tc', 'extLst',)

    @property
    def cells(self) -> list['TableCell']:
        return self.get_elements_by_type(TableCell)


class Table(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'tbl'
    default_order = ('tblPr', 'tblGrid', 'tr',)

    @property
    def rows(self) -> list['TableRow']:
        return self.get_elements_by_type(TableRow)
