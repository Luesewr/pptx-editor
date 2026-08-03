from abc import abstractmethod

from pptx_editor.attributes.table import Height, Width
from pptx_editor.properties.xml_element import RequiredXmlElementProperty, XmlElementProperty
from pptx_editor.properties.attribute import RequiredIntegerAttributeProperty
from pptx_editor.properties.table import TableColumnCellProperty, TableRowProperty, TableColumnProperty, TableRowCellProperty
from pptx_editor.xml_element import XmlElement
from pptx_editor.exceptions import PowerpointIntegrityError
from pptx_editor.xml_elements.text import TextBody


class TableCell(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'tc'
    default_order = ('txBody', 'tcPr', 'extLst',)

    text_body = XmlElementProperty(TextBody)

    def __repr__(self) -> str:
        return f"TableCell(text_body={self.text_body.paragraph_texts if self.text_body else '<EMPTY>'})"


class TableSubcollection(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = None
    is_abstract = True

    @abstractmethod
    def is_row(self) -> bool:
        """Determine if this subcollection represents rows or columns."""

    @abstractmethod
    def is_column(self) -> bool:
        """Determine if this subcollection represents rows or columns."""

    def __init_subclass__(cls, **kwargs):
        cls.is_abstract = False
        super().__init_subclass__(**kwargs)

class TableRow(TableSubcollection):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'tr'
    default_order = ('tc', 'extLst',)

    cells = TableRowCellProperty(TableCell)
    height = RequiredIntegerAttributeProperty(Height)

    @property
    def table(self) -> 'Table':
        parent = self.parent
        while parent is not None:
            if isinstance(parent, Table):
                return parent
            parent = parent.parent
        raise PowerpointIntegrityError("TableRow is not part of a Table.")

    def is_row(self) -> bool:
        return True

    def is_column(self) -> bool:
        return False


class GridColumn(TableSubcollection):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'gridCol'
    default_order = ('extLst',)

    cells = TableColumnCellProperty(TableCell)
    width = RequiredIntegerAttributeProperty(Width)

    @property
    def table(self) -> 'Table':
        parent = self.parent
        while parent is not None:
            if isinstance(parent, Table):
                return parent
            parent = parent.parent
        raise PowerpointIntegrityError("GridColumn is not part of a Table.")

    def is_row(self) -> bool:
        return False

    def is_column(self) -> bool:
        return True


class TableGrid(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'tblGrid'
    default_order = ('gridCol',)

    @property
    def grid_columns(self) -> list['GridColumn']:
        return self.get_elements_by_type(GridColumn)


class Table(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'tbl'
    default_order = ('tblPr', 'tblGrid', 'tr',)

    table_grid = RequiredXmlElementProperty(TableGrid)
    rows = TableRowProperty(TableRow)
    columns = TableColumnProperty(GridColumn)
