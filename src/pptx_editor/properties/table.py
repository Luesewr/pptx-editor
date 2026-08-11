import abc

from typing import Sequence, TYPE_CHECKING

from pptx_editor.api_lists.table import TableRowCellList, TableColumnCellList, TableSubcollectionCellList, TableRowList, TableColumnList

if TYPE_CHECKING:
    from pptx_editor.xml_elements.table import TableCell

class TableRowProperty:
    def __init__(self, row_cls):
        self.row_cls = row_cls

    def __get__(self, instance, owner) -> 'TableRowList':
        if instance is None:
            return self
        return TableRowList(instance, self.row_cls)

    def __set__(self, instance, value):
        if not isinstance(value, list):
            raise TypeError(f"Expected a list of TableRow instances, got {type(value).__name__}.")
        row_list = TableRowList(instance, self.row_cls)
        row_list.clear()
        for row in value:
            row_list.append(row)

class TableColumnProperty:
    def __init__(self, column_cls):
        self.column_cls = column_cls

    def __get__(self, instance, owner) -> 'TableColumnList':
        if instance is None:
            return self
        return TableColumnList(instance, self.column_cls)

    def __set__(self, instance, value):
        column_list = TableColumnList(instance, self.column_cls)
        column_list.clear()
        for column in value:
            column_list.append(column)

class TableSubcollectionProperty(abc.ABC):
    @abc.abstractmethod
    def __get__(self, instance, owner) -> 'TableSubcollectionCellList':
        """Return the TableSubcollection instance for this property."""


class TableRowCellProperty(TableSubcollectionProperty):
    def __init__(self, cell_cls):
        self.cell_cls = cell_cls

    def __get__(self, instance, owner) -> 'TableRowCellList':
        if instance is None:
            return self
        return TableRowCellList(instance, self.cell_cls)

    def __set__(self, instance, value: Sequence['TableCell']):
        row_list = TableRowCellList(instance, self.cell_cls)

        if len(value) != len(row_list):
            raise ValueError(f"Expected a list of {len(row_list)} TableCell instances, got {len(value)}.")

        cells = list(value)
        for i, cell in enumerate(cells):
            row_list[i] = cell

class TableColumnCellProperty(TableSubcollectionProperty):
    def __init__(self, cell_cls):
        self.cell_cls = cell_cls

    def __get__(self, instance, owner) -> 'TableColumnCellList':
        if instance is None:
            return self
        return TableColumnCellList(instance, self.cell_cls)

    def __set__(self, instance, value: Sequence['TableCell']):
        column_list = TableColumnCellList(instance, self.cell_cls)

        if len(value) != len(column_list):
            raise ValueError(f"Expected a list of {len(column_list)} TableCell instances, got {len(value)}.")

        cells = list(value)

        for i, cell in enumerate(cells):
            column_list[i] = cell
