import abc

from collections.abc import MutableSequence
from typing import TYPE_CHECKING, Sequence, overload

from pptx_editor.xml_element import XmlElement

if TYPE_CHECKING:
    from pptx_editor.xml_elements.table import TableRow, GridColumn, Table, TableCell, TableSubcollection

class TableRowList(MutableSequence['TableRow']):
    def __init__(self, table: 'Table', row_cls: type['TableRow']):
        self.table = table
        self.row_cls = row_cls

    def __len__(self) -> int:
        return len(self.table.get_elements_by_type(self.row_cls))

    @overload
    def __getitem__(self, index: int) -> 'TableRow': ...

    @overload
    def __getitem__(self, index: slice) -> list['TableRow']: ...

    def __getitem__(self, index: int | slice) -> 'TableRow | list[TableRow]':
        if isinstance(index, int):
            return self.table.get_elements_by_type(self.row_cls)[index]

        if isinstance(index, slice):
            return [self[i] for i in range(*index.indices(len(self)))]

        raise TypeError(f"Invalid index type: {type(index).__name__}. Expected int or slice.")

    @overload
    def __setitem__(self, index: int, value: 'TableSubcollection | Sequence[TableCell] | Sequence[TableSubcollection | Sequence[TableCell]]') -> None: ...

    @overload
    def __setitem__(self, index: slice, value: 'Sequence[TableSubcollection | Sequence[TableCell]]') -> None: ...

    def __setitem__(self, index: int | slice, value: 'TableSubcollection | Sequence[TableCell] | Sequence[TableSubcollection | Sequence[TableCell]]') -> None:
        if isinstance(index, int):
            if isinstance(value, XmlElement) and value.is_row():
                row = value.copy() if value.table is self.table else value
            else:
                row = self.row_cls()
                row.height = 0
                row.cells.extend([cell.copy() for cell in value.cells] if isinstance(value, XmlElement) else value)

            self.table.replace_element(self[index], row)
        elif isinstance(index, slice):
            if not isinstance(value, Sequence):
                raise TypeError(f"Expected a sequence for slice assignment, got {type(value).__name__}.")
            for i, v in zip(range(*index.indices(len(self))), value):
                self[i] = v
        else:
            raise TypeError(f"Invalid index type: {type(index).__name__}. Expected int or slice.")

    @overload
    def __delitem__(self, index: int) -> None: ...

    @overload
    def __delitem__(self, index: slice) -> None: ...

    def __delitem__(self, index: int | slice) -> None:
        if isinstance(index, int):
            rows = [self[index]]
        elif isinstance(index, slice):
            rows = [self[i] for i in range(*index.indices(len(self)))]
        else:
            raise TypeError(f"Invalid index type: {type(index).__name__}. Expected int or slice.")

        for row in rows:
            self.table.remove_element(row)

    def insert(self, index: int, value: 'TableSubcollection | Sequence[TableCell]') -> None:
        if isinstance(value, XmlElement) and value.is_row():
            row = value.copy() if value.table is self.table else value
        else:
            row = self.row_cls()
            row.height = 0
            row.cells.extend([cell.copy() for cell in value.cells] if isinstance(value, XmlElement) else value)

        reference_row = self[index - 1] if index - 1 >= 0 and index - 1 < len(self) else None

        if reference_row is not None:
            self.table.insert_element_after(row, reference_row)
        else:
            self.table.auto_add_element(row)

    def __repr__(self) -> str:
        return f"TableRowList{list(self)}"


class TableColumnList(MutableSequence['GridColumn']):
    def __init__(self, table: 'Table', column_cls: type['GridColumn']):
        self.table = table
        self.column_cls = column_cls

    def __len__(self) -> int:
        return len(self.table.table_grid.grid_columns)

    @overload
    def __getitem__(self, index: int) -> 'GridColumn': ...

    @overload
    def __getitem__(self, index: slice) -> list['GridColumn']: ...

    def __getitem__(self, index: int | slice) -> 'GridColumn | list[GridColumn]':
        if isinstance(index, int):
            return self.table.table_grid.grid_columns[index]

        if isinstance(index, slice):
            return [self[i] for i in range(*index.indices(len(self)))]

        raise TypeError(f"Invalid index type: {type(index).__name__}. Expected int or slice.")

    @overload
    def __setitem__(self, index: int, value: 'TableSubcollection | Sequence[TableCell]') -> None: ...

    @overload
    def __setitem__(self, index: slice, value: 'Sequence[TableSubcollection | Sequence[TableCell]]') -> None: ...

    def __setitem__(self, index: int | slice, value: 'TableSubcollection | Sequence[TableCell] | Sequence[TableSubcollection | Sequence[TableCell]]') -> None:
        if isinstance(index, int):
            if isinstance(value, XmlElement) and value.is_column():
                column = value.copy() if value.table is self.table else value
                cells = list(value.cells)
            else:
                column = self.column_cls()
                column.width = 0
                cells = [cell.copy() for cell in value.cells] if isinstance(value, XmlElement) else value

            self.table.table_grid.replace_element(self[index], column)

            for row, cell in zip(self.table.rows, cells):
                row.cells[index] = cell

        elif isinstance(index, slice):
            if not isinstance(value, Sequence):
                raise TypeError(f"Expected a list for slice assignment, got {type(value).__name__}.")
            for i, v in zip(range(*index.indices(len(self))), value):
                self[i] = v
        else:
            raise TypeError(f"Invalid index type: {type(index).__name__}. Expected int or slice.")

    @overload
    def __delitem__(self, index: int) -> None: ...

    @overload
    def __delitem__(self, index: slice) -> None: ...

    def __delitem__(self, index: int | slice) -> None:
        if isinstance(index, int):
            columns = [self[index]]
        elif isinstance(index, slice):
            columns = [self[i] for i in range(*index.indices(len(self)))]
        else:
            raise TypeError(f"Invalid index type: {type(index).__name__}. Expected int or slice.")

        for column in columns:
            for row in self.table.rows:
                del row.cells[self.table.table_grid.grid_columns.index(column)]
            self.table.table_grid.remove_element(column)

    def insert(self, index: int, value: 'TableSubcollection | Sequence[TableCell]') -> None:
        if isinstance(value, XmlElement) and value.is_column():
            column = value.copy() if value.table is self.table else value
        else:
            column = self.column_cls()
            column.width = 0

        if isinstance(value, XmlElement):
            cells = list(value.cells)
        else:
            cells = value

        for row, cell in zip(self.table.rows, cells):
            row.cells.insert(index, cell)

        reference_column = self[index - 1] if index - 1 >= 0 and index - 1 < len(self) else None

        if reference_column is not None:
            self.table.table_grid.insert_element_after(column, reference_column)
        else:
            self.table.table_grid.auto_add_element(column)

    def __repr__(self) -> str:
        return f"TableColumnList{list(self)}"

class TableSubcollectionCellList(MutableSequence['TableCell'], abc.ABC):
    @abc.abstractmethod
    def __init__(self, subcollection: 'TableSubcollection', cell_cls: type['TableCell']): ...

    @abc.abstractmethod
    def __getitem__(self, index: int | slice) -> 'TableCell | list[TableCell]': ...

    @abc.abstractmethod
    def __setitem__(self, index: int | slice, value: 'TableCell | Sequence[TableCell]') -> None: ...

    @abc.abstractmethod
    def __delitem__(self, index: int | slice) -> None: ...


class TableRowCellList(TableSubcollectionCellList):
    def __init__(self, row: 'TableRow', cell_cls: type['TableCell']):
        self.row = row
        self.cell_cls = cell_cls

    def __len__(self) -> int:
        return len(self.row.get_elements_by_type(self.cell_cls))

    @overload
    def __getitem__(self, index: int) -> 'TableCell': ...

    @overload
    def __getitem__(self, index: slice) -> list['TableCell']: ...

    def __getitem__(self, index: int | slice) -> 'TableCell | Sequence[TableCell]':
        if isinstance(index, int):
            return self.row.get_elements_by_type(self.cell_cls)[index]

        if isinstance(index, slice):
            return [self[i] for i in range(*index.indices(len(self)))]

        raise TypeError(f"Invalid index type: {type(index).__name__}. Expected int or slice.")

    def __setitem__(self, index: int | slice, value: 'TableCell | Sequence[TableCell]') -> None:
        if isinstance(index, int):
            self.row.replace_element(self[index], value)
        elif isinstance(index, slice):
            if not isinstance(value, Sequence):
                raise TypeError(f"Expected a sequence for slice assignment, got {type(value).__name__}.")
            for i, v in zip(range(*index.indices(len(self))), value):
                self[i] = v
        else:
            raise TypeError(f"Invalid index type: {type(index).__name__}. Expected int or slice.")

    @overload
    def __delitem__(self, index: int) -> None: ...

    @overload
    def __delitem__(self, index: slice) -> None: ...

    def __delitem__(self, index: int | slice) -> None:
        if isinstance(index, int):
            cells = [self[index]]
        elif isinstance(index, slice):
            cells = [self[i] for i in range(*index.indices(len(self)))]
        else:
            raise TypeError(f"Invalid index type: {type(index).__name__}. Expected int or slice.")

        for cell in cells:
            self.row.remove_element(cell)

    def insert(self, index: int, value: 'TableCell') -> None:
        reference_cell = self[index - 1] if index - 1 >= 0 and index - 1 < len(self) else None

        if reference_cell is not None:
            self.row.insert_element_after(value, reference_cell)
        else:
            self.row.auto_add_element(value)

    def __repr__(self) -> str:
        return f"TableRowCellList{list(self)}"


class TableColumnCellList(TableSubcollectionCellList):
    def __init__(self, column: 'GridColumn', cell_cls: type['TableCell']):
        self.column = column
        self.cell_cls = cell_cls

    def __len__(self) -> int:
        if self.column.table is None:
            raise ValueError("The column is not associated with a table (Likely a copied column).")

        return len(self.column.table.rows)

    @overload
    def __getitem__(self, index: int) -> 'TableCell': ...

    @overload
    def __getitem__(self, index: slice) -> list['TableCell']: ...

    def __getitem__(self, index: int | slice) -> 'TableCell | Sequence[TableCell]':
        if isinstance(index, int):
            if self.column.table is None:
                raise ValueError("The column is not associated with a table (Likely a copied column).")

            column_index = self.column.table.columns.index(self.column)
            return self.column.table.rows[index].cells[column_index]

        if isinstance(index, slice):
            return [self[i] for i in range(*index.indices(len(self)))]

        raise TypeError(f"Invalid index type: {type(index).__name__}. Expected int or slice.")

    @overload
    def __setitem__(self, index: int, value: 'TableCell') -> None: ...

    @overload
    def __setitem__(self, index: slice, value: Sequence['TableCell']) -> None: ...

    def __setitem__(self, index: int | slice, value: 'TableCell | Sequence[TableCell]') -> None:
        if isinstance(index, int):
            if self.column.table is None:
                raise ValueError("The column is not associated with a table (Likely a copied column).")

            column_index = self.column.table.columns.index(self.column)
            self.column.table.rows[index].cells[column_index] = value
        elif isinstance(index, slice):
            if not isinstance(value, Sequence):
                raise TypeError(f"Expected a sequence for slice assignment, got {type(value).__name__}.")
            for i, v in zip(range(*index.indices(len(self))), value):
                self[i] = v
        else:
            raise TypeError(f"Invalid index type: {type(index).__name__}. Expected int or slice.")

    @overload
    def __delitem__(self, index: int) -> None: ...

    @overload
    def __delitem__(self, index: slice) -> None: ...

    def __delitem__(self, index: int | slice) -> None:
        if isinstance(index, int):
            cells = [self[index]]
        elif isinstance(index, slice):
            cells = [self[i] for i in range(*index.indices(len(self)))]
        else:
            raise TypeError(f"Invalid index type: {type(index).__name__}. Expected int or slice.")

        for cell in cells:
            row_index = self.index(cell)
            column_index = self.column.table.columns.index(self.column)
            self.column.table.rows[row_index].cells.pop(column_index)

    def insert(self, index: int, value: 'TableCell') -> None:
        column_index = self.column.table.columns.index(self.column)
        self.column.table.rows[index].cells.insert(column_index, value)

    def __repr__(self) -> str:
        return f"TableColumnCellList{list(self)}"
