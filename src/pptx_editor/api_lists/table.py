from collections.abc import MutableSequence
from typing import TYPE_CHECKING, overload

if TYPE_CHECKING:
    from pptx_editor.xml_elements.table import TableRow, GridColumn, Table, TableCell

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
    def __setitem__(self, index: int, value: 'TableRow | list[TableCell]') -> None: ...

    @overload
    def __setitem__(self, index: slice, value: 'list[TableRow | list[TableCell]]') -> None: ...

    def __setitem__(self, index: int | slice, value: 'TableRow | list[TableCell] | list[TableRow | list[TableCell]]') -> None:
        if isinstance(index, int):
            if not isinstance(value, list):
                row_copy = value.copy()
                self.table.replace_element(self[index], row_copy)
            else:
                self[index].cells = value
        elif isinstance(index, slice):
            if not isinstance(value, list):
                raise TypeError(f"Expected a list for slice assignment, got {type(value).__name__}.")
            for i, v in zip(range(*index.indices(len(self))), value):
                self[i] = v
        else:
            raise TypeError(f"Invalid index type: {type(index).__name__}. Expected int or slice.")

    def __delitem__(self, index: int) -> None:
        row = self[index]
        self.table.remove_element(row)

    def insert(self, index: int, value: 'TableRow | GridColumn | list[TableCell]') -> None:
        if isinstance(value, list) or value.is_column():
            row = self.row_cls()
            row.height = 0
            row.cells.extend(value if isinstance(value, list) else [cell.copy() for cell in value.cells])
        else:
            row = value.copy()

        reference_row = self[index] if index < len(self) else None

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
    def __setitem__(self, index: int, value: 'GridColumn | list[TableCell]') -> None: ...

    @overload
    def __setitem__(self, index: slice, value: 'list[GridColumn | list[TableCell]]') -> None: ...

    def __setitem__(self, index: int | slice, value: 'GridColumn | list[TableCell] | list[GridColumn | list[TableCell]]') -> None:
        if isinstance(index, int):
            if not isinstance(value, list):
                cells = value.cells
                self.table.table_grid.replace_element(self[index], value)
            else:
                cells = value

            for row, cell in zip(self.table.rows, cells):
                cell_copy = cell.copy()
                row.cells[index] = cell_copy
        elif isinstance(index, slice):
            if not isinstance(value, list):
                raise TypeError(f"Expected a list for slice assignment, got {type(value).__name__}.")
            for i, v in zip(range(*index.indices(len(self))), value):
                self[i] = v
        else:
            raise TypeError(f"Invalid index type: {type(index).__name__}. Expected int or slice.")

    def __delitem__(self, index: int) -> None:
        column = self[index]

        for row in self.table.rows:
            del row.cells[index]

        self.table.table_grid.remove_element(column)

    def insert(self, index: int, value: 'GridColumn | TableRow | list[TableCell]') -> None:
        if isinstance(value, list) or value.is_row():
            column = self.column_cls()
            column.width = 0
        else:
            column = value

        if not isinstance(value, list):
            cells = value.cells
        else:
            cells = value

        for row, cell in zip(self.table.rows, cells):
            row.cells.insert(index, cell)

        reference_column = self[index] if index < len(self) else None

        if reference_column is not None:
            self.table.table_grid.insert_element_after(column, reference_column)
        else:
            self.table.table_grid.auto_add_element(column)

    def __repr__(self) -> str:
        return f"TableColumnList{list(self)}"


class TableRowCellList(MutableSequence['TableCell']):
    def __init__(self, row: 'TableRow', cell_cls: type['TableCell']):
        self.row = row
        self.cell_cls = cell_cls

    def __len__(self) -> int:
        return len(self.row.get_elements_by_type(self.cell_cls))

    @overload
    def __getitem__(self, index: int) -> 'TableCell': ...

    @overload
    def __getitem__(self, index: slice) -> list['TableCell']: ...

    def __getitem__(self, index: int | slice) -> 'TableCell | list[TableCell]':
        if isinstance(index, int):
            return self.row.get_elements_by_type(self.cell_cls)[index]

        if isinstance(index, slice):
            return [self[i] for i in range(*index.indices(len(self)))]

        raise TypeError(f"Invalid index type: {type(index).__name__}. Expected int or slice.")

    @overload
    def __setitem__(self, index: int, value: 'TableCell') -> None: ...

    @overload
    def __setitem__(self, index: slice, value: list['TableCell']) -> None: ...

    def __setitem__(self, index: int | slice, value: 'TableCell | list[TableCell]') -> None:
        if isinstance(index, int):
            self.row.replace_element(self[index], value)
        elif isinstance(index, slice):
            if not isinstance(value, list):
                raise TypeError(f"Expected a list for slice assignment, got {type(value).__name__}.")
            for i, v in zip(range(*index.indices(len(self))), value):
                self[i] = v
        else:
            raise TypeError(f"Invalid index type: {type(index).__name__}. Expected int or slice.")

    def __delitem__(self, index: int) -> None:
        cell = self[index]
        self.row.remove_element(cell)

    def insert(self, index: int, value: 'TableCell') -> None:
        reference_cell = self[index] if index < len(self) else None

        if reference_cell is not None:
            self.row.insert_element_after(value, reference_cell)
        else:
            self.row.auto_add_element(value)

    def __repr__(self) -> str:
        return f"TableRowCellList{list(self)}"


class TableColumnCellList(MutableSequence['TableCell']):
    def __init__(self, column: 'GridColumn', cell_cls: type['TableCell']):
        self.column = column
        self.cell_cls = cell_cls

    def __len__(self) -> int:
        return len(self.column.table.rows)

    @overload
    def __getitem__(self, index: int) -> 'TableCell': ...

    @overload
    def __getitem__(self, index: slice) -> list['TableCell']: ...

    def __getitem__(self, index: int | slice) -> 'TableCell | list[TableCell]':
        if isinstance(index, int):
            column_index = self.column.table.table_grid.grid_columns.index(self.column)
            return self.column.table.rows[index].cells[column_index]

        if isinstance(index, slice):
            return [self[i] for i in range(*index.indices(len(self)))]

        raise TypeError(f"Invalid index type: {type(index).__name__}. Expected int or slice.")

    @overload
    def __setitem__(self, index: int, value: 'TableCell') -> None: ...

    @overload
    def __setitem__(self, index: slice, value: list['TableCell']) -> None: ...

    def __setitem__(self, index: int | slice, value: 'TableCell | list[TableCell]') -> None:
        if isinstance(index, int):
            column_index = self.column.table.table_grid.grid_columns.index(self.column)
            self.column.table.rows[index].cells[column_index] = value
        elif isinstance(index, slice):
            if not isinstance(value, list):
                raise TypeError(f"Expected a list for slice assignment, got {type(value).__name__}.")
            for i, v in zip(range(*index.indices(len(self))), value):
                self[i] = v
        else:
            raise TypeError(f"Invalid index type: {type(index).__name__}. Expected int or slice.")

    def __delitem__(self, index: int) -> None:
        column_index = self.column.table.table_grid.grid_columns.index(self.column)
        del self.column.table.rows[index].cells[column_index]

    def insert(self, index: int, value: 'TableCell') -> None:
        column_index = self.column.table.table_grid.grid_columns.index(self.column)
        self.column.table.rows[index].cells.insert(column_index, value)

    def __repr__(self) -> str:
        return f"TableColumnCellList{list(self)}"
