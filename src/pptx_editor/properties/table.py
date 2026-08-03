from pptx_editor.api_lists.table import TableRowCellList, TableColumnCellList, TableRowList, TableColumnList

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
        if not isinstance(value, list):
            raise TypeError(f"Expected a list of GridColumn instances, got {type(value).__name__}.")
        column_list = TableColumnList(instance, self.column_cls)
        column_list.clear()
        for column in value:
            column_list.append(column)

class TableRowCellProperty:
    def __init__(self, cell_cls):
        self.cell_cls = cell_cls

    def __get__(self, instance, owner) -> 'TableRowCellList':
        if instance is None:
            return self
        return TableRowCellList(instance, self.cell_cls)

    def __set__(self, instance, value):
        if not isinstance(value, list):
            raise TypeError(f"Expected a list of TableCell instances, got {type(value).__name__}.")
        row_list = TableRowCellList(instance, self.cell_cls)
        row_list.clear()
        for cell in value:
            row_list.append(cell)

class TableColumnCellProperty:
    def __init__(self, cell_cls):
        self.cell_cls = cell_cls

    def __get__(self, instance, owner) -> 'TableColumnCellList':
        if instance is None:
            return self
        return TableColumnCellList(instance, self.cell_cls)

    def __set__(self, instance, value):
        if not isinstance(value, list):
            raise TypeError(f"Expected a list of TableCell instances, got {type(value).__name__}.")
        column_list = TableColumnCellList(instance, self.cell_cls)
        column_list.clear()
        for cell in value:
            column_list.append(cell)
