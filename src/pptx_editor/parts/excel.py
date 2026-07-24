from pptx_editor.part import Part
from pathlib import PurePosixPath

class ExcelWorksheet(Part):
    default_content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    default_base_path = PurePosixPath('/ppt/embeddings')
    default_part_name = 'Microsoft_Excel_Worksheet'
    default_extension = 'xlsx'

    def __init__(self, main_part: 'Part', file_path: PurePosixPath | None = None, content_type: str | None = None, is_default: bool = False):
        super().__init__(main_part, file_path, content_type, is_default)
        print("hey")
