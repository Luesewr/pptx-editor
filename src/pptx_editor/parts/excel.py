from pptx_editor.part import Part
from pathlib import PurePosixPath

class ExcelWorksheet(Part):
    default_content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    default_base_path = PurePosixPath('/ppt/embeddings')
    default_part_name = 'Microsoft_Excel_Worksheet'
    default_extension = 'xlsx'
