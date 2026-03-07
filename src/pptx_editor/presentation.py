from typing import IO
from zipfile import ZipFile

from pptx_editor.content_types import ContentTypes

class Presentation():
    def __init__(self, file: IO):
        zip_file = ZipFile(file)
        content_types_file = zip_file.read('[Content_Types].xml')
        content_types = ContentTypes(content_types_file)
        content_types._parse_content_data(zip_file)
