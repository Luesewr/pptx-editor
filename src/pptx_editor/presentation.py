from typing import IO
from zipfile import ZipFile

from pptx_editor.content_types import ContentTypes

class Presentation():
    def __init__(self, file: IO):
        zip_file = ZipFile(file)
        content_types = zip_file.read('[Content_Types].xml')
        parts = ContentTypes(content_types)
