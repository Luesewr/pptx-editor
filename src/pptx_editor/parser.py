from zipfile import ZipFile
from typing import IO

class Parser:
    def __init__(self, file: IO):
        self.zip_file = ZipFile(file)
        self.parts = {}

    def read_file(self, file_path: str) -> bytes:
        return self.zip_file.read(file_path)

    def add_part(self, part):
        self.parts[part.file_path] = part

    def has_part(self, file_path: str) -> bool:
        return file_path in self.parts
