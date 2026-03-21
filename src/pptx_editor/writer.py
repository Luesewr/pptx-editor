from io import BytesIO
from zipfile import ZipFile

class Writer:
    def write_to_buffer(self, presentation):
        buffer = BytesIO()

        with ZipFile(buffer, 'w') as zip_file:
            # content_types =


            for part in presentation.parts:
                zip_file.writestr(part, part.content)

        buffer.seek(0)
        return buffer