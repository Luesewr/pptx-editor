from pptx_editor.part import Part

class Base(Part):
    default_content_type = None
    default_base_path = ''
    default_part_name = None

    def __init__(self, base: 'Base | None', file_path: str | None = None, content_type: str | None = None):
        self.parts: list['Part'] = []
        super().__init__(base, file_path, content_type)

    def add_part(self, part: 'Part'):
        self.parts.append(part)
