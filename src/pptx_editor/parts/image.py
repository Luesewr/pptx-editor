from pathlib import PurePosixPath

from pptx_editor.part import Part


class ImagePart(Part):
    default_base_path: PurePosixPath | None = PurePosixPath('/ppt/media')
    default_part_name: str | None = 'image{i}'
    default_shared: bool = True
    is_abstract: bool = True

    def __init_subclass__(cls, **kwargs):
        cls.is_abstract = False
        super().__init_subclass__(**kwargs)


class PngImagePart(ImagePart):
    default_content_type = 'image/png'
    default_extension: str | None = 'png'


class JpegImagePart(ImagePart):
    default_content_type = 'image/jpeg'
    default_extension: str | None = 'jpeg'


class JpgImagePart(ImagePart):
    default_content_type = 'image/jpg'
    default_extension: str | None = 'jpg'


class GifImagePart(ImagePart):
    default_content_type = 'image/gif'
    default_extension: str | None = 'gif'


class BmpImagePart(ImagePart):
    default_content_type = 'image/bmp'
    default_extension: str | None = 'bmp'


class TiffImagePart(ImagePart):
    default_content_type = 'image/tiff'
    default_extension: str | None = 'tiff'


class WebPImagePart(ImagePart):
    default_content_type = 'image/webp'
    default_extension: str | None = 'webp'


class SvgImagePart(ImagePart):
    default_content_type = 'image/svg+xml'
    default_extension: str | None = 'svg'


class EmfImagePart(ImagePart):
    default_content_type = 'image/x-emf'
    default_extension: str | None = 'emf'


class WmfImagePart(ImagePart):
    default_content_type = 'image/x-wmf'
    default_extension: str | None = 'wmf'


class AvifImagePart(ImagePart):
    default_content_type = 'image/avif'
    default_extension: str | None = 'avif'
