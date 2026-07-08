from pptx_editor.attribute import Attribute
from pptx_editor.enums.ST_TextLanguageID import VALID_LANGUAGE_IDS, ST_TextLanguage, ST_TextRegion

class Typeface(Attribute):
    default_namespace = None
    default_prefix = None
    default_name = 'typeface'
    default_element_names = ['latin', 'ea', 'cs', 'sym', 'buFont']

class Bold(Attribute):
    default_namespace = None
    default_prefix = None
    default_name = 'b'
    default_element_names = ['rPr', 'defRPr']

class Italic(Attribute):
    default_namespace = None
    default_prefix = None
    default_name = 'i'
    default_element_names = ['rPr', 'defRPr']

class Underline(Attribute):
    default_namespace = None
    default_prefix = None
    default_name = 'u'
    default_element_names = ['rPr', 'defRPr']

class Strikethrough(Attribute):
    default_namespace = None
    default_prefix = None
    default_name = 'strike'
    default_element_names = ['rPr', 'defRPr']

class FontSize(Attribute):
    default_namespace = None
    default_prefix = None
    default_name = 'sz'
    default_element_names = ['rPr', 'defRPr']

class Dirty(Attribute):
    default_namespace = None
    default_prefix = None
    default_name = 'dirty'
    default_element_names = ['rPr', 'defRPr']

class Error(Attribute):
    default_namespace = None
    default_prefix = None
    default_name = 'err'
    default_element_names = ['rPr', 'defRPr']

class Language(Attribute):
    default_namespace = None
    default_prefix = None
    default_name = 'lang'
    default_element_names = ['rPr', 'defRPr']

    @staticmethod
    def from_enum(language: 'ST_TextLanguage', region: 'ST_TextRegion | None' = None) -> str:
        for lang_id in VALID_LANGUAGE_IDS:
            if lang_id.startswith(language.value):
                if (region is None and lang_id.endswith(language.value)) or (region is not None and lang_id.endswith(region.value)):
                    return lang_id

        raise ValueError(f"Invalid language ID for language '{language}' and region '{region}'")

