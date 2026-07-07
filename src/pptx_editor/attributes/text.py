from pptx_editor.attribute import Attribute

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
