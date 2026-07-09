from pptx_editor.attribute import Attribute


class Hue(Attribute):
    default_namespace = None
    default_prefix = None
    default_name = 'hue'
    default_element_names = ['hslClr']

class Saturation(Attribute):
    default_namespace = None
    default_prefix = None
    default_name = 'sat'
    default_element_names = ['hslClr']

class Luminance(Attribute):
    default_namespace = None
    default_prefix = None
    default_name = 'lum'
    default_element_names = ['hslClr']

class Red(Attribute):
    default_namespace = None
    default_prefix = None
    default_name = 'r'
    default_element_names = ['scrgbClr']

class Green(Attribute):
    default_namespace = None
    default_prefix = None
    default_name = 'g'
    default_element_names = ['scrgbClr']

class Blue(Attribute):
    default_namespace = None
    default_prefix = None
    default_name = 'b'
    default_element_names = ['scrgbClr']

class ColorValue(Attribute):
    default_namespace = None
    default_prefix = None
    default_name = 'val'
    default_element_names = ['srgbClr', 'prstClr', 'schemeClr', 'scrgbClr', 'hslClr']
