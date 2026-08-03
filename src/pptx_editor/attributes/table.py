from pptx_editor.attribute import Attribute

class Width(Attribute):
    default_namespace = None
    default_prefix = None
    default_name = 'w'
    default_element_names = ['gridCol',]

class Height(Attribute):
    default_namespace = None
    default_prefix = None
    default_name = 'h'
    default_element_names = ['tr',]
