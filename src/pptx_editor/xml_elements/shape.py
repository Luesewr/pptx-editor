from pptx_editor.properties.xml_element import XmlElementProperty
from pptx_editor.xml_element import XmlElement
from pptx_editor.xml_elements.text import TextBody

class AbstractShape(XmlElement):
    is_abstract = True

    text_body = XmlElementProperty(TextBody)

    def __init_subclass__(cls, **kwargs):
        cls.is_abstract = False
        super().__init_subclass__(**kwargs)


class NonVisualShapeProperties(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'nvSpPr'
    default_order = ('cNvPr', 'cNvSpPr',)


class NonVisualGroupShapeProperties(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'nvGrpSpPr'
    default_order = ('cNvPr', 'cNvGrpSpPr',)


class GroupShapeProperties(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'grpSpPr'
    default_order = (
        'xfrm', ('blipFill', 'gradFill', 'grpFill', 'noFill',
        'pattFill', 'solidFill',), ('effectDag', 'effectLst',),
        'scene3d', 'extLst',
    )


class GroupShape(AbstractShape):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'grpSp'
    default_order = ('nvGrpSpPr', 'grpSpPr', ('sp', 'grpSp', 'graphicFrame', 'cxnSp', 'pic',),)


class Picture(AbstractShape):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'pic'
    default_order = ('nvPicPr', 'blipFill', 'spPr', 'style',)


class GraphicFrame(AbstractShape):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'graphicFrame'
    default_order = ('nvGraphicFramePr', 'xfrm', 'graphic',)


class ConnectionShape(AbstractShape):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'cxnSp'
    default_order = ('nvCxnSpPr', 'spPr', 'style',)


class Shape(AbstractShape):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'sp'
    default_order = ('nvSpPr', 'spPr', 'style', 'txBody',)

    non_visual_shape_properties = XmlElementProperty(NonVisualShapeProperties)
