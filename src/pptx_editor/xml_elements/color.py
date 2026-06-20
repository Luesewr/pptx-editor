from abc import ABC, abstractmethod
from typing import TYPE_CHECKING, Self

from pptx_editor.parts.xml_part import XmlPart
from pptx_editor.xml_element import XmlElement
from pptx_editor.attribute import Attribute
from pptx_editor.enums.ST_PresetColorVal import ST_PresetColorVal
from pptx_editor.exceptions import PowerpointIntegrityError

if TYPE_CHECKING:
    from pptx_editor.parts.theme import Theme

class ColorMapOverride(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'clrMapOvr'

class ColorMap(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'clrMap'

    @property
    def colors(self) -> dict[str, str]:
        return {element.name: element.value for element in self.attributes}

class OverrideColorMapping(ColorMap):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'overrideClrMapping'

class MasterColorMapping(ColorMap):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'masterClrMapping'

class Color(XmlElement, ABC):
    is_abstract = True

    @abstractmethod
    def as_rgb(self) -> tuple[int, int, int]:
        pass

    @classmethod
    @abstractmethod
    def from_rgb(cls, part: 'XmlPart', r: int, g: int, b: int) -> Self:
        pass

    def __init_subclass__(cls, **kwargs):
        cls.is_abstract = False
        super().__init_subclass__(**kwargs)


class HslColor(Color):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'hslClr'

    def as_rgb(self) -> tuple[int, int, int]:
        h_attribute = self.get_attribute('hue', 'a')
        s_attribute = self.get_attribute('sat', 'a')
        l_attribute = self.get_attribute('lum', 'a')

        if h_attribute is None or s_attribute is None or l_attribute is None:
            raise PowerpointIntegrityError('The hslClr element is missing one or more required attributes.')

        h = int(h_attribute.value)
        s = int(s_attribute.value)
        l = int(l_attribute.value)

        # Convert HSL to RGB
        if s == 0:
            r = g = b = l
        else:
            q = l * (1 + s) if l < 0.5 else l + s - l * s
            p = 2 * l - q
            r = int(_hue_to_rgb(p, q, h + 1/3) * 255)
            g = int(_hue_to_rgb(p, q, h) * 255)
            b = int(_hue_to_rgb(p, q, h - 1/3) * 255)

        return r, g, b

    @classmethod
    def from_rgb(cls, part: 'XmlPart', r: int, g: int, b: int) -> Self:
        r /= 255
        g /= 255
        b /= 255

        max_val = max(r, g, b)
        min_val = min(r, g, b)
        h = s = l = (max_val + min_val) / 2

        if max_val == min_val:
            h = s = 0
        else:
            d = max_val - min_val
            s = d / (1 - abs(2 * l - 1))
            if max_val == r:
                h = (g - b) / d + (6 if g < b else 0)
            elif max_val == g:
                h = (b - r) / d + 2
            else:
                h = (r - g) / d + 4
            h /= 6

        attributes = (
            Attribute('a', 'hue', str(int(h * 360))),
            Attribute('a', 'sat', str(int(s * 100))),
            Attribute('a', 'lum', str(int(l * 100)))
        )

        return cls(attributes=attributes, part=part)


class PresetColor(Color):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'prstClr'

    def as_rgb(self) -> tuple[int, int, int]:
        val_attribute = self.get_attribute('val')

        if val_attribute is None:
            raise PowerpointIntegrityError('The prstClr element is missing the required val attribute.')

        val = val_attribute.value

        if val not in ST_PresetColorVal:
            raise PowerpointIntegrityError(f'Invalid preset color value: {val}')

        r, g, b = ST_PresetColorVal[val]
        return r, g, b

    @classmethod
    def from_rgb(cls, part: 'XmlPart', r: int, g: int, b: int) -> Self:
        for val, (pr, pg, pb) in ST_PresetColorVal.items():
            if (r, g, b) == (pr, pg, pb):
                attributes = (Attribute('a', 'val', val),)
                return cls(attributes=attributes, part=part)

        raise ValueError(f'No preset color matches the RGB value ({r}, {g}, {b})')


class SchemeColor(Color):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'schemeClr'

    def as_rgb(self) -> tuple[int, int, int]:
        val_attribute = self.get_attribute('val')

        if val_attribute is None:
            raise PowerpointIntegrityError('The schemeClr element is missing the required val attribute.')

        val = val_attribute.value

        color_map: dict[str, str] | None = getattr(self.part, 'color_map', None)

        if color_map is None:
            raise PowerpointIntegrityError('SchemeColor element does not have a color map to resolve the scheme color value.')

        if val not in color_map:
            raise PowerpointIntegrityError(f'Scheme color value {val} not found in color map.')

        color_value = color_map[val]

        if not hasattr(self.part, 'theme'):
            raise PowerpointIntegrityError('SchemeColor element does not have an associated Theme.')

        theme: Theme = getattr(self.part, 'theme')

        return theme.get_color(color_value).as_rgb()

    @classmethod
    def from_rgb(cls, part: 'XmlPart', r: int, g: int, b: int) -> Self:
        if not hasattr(part, 'color_map'):
            raise PowerpointIntegrityError('SchemeColor class does not have a color map to resolve the scheme color value.')

        color_map: dict[str, str] = getattr(part, 'color_map')

        if not hasattr(part, 'theme'):
            raise PowerpointIntegrityError('SchemeColor class does not have an associated Theme.')

        theme: Theme = getattr(part, 'theme')

        for name, color_value in color_map.items():
            if theme.get_color(color_value).as_rgb() == (r, g, b):
                attributes = (Attribute('a', 'val', name),)
                return cls(attributes=attributes, part=part)

        raise ValueError(f'No scheme color matches the RGB value ({r}, {g}, {b})')


class RgbColorModelPercentage(Color):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'scrgbClr'

    def as_rgb(self) -> tuple[int, int, int]:
        r_attribute = self.get_attribute('r', 'a')
        g_attribute = self.get_attribute('g', 'a')
        b_attribute = self.get_attribute('b', 'a')

        if r_attribute is None or g_attribute is None or b_attribute is None:
            raise PowerpointIntegrityError('The scrgbClr element is missing one or more required attributes.')

        r = int(r_attribute.value) * 255 / 100000
        g = int(g_attribute.value) * 255 / 100000
        b = int(b_attribute.value) * 255 / 100000

        return r, g, b

    @classmethod
    def from_rgb(cls, part: 'XmlPart', r: int, g: int, b: int) -> Self:
        attributes = (
            Attribute('a', 'r', str(int(r * 100000 / 255))),
            Attribute('a', 'g', str(int(g * 100000 / 255))),
            Attribute('a', 'b', str(int(b * 100000 / 255))),
        )

        return cls(attributes=attributes, part=part)


class RgbColorModelHex(Color):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'srgbClr'

    def as_rgb(self) -> tuple[int, int, int]:
        val_attribute = self.get_attribute('val')

        if val_attribute is None:
            raise PowerpointIntegrityError('The srgbClr element is missing the required val attribute.')

        val = val_attribute.value

        if len(val) != 6:
            raise PowerpointIntegrityError(f'Invalid srgbClr value: {val}. Expected a 6-character hex string.')

        r = int(val[0:2], 16)
        g = int(val[2:4], 16)
        b = int(val[4:6], 16)

        return r, g, b

    @classmethod
    def from_rgb(cls, part: 'XmlPart', r: int, g: int, b: int) -> Self:
        val = f'{r:02X}{g:02X}{b:02X}'
        attributes = (Attribute('a', 'val', val),)
        return cls(attributes=attributes, part=part)


class SystemColor(Color):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'sysClr'

    def as_rgb(self) -> tuple[int, int, int]:
        raise NotImplementedError('SystemColor does not support conversion to RGB since it depends on the system color scheme.')

    @classmethod
    def from_rgb(cls, part: 'XmlPart', r: int, g: int, b: int) -> Self:
        raise NotImplementedError('SystemColor does not support creation from RGB since it depends on the system color scheme.')


def _hue_to_rgb(p, q, t):
    if t < 0:
        t += 1
    if t > 1:
        t -= 1
    if t < 1/6:
        return p + (q - p) * 6 * t
    if t < 1/2:
        return q
    if t < 2/3:
        return p + (q - p) * (2/3 - t) * 6
    return p
