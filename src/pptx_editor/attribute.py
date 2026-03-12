import sys

from lxml import etree

class AttributeValue:
    def __init__(self, name: str, value: str):
        q = etree.QName(name)
        self.namespace = sys.intern(q.namespace) if q.namespace else None
        self.name = q.localname
        self.value = value

    def __str__(self):
        return f"{self.namespace}:{self.name}={self.value}" if self.namespace else f"{self.name}={self.value}"

class Attribute:
    def __init__(self, xml: etree._Element):
        q = etree.QName(xml)
        self.name = q.localname
        self.namespace = sys.intern(q.namespace) if q.namespace else None
        self.values = [AttributeValue(str(key), str(value)) for key, value in xml.attrib.items()]
        self.attributes = [Attribute(child) for child in xml]

    def get_values(self, name: str, namespace: str | None = None) -> list[AttributeValue]:
        return [value for value in self.values if value.name == name and value.namespace == namespace]

    def pretty_print(self, indent=0):
        indent_str = ' ' * indent
        print(f"{indent_str}{self}")
        for child in self.attributes:
            child.pretty_print(indent + 2)

    def __str__(self):
        return f"Attribute(name={self.name}, namespace={self.namespace}, values={[str(value) for value in self.values]})"
