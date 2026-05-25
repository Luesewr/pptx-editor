"""pptx-editor: A Python library for editing PowerPoint (.pptx) files."""

__version__ = "0.1.0"

import pkgutil
import importlib

import pptx_editor.parts
import pptx_editor.xml_elements
import pptx_editor.attributes

from pptx_editor.parts.presentation import Presentation


def import_submodules(package):
    for _, module_name, _ in pkgutil.iter_modules(package.__path__):
        importlib.import_module(f"{package.__name__}.{module_name}")

import_submodules(pptx_editor.parts)
import_submodules(pptx_editor.xml_elements)
import_submodules(pptx_editor.attributes)
