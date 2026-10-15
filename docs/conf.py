# -*- coding: utf-8 -*-
"""
Sphinx configuration file for OSTRAM documentation.

Author: Climate Lead Group, Andrey Salazar-Vargas
Date: 2025
"""

project = "OSTRAM"
copyright = "2026, Climate Lead Group"
author = "Climate Lead Group"
release = "2026.10"
version = "2026.10"

extensions = [
    "myst_parser",
    "sphinx.ext.autosectionlabel",
    "sphinxcontrib.mermaid",
]

myst_enable_extensions = [
    "colon_fence",
    "fieldlist",
]

myst_heading_anchors = 3

# Render plain ```mermaid fences (as GitHub does) through the mermaid directive.
myst_fence_as_directive = ["mermaid"]

# Neutral palette stays readable on both Furo backgrounds.
mermaid_light_theme = "neutral"
mermaid_dark_theme = "neutral"

source_suffix = {
    ".rst": "restructuredtext",
    ".md": "markdown",
}

templates_path = ["_templates"]
exclude_patterns = ["_build"]

html_theme = "furo"
html_theme_options = {
    "source_repository": "https://github.com/Climate-Lead-Group/OSTRAM/",
    "source_branch": "main",
    "source_directory": "docs/",
}
