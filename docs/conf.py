# -*- coding: utf-8 -*-
"""
Sphinx configuration file for OSTRAM documentation.

Author: Climate Lead Group, Andrey Salazar-Vargas
Date: 2025
"""

project = "OSTRAM"
copyright = "2025, Climate Lead Group"
author = "Climate Lead Group"
release = "1.0"

extensions = [
    "myst_parser",
    "sphinx.ext.autosectionlabel",
]

myst_enable_extensions = [
    "colon_fence",
    "fieldlist",
]

myst_heading_anchors = 3

source_suffix = {
    ".rst": "restructuredtext",
    ".md": "markdown",
}

templates_path = ["_templates"]
exclude_patterns = ["_build"]

html_theme = "sphinx_rtd_theme"
html_theme_options = {
    "navigation_depth": 3,
    "collapse_navigation": False,
    "titles_only": False,
}
html_static_path = ["_static"]
