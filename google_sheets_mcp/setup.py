#!/usr/bin/env python3
"""
Setup script for the Google Sheets MCP Server package.
"""

from setuptools import setup, find_packages
import os

# Read the contents of the README file
with open("README.md", encoding="utf-8") as f:
    long_description = f.read()

# Read the requirements from requirements.txt
with open("requirements.txt", encoding="utf-8") as f:
    requirements = [line.strip() for line in f if line.strip() and not line.startswith("#")]

# Development requirements (excluding comments and empty lines)
dev_requirements = [
    req for req in requirements 
    if any(dev_tool in req for dev_tool in ["pytest", "black", "isort", "mypy", "pylint", "cov"])
]

# Main requirements (excluding development requirements)
main_requirements = [req for req in requirements if req not in dev_requirements]

setup(
    name="google-sheets-mcp",
    version="0.1.0",
    description="MCP server for Google Sheets integration",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="Your Name",
    author_email="your.email@example.com",
    url="https://github.com/yourusername/google-sheets-mcp",
    packages=find_packages(),
    python_requires=">=3.9",
    install_requires=main_requirements,
    extras_require={
        "dev": dev_requirements,
    },
    entry_points={
        "console_scripts": [
            "google-sheets-mcp=google_sheets_mcp.src.server:main",
        ],
    },
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Topic :: Software Development :: Libraries :: Python Modules",
    ],
    keywords="mcp, google sheets, spreadsheet, api, model context protocol",
    project_urls={
        "Bug Reports": "https://github.com/yourusername/google-sheets-mcp/issues",
        "Source": "https://github.com/yourusername/google-sheets-mcp",
    },
)
