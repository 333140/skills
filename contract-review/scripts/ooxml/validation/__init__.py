"""Validator exports for lightweight OOXML checks."""

from .docx import DOCXSchemaValidator as DocxValidator

DOCXSchemaValidator = DocxValidator

__all__ = [
    "DOCXSchemaValidator",
]
