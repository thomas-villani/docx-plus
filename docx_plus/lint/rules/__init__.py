"""Built-in lint rules.

Importing this package runs every rule module's :func:`~docx_plus.lint.registry.rule`
decorators, which is what populates the registry. One module per cluster;
adding a rule is a single new function, with no central list to update.
"""

from docx_plus.lint.rules import formatting, references, structure, styles, typography

__all__ = ["formatting", "references", "structure", "styles", "typography"]
