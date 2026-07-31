"""Profiles — the one place a house opinion is allowed to live.

The rule kinds keep the library's own opinions out of the linter:
``consistency`` and ``structural`` rules judge a document against itself,
and ``policy`` rules — the ones that need somebody to say what "right"
looks like — are inert without a target. A **profile** is where those
targets come from, and the reason `docx_plus` can ship no house style
while still being useful to a team that has one.

```json
{
  "rules": {
    "double-space":     {"enabled": false},
    "style-drift":      {"severity": "error"},
    "font-outliers":    {"enabled": true, "options": {"max_share": 0.02}}
  }
}
```

**No rule reads ``options`` yet**, because no ``policy`` rule ships yet.
The loader is here now on purpose: a profile is the *interface* between
a rule catalogue and a team's conventions, and settling that shape while
there is nothing to break is cheaper than retrofitting it around whichever
policy rule happens to land first.

What a profile deliberately does **not** do is select rules. ``--rule`` /
``--exclude`` stay the caller's, applied after the profile, so a profile
never stops someone asking a specific question of a specific document.
"""

from __future__ import annotations

import json
from collections.abc import Mapping
from dataclasses import dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING, Any

from docx_plus.core import DocxPlusError

if TYPE_CHECKING:
    from docx_plus.lint.models import Severity


_SEVERITIES = frozenset({"error", "warning", "info"})

DEFAULT_PROFILE_NAME = "docx-plus-lint.json"
"""The file name :meth:`Profile.discover` looks for, so a repository can
check its conventions in beside the documents they govern."""


class InvalidProfileError(DocxPlusError, ValueError):
    """Raised when a profile is malformed.

    Loudly, and on load rather than on use: a profile with a typo in a rule
    id would otherwise configure nothing and read exactly like a profile
    that was working.
    """


@dataclass(frozen=True)
class RuleSettings:
    """What a profile says about one rule.

    Attributes:
        enabled: Force the rule on or off, overriding its ``default_on``.
            ``None`` leaves the default alone.
        severity: Report this rule's findings at another severity — the
            knob a team reaches for first, since "we treat drift as an
            error" is a house opinion that changes no detection.
        options: Rule-specific values. Untouched by the loader beyond
            being required to be a JSON object; each rule validates its
            own.
    """

    enabled: bool | None = None
    severity: Severity | None = None
    options: Mapping[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class Profile:
    """A loaded lint profile.

    Attributes:
        rules: Settings by rule id. A rule the profile does not mention
            keeps its registered behaviour, so a profile is a set of
            deltas rather than a replacement catalogue.
    """

    rules: Mapping[str, RuleSettings] = field(default_factory=dict)

    def settings(self, rule_id: str) -> RuleSettings:
        """What this profile says about ``rule_id``, or the empty settings."""
        return self.rules.get(rule_id, RuleSettings())

    def enabled(self, rule_id: str, *, default: bool) -> bool:
        """Whether ``rule_id`` runs, given its registered ``default``."""
        override = self.settings(rule_id).enabled
        return default if override is None else override

    def severity(self, rule_id: str, *, default: Severity) -> Severity:
        """The severity to report ``rule_id`` at, given its registered ``default``."""
        return self.settings(rule_id).severity or default

    def option(self, rule_id: str, key: str, default: Any = None) -> Any:
        """One rule-specific value, or ``default``.

        The hook every ``policy`` rule will read its target through. Nothing
        calls it yet — see the module docstring.
        """
        return self.settings(rule_id).options.get(key, default)

    @classmethod
    def load(cls, source: str | Path | Mapping[str, Any] | None) -> Profile:
        """Build a profile from a path, an already-parsed mapping, or nothing.

        Args:
            source: A path to a JSON file, a mapping in the same shape, or
                ``None`` for the empty profile — which changes nothing, so
                every caller can pass its argument straight through without
                branching.

        Returns:
            The profile.

        Raises:
            InvalidProfileError: If the file is unreadable, is not JSON, or
                is not in the documented shape.

        Example:
            >>> profile = Profile.load({"rules": {"double-space": {"enabled": False}}})
            >>> profile.enabled("double-space", default=True)
            False
            >>> profile.enabled("style-drift", default=True)
            True
        """
        if source is None:
            return cls()
        if isinstance(source, str | Path):
            return cls._parse(_read(Path(source)), origin=str(source))
        return cls._parse(source, origin="the supplied mapping")

    @classmethod
    def discover(cls, start: str | Path) -> Profile:
        """Find and load :data:`DEFAULT_PROFILE_NAME` at or above ``start``.

        Walks up from ``start`` (a file or a directory) to the filesystem
        root, so running the linter from anywhere inside a project picks up
        the conventions checked in at its top.

        Args:
            start: Where to begin looking.

        Returns:
            The first profile found, or the empty profile if there is none.

        Raises:
            InvalidProfileError: If a profile is found and is malformed.
                Silently ignoring a broken checked-in profile would be
                worse than not having one.
        """
        here = Path(start).resolve()
        if here.is_file():
            here = here.parent
        for directory in (here, *here.parents):
            candidate = directory / DEFAULT_PROFILE_NAME
            if candidate.is_file():
                return cls.load(candidate)
        return cls()

    @classmethod
    def _parse(cls, raw: Mapping[str, Any], *, origin: str) -> Profile:
        """Validate the documented shape, naming ``origin`` in every complaint."""
        if not isinstance(raw, Mapping):
            raise InvalidProfileError(f"{origin}: a profile must be a JSON object")

        unknown = set(raw) - {"rules"}
        if unknown:
            raise InvalidProfileError(
                f"{origin}: unknown profile key(s): {', '.join(sorted(unknown))}"
            )

        rules_raw = raw.get("rules", {})
        if not isinstance(rules_raw, Mapping):
            raise InvalidProfileError(f"{origin}: 'rules' must be an object keyed by rule id")

        return cls(
            rules={
                rule_id: _rule_settings(settings, origin=f"{origin}: rule {rule_id!r}")
                for rule_id, settings in rules_raw.items()
            }
        )


def _rule_settings(raw: object, *, origin: str) -> RuleSettings:
    """Validate one rule's block."""
    if not isinstance(raw, Mapping):
        raise InvalidProfileError(f"{origin}: must be an object")

    unknown = set(raw) - {"enabled", "severity", "options"}
    if unknown:
        raise InvalidProfileError(f"{origin}: unknown key(s): {', '.join(sorted(unknown))}")

    enabled = raw.get("enabled")
    if enabled is not None and not isinstance(enabled, bool):
        raise InvalidProfileError(f"{origin}: 'enabled' must be true or false")

    severity = raw.get("severity")
    if severity is not None and severity not in _SEVERITIES:
        raise InvalidProfileError(
            f"{origin}: 'severity' must be one of {', '.join(sorted(_SEVERITIES))}"
        )

    options = raw.get("options", {})
    if not isinstance(options, Mapping):
        raise InvalidProfileError(f"{origin}: 'options' must be an object")

    return RuleSettings(enabled=enabled, severity=severity, options=dict(options))


def _read(path: Path) -> Mapping[str, Any]:
    """Read and parse a profile file."""
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:
        raise InvalidProfileError(f"{path}: cannot read profile ({exc})") from exc
    try:
        parsed = json.loads(text)
    except json.JSONDecodeError as exc:
        raise InvalidProfileError(f"{path}: not valid JSON ({exc})") from exc
    if not isinstance(parsed, dict):
        raise InvalidProfileError(f"{path}: a profile must be a JSON object")
    return parsed


__all__ = [
    "DEFAULT_PROFILE_NAME",
    "InvalidProfileError",
    "Profile",
    "RuleSettings",
]
