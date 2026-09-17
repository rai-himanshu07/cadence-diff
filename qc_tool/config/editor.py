"""Lossless draft operations for the typed reporting-contract editor."""

from __future__ import annotations

import copy
import hashlib
from dataclasses import dataclass
from pathlib import Path
from typing import Any, cast

import yaml

from qc_tool.config.lint import lint_profile
from qc_tool.config.profile import (
    DeliverableProfile,
    load_profile,
    profile_path,
    save_profile,
)

PathPart = str | int
DraftPath = tuple[PathPart, ...]
Schema = dict[str, Any]

PROFILE_SCHEMA: Schema = DeliverableProfile.model_json_schema(by_alias=True)
_DEFINITIONS: dict[str, Schema] = PROFILE_SCHEMA.get("$defs", {})


def profile_to_yaml(profile: DeliverableProfile) -> str:
    """Render the complete validated profile as canonical editor YAML."""
    return yaml.safe_dump(
        profile.model_dump(mode="json", by_alias=True),
        sort_keys=False,
    )


def yaml_to_profile(text: str) -> DeliverableProfile:
    """Parse editor YAML without accepting a scalar or sequence root."""
    raw = yaml.safe_load(text or "") or {}
    if not isinstance(raw, dict):
        raise ValueError("profile YAML must be a mapping")
    return DeliverableProfile.model_validate(raw)


def source_sha256(path: Path) -> str:
    """Hash the exact YAML bytes used to open an editor session."""
    return hashlib.sha256(path.read_bytes()).hexdigest()


def resolve_schema(schema: Schema) -> tuple[Schema, bool]:
    """Resolve local Pydantic references and return nullable state."""
    resolved = schema
    while "$ref" in resolved:
        name = str(resolved["$ref"]).rsplit("/", 1)[-1]
        resolved = _DEFINITIONS[name]
    variants = resolved.get("anyOf")
    if isinstance(variants, list):
        nullable = any(item.get("type") == "null" for item in variants)
        concrete = next(
            (item for item in variants if item.get("type") != "null"),
            {"type": "string"},
        )
        concrete_resolved, _ = resolve_schema(concrete)
        return concrete_resolved, nullable
    return resolved, False


def default_for_schema(schema: Schema) -> object:
    """Build a typed empty value for a newly added list or mapping entry."""
    resolved, nullable = resolve_schema(schema)
    if "default" in resolved:
        return copy.deepcopy(resolved["default"])
    if nullable:
        return None
    if "enum" in resolved:
        values = resolved["enum"]
        return values[0] if values else ""
    kind = resolved.get("type")
    if kind == "object":
        if "additionalProperties" in resolved and not resolved.get("properties"):
            return {}
        return {
            name: default_for_schema(field_schema)
            for name, field_schema in resolved.get("properties", {}).items()
        }
    if kind == "array":
        return []
    if kind == "boolean":
        return False
    if kind == "integer":
        return 0
    if kind == "number":
        return 0.0
    return ""


def schema_leaf_paths(
    schema: Schema = PROFILE_SCHEMA,
    prefix: tuple[str, ...] = (),
) -> set[tuple[str, ...]]:
    """Enumerate all schema leaves, using wildcards for dynamic containers."""
    resolved, _ = resolve_schema(schema)
    kind = resolved.get("type")
    if kind == "object" and resolved.get("properties"):
        paths: set[tuple[str, ...]] = set()
        for name, child in resolved["properties"].items():
            paths.update(schema_leaf_paths(child, (*prefix, name)))
        return paths
    if kind == "object" and "additionalProperties" in resolved:
        return schema_leaf_paths(resolved["additionalProperties"], (*prefix, "*"))
    if kind == "array":
        return schema_leaf_paths(resolved.get("items", {}), (*prefix, "[]"))
    return {prefix}


def schema_at_path(path: tuple[str, ...]) -> Schema:
    """Return the schema node reached by fields and dynamic-container markers."""
    schema = PROFILE_SCHEMA
    for part in path:
        resolved, _ = resolve_schema(schema)
        if part == "[]":
            schema = resolved["items"]
        elif part == "*":
            schema = resolved["additionalProperties"]
        else:
            schema = resolved["properties"][part]
    return schema


def _materialized_list(
    payload: dict[str, Any], path: DraftPath
) -> tuple[dict[str, Any], list[object]]:
    updated = copy.deepcopy(payload)
    node: object = updated
    for index, part in enumerate(path):
        is_last = index == len(path) - 1
        if isinstance(part, str):
            if not isinstance(node, dict):
                raise TypeError(f"invalid draft path component {part!r}")
            mapping = cast(dict[str, object], node)
            if part not in mapping:
                mapping[part] = (
                    [] if is_last or isinstance(path[index + 1], int) else {}
                )
            node = mapping[part]
        else:
            if not isinstance(node, list):
                raise TypeError(f"invalid draft path component {part!r}")
            node = cast(list[object], node)[part]
    if not isinstance(node, list):
        raise TypeError("draft path is not a list")
    return updated, node


@dataclass(slots=True)
class ProfileDraft:
    """One complete profile draft shared by the form and YAML views."""

    selected_name: str
    payload: dict[str, Any]
    opened_hash: str | None
    source_existed: bool
    dirty: bool = False

    @classmethod
    def from_profile(
        cls,
        profile: DeliverableProfile,
    ) -> ProfileDraft:
        return cls(
            selected_name=profile.name,
            payload=profile.model_dump(mode="json", by_alias=True),
            opened_hash=None,
            source_existed=False,
        )

    @classmethod
    def from_path(cls, path: Path) -> ProfileDraft:
        profile = load_profile(path)
        return cls(
            selected_name=profile.name,
            payload=profile.model_dump(mode="json", by_alias=True),
            opened_hash=source_sha256(path),
            source_existed=True,
        )

    def profile(self) -> DeliverableProfile:
        return DeliverableProfile.model_validate(self.payload)

    def yaml(self) -> str:
        return yaml.safe_dump(self.payload, sort_keys=False)

    def apply_yaml(self, text: str) -> None:
        profile = yaml_to_profile(text)
        self.payload = profile.model_dump(mode="json", by_alias=True)
        self.dirty = True

    def get(self, path: DraftPath) -> object:
        node: object = self.payload
        for part in path:
            if isinstance(part, str):
                if not isinstance(node, dict):
                    raise TypeError(f"invalid draft path component {part!r}")
                node = cast(dict[str, object], node)[part]
            else:
                if not isinstance(node, list):
                    raise TypeError(f"invalid draft path component {part!r}")
                node = cast(list[object], node)[part]
        return node

    def set(self, path: DraftPath, value: object) -> None:
        if not path:
            raise ValueError("cannot replace the profile root")
        parent = self.get(path[:-1]) if len(path) > 1 else self.payload
        part = path[len(path) - 1]
        if isinstance(part, str):
            if not isinstance(parent, dict):
                raise TypeError(f"invalid draft path component {part!r}")
            parent[part] = value
        else:
            if not isinstance(parent, list):
                raise TypeError(f"invalid draft path component {part!r}")
            parent[part] = value
        self.dirty = True

    def append(self, path: DraftPath, item_schema: Schema) -> None:
        try:
            values = self.get(path)
        except KeyError:
            payload, values = _materialized_list(self.payload, path)
            self.payload = payload
        if not isinstance(values, list):
            raise TypeError("draft path is not a list")
        values.append(default_for_schema(item_schema))
        self.dirty = True

    def remove(self, path: DraftPath, index: int) -> None:
        values = self.get(path)
        if not isinstance(values, list):
            raise TypeError("draft path is not a list")
        values.pop(index)
        self.dirty = True

    def move(self, path: DraftPath, index: int, offset: int) -> None:
        values = self.get(path)
        if not isinstance(values, list):
            raise TypeError("draft path is not a list")
        target = index + offset
        if target < 0 or target >= len(values):
            return
        values[index], values[target] = values[target], values[index]
        self.dirty = True

    def set_mapping(self, path: DraftPath, key: str, value: object) -> None:
        try:
            mapping = self.get(path)
        except KeyError:
            self.set(path, {})
            mapping = self.get(path)
        if not isinstance(mapping, dict):
            raise TypeError("draft path is not a mapping")
        key = key.strip()
        if not key:
            raise ValueError("mapping key cannot be blank")
        if key in mapping:
            raise ValueError(f"mapping key {key!r} already exists")
        mapping[key] = value
        self.dirty = True

    def rename_mapping_key(self, path: DraftPath, old: str, new: str) -> None:
        mapping = self.get(path)
        if not isinstance(mapping, dict):
            raise TypeError("draft path is not a mapping")
        new = new.strip()
        if not new:
            raise ValueError("mapping key cannot be blank")
        if new != old and new in mapping:
            raise ValueError(f"mapping key {new!r} already exists")
        items = [(new if key == old else key, value) for key, value in mapping.items()]
        mapping.clear()
        mapping.update(items)
        self.dirty = True

    def remove_mapping(self, path: DraftPath, key: str) -> None:
        mapping = self.get(path)
        if not isinstance(mapping, dict):
            raise TypeError("draft path is not a mapping")
        del mapping[key]
        self.dirty = True


def save_draft(draft: ProfileDraft, profiles_dir: Path) -> DeliverableProfile:
    """Validate and atomically save a draft if its source is still current."""
    if draft.selected_name == "default":
        raise ValueError("the default profile is immutable")
    profile = draft.profile()
    source = profile_path(profiles_dir, draft.selected_name)
    target = profile_path(profiles_dir, profile.name)
    if source.exists() != draft.source_existed:
        raise ValueError("profile changed while this editor was open; reopen it")
    if source.exists() and source_sha256(source) != draft.opened_hash:
        raise ValueError("profile changed while this editor was open; reopen it")
    if target != source and target.exists():
        raise ValueError(f"profile {profile.name!r} already exists")
    errors = [issue for issue in lint_profile(profile) if issue.level == "error"]
    if errors:
        raise ValueError(f"{errors[0].where}: {errors[0].message}")
    save_profile(profile, target)
    if target != source and source.exists():
        try:
            source.unlink()
        except OSError:
            target.unlink(missing_ok=True)
            raise
    draft.selected_name = profile.name
    draft.opened_hash = source_sha256(target)
    draft.source_existed = True
    draft.dirty = False
    return profile
