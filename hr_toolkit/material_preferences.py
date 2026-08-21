from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from typing import Any, Iterable, Mapping

from .common.filenames import safe_filename
from .tools.material_collector import MATERIAL_SYNONYMS


BUILTIN_MATERIALS: tuple[str, ...] = tuple(MATERIAL_SYNONYMS)
BUILTIN_MATERIAL_PRESETS: dict[str, tuple[str, ...]] = {
    "入职材料": ("身份证", "劳动合同"),
    "证书材料": ("特种证书", "资格证书"),
}

_MAX_CUSTOM_MATERIALS = 100
_MAX_CUSTOM_PRESETS = 100
_MAX_NAME_LENGTH = 40
_INVALID_MATERIAL_NAME = re.compile(r'[<>:"/\\|?*\x00-\x1f,，、;；]')
_CONTROL_CHARACTER = re.compile(r"[\x00-\x1f\x7f]")


@dataclass(frozen=True)
class MaterialRemovalResult:
    updated_presets: tuple[str, ...] = ()
    removed_presets: tuple[str, ...] = ()


def _normalize_name(value: Any, *, label: str, material: bool = False) -> str:
    if not isinstance(value, str):
        raise ValueError(f"{label}不能为空。")
    name = unicodedata.normalize("NFKC", value)
    name = re.sub(r"\s+", " ", name).strip()
    if not name:
        raise ValueError(f"{label}不能为空。")
    if len(name) > _MAX_NAME_LENGTH:
        raise ValueError(f"{label}不能超过 {_MAX_NAME_LENGTH} 个字符。")
    if _CONTROL_CHARACTER.search(name):
        raise ValueError(f"{label}包含不支持的控制字符。")
    if material and _INVALID_MATERIAL_NAME.search(name):
        raise ValueError(f"{label}不能包含逗号或 Windows 文件名不支持的字符。")
    if material and safe_filename(name, fallback="") != name:
        raise ValueError(f"{label}不是安全的文件夹名称，请换一个名称。")
    return name


def _name_key(value: str) -> str:
    return unicodedata.normalize("NFKC", value).casefold()


def _deduplicate_names(values: Iterable[Any], *, material: bool) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for raw_value in values:
        try:
            value = _normalize_name(
                raw_value,
                label="材料名称" if material else "预设名称",
                material=material,
            )
        except ValueError:
            continue
        key = _name_key(value)
        if key in seen:
            continue
        seen.add(key)
        result.append(value)
    return result


class MaterialPreferences:
    """Local material catalog and preset configuration.

    The class deliberately contains no Tk state so it can be validated before the
    GUI is created and safely persisted in the existing workspace-ui.json file.
    """

    def __init__(
        self,
        *,
        custom_materials: Iterable[Any] = (),
        custom_presets: Mapping[Any, Any] | None = None,
    ) -> None:
        builtin_keys = {_name_key(name) for name in BUILTIN_MATERIALS}
        custom: list[str] = []
        for name in _deduplicate_names(custom_materials, material=True):
            if _name_key(name) in builtin_keys:
                continue
            custom.append(name)
            if len(custom) >= _MAX_CUSTOM_MATERIALS:
                break
        self._custom_materials = custom
        self._custom_presets: dict[str, tuple[str, ...]] = {}

        if not isinstance(custom_presets, Mapping):
            return
        builtin_preset_keys = {_name_key(name) for name in BUILTIN_MATERIAL_PRESETS}
        for raw_name, raw_materials in custom_presets.items():
            if len(self._custom_presets) >= _MAX_CUSTOM_PRESETS:
                break
            try:
                name = _normalize_name(raw_name, label="预设名称")
            except ValueError:
                continue
            key = _name_key(name)
            if key in builtin_preset_keys or any(
                _name_key(existing) == key for existing in self._custom_presets
            ):
                continue
            if not isinstance(raw_materials, (list, tuple)):
                continue
            materials = self._validated_materials(raw_materials, ignore_unknown=True)
            # Stale presets that lost every material are omitted on load. This keeps
            # the persisted model valid after a custom material is deleted manually.
            if materials:
                self._custom_presets[name] = materials

    @classmethod
    def from_payload(cls, payload: Any) -> "MaterialPreferences":
        if not isinstance(payload, Mapping):
            return cls()
        raw_materials = payload.get("custom_materials", ())
        if not isinstance(raw_materials, (list, tuple)):
            raw_materials = ()
        raw_presets = payload.get("custom_presets", {})
        if not isinstance(raw_presets, Mapping):
            raw_presets = {}
        return cls(custom_materials=raw_materials, custom_presets=raw_presets)

    @property
    def custom_materials(self) -> tuple[str, ...]:
        return tuple(self._custom_materials)

    @property
    def custom_presets(self) -> dict[str, tuple[str, ...]]:
        return dict(self._custom_presets)

    @property
    def available_materials(self) -> tuple[str, ...]:
        return BUILTIN_MATERIALS + tuple(self._custom_materials)

    @property
    def preset_names(self) -> tuple[str, ...]:
        return tuple(BUILTIN_MATERIAL_PRESETS) + tuple(self._custom_presets)

    def to_payload(self) -> dict[str, Any]:
        return {
            "version": 1,
            "custom_materials": list(self._custom_materials),
            "custom_presets": {
                name: list(materials)
                for name, materials in self._custom_presets.items()
            },
        }

    def get_preset(self, name: str) -> tuple[str, ...] | None:
        key = _name_key(str(name))
        for preset_name, materials in BUILTIN_MATERIAL_PRESETS.items():
            if _name_key(preset_name) == key:
                return materials
        for preset_name, materials in self._custom_presets.items():
            if _name_key(preset_name) == key:
                return materials
        return None

    def is_builtin_preset(self, name: str) -> bool:
        key = _name_key(str(name))
        return any(_name_key(item) == key for item in BUILTIN_MATERIAL_PRESETS)

    def add_material(self, raw_name: Any) -> str:
        name = _normalize_name(raw_name, label="材料名称", material=True)
        key = _name_key(name)
        if any(_name_key(item) == key for item in self.available_materials):
            raise ValueError(f"材料“{name}”已经存在。")
        if len(self._custom_materials) >= _MAX_CUSTOM_MATERIALS:
            raise ValueError(f"自定义材料最多可添加 {_MAX_CUSTOM_MATERIALS} 种。")
        self._custom_materials.append(name)
        return name

    def remove_material(self, raw_name: Any) -> MaterialRemovalResult:
        name = _normalize_name(raw_name, label="材料名称", material=True)
        key = _name_key(name)
        matched = next(
            (item for item in self._custom_materials if _name_key(item) == key),
            None,
        )
        if matched is None:
            if any(_name_key(item) == key for item in BUILTIN_MATERIALS):
                raise ValueError("内置材料不能删除。")
            raise ValueError(f"没有找到自定义材料“{name}”。")

        self._custom_materials.remove(matched)
        updated: list[str] = []
        removed: list[str] = []
        for preset_name, materials in list(self._custom_presets.items()):
            filtered = tuple(item for item in materials if _name_key(item) != key)
            if filtered == materials:
                continue
            if filtered:
                self._custom_presets[preset_name] = filtered
                updated.append(preset_name)
            else:
                del self._custom_presets[preset_name]
                removed.append(preset_name)
        return MaterialRemovalResult(tuple(updated), tuple(removed))

    def save_preset(
        self,
        raw_name: Any,
        materials: Iterable[Any],
        *,
        replacing: str | None = None,
    ) -> str:
        name = _normalize_name(raw_name, label="预设名称")
        selected = self._validated_materials(materials)
        if not selected:
            raise ValueError("预设至少需要选择一种材料。")

        name_key = _name_key(name)
        if any(_name_key(item) == name_key for item in BUILTIN_MATERIAL_PRESETS):
            raise ValueError("内置预设不能覆盖，请使用其他名称。")

        replacing_name: str | None = None
        if replacing is not None:
            replacing_key = _name_key(str(replacing))
            replacing_name = next(
                (item for item in self._custom_presets if _name_key(item) == replacing_key),
                None,
            )
            if replacing_name is None:
                raise ValueError("要编辑的自定义预设不存在。")

        collision = next(
            (
                item
                for item in self._custom_presets
                if _name_key(item) == name_key and item != replacing_name
            ),
            None,
        )
        if collision is not None:
            raise ValueError(f"预设“{name}”已经存在。")
        if replacing_name is None and len(self._custom_presets) >= _MAX_CUSTOM_PRESETS:
            raise ValueError(f"自定义预设最多可保存 {_MAX_CUSTOM_PRESETS} 个。")

        if replacing_name is not None:
            items = list(self._custom_presets.items())
            self._custom_presets.clear()
            for existing_name, existing_materials in items:
                if existing_name == replacing_name:
                    self._custom_presets[name] = selected
                else:
                    self._custom_presets[existing_name] = existing_materials
        else:
            self._custom_presets[name] = selected
        return name

    def rename_preset(self, current_name: str, new_name: Any) -> str:
        materials = self.get_preset(current_name)
        if materials is None or self.is_builtin_preset(current_name):
            raise ValueError("只能重命名自定义预设。")
        return self.save_preset(new_name, materials, replacing=current_name)

    def delete_preset(self, name: str) -> None:
        if self.is_builtin_preset(name):
            raise ValueError("内置预设不能删除。")
        key = _name_key(str(name))
        matched = next(
            (item for item in self._custom_presets if _name_key(item) == key),
            None,
        )
        if matched is None:
            raise ValueError("要删除的自定义预设不存在。")
        del self._custom_presets[matched]

    def _validated_materials(
        self,
        values: Iterable[Any],
        *,
        ignore_unknown: bool = False,
    ) -> tuple[str, ...]:
        requested = _deduplicate_names(values, material=True)
        requested_keys = {_name_key(item) for item in requested}
        ordered = tuple(
            item for item in self.available_materials if _name_key(item) in requested_keys
        )
        if not ignore_unknown:
            available_keys = {_name_key(item) for item in self.available_materials}
            unknown = [item for item in requested if _name_key(item) not in available_keys]
            if unknown:
                raise ValueError(f"预设包含不存在的材料：{'、'.join(unknown)}。")
        return ordered
