"""Bounded and virtualized list models for the Qt Quick front end."""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from typing import Any

from .compat import QAbstractListModel, QModelIndex, Qt, USER_ROLE


class ObjectListModel(QAbstractListModel):
    """A reset-efficient fixed-role model consumed by QML ``ListView``.

    Only visible delegates are instantiated by ListView, so tens of thousands
    of file records do not create tens of thousands of controls.
    """

    def __init__(self, roles: Iterable[str], parent=None) -> None:
        super().__init__(parent)
        self._roles = tuple(dict.fromkeys(str(role) for role in roles))
        self._role_numbers = {
            USER_ROLE + index + 1: role for index, role in enumerate(self._roles)
        }
        self._role_lookup = {role: number for number, role in self._role_numbers.items()}
        self._items: list[dict[str, Any]] = []

    def roleNames(self):  # noqa: N802 - Qt override
        return {
            number: role.encode("utf-8")
            for number, role in self._role_numbers.items()
        }

    def rowCount(self, parent=QModelIndex()):  # noqa: N802 - Qt override
        if parent.isValid():
            return 0
        return len(self._items)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid() or index.row() < 0 or index.row() >= len(self._items):
            return None
        role_name = self._role_numbers.get(int(role))
        if role_name is None:
            return None
        return self._items[index.row()].get(role_name)

    def set_items(self, items: Iterable[Mapping[str, Any]]) -> None:
        normalized = [
            {role: item.get(role) for role in self._roles}
            for item in items
        ]
        self.beginResetModel()
        self._items = normalized
        self.endResetModel()

    def append(self, item: Mapping[str, Any], *, maximum: int | None = None) -> None:
        if maximum is not None and maximum > 0 and len(self._items) >= maximum:
            remove_count = len(self._items) - maximum + 1
            self.beginRemoveRows(QModelIndex(), 0, remove_count - 1)
            del self._items[:remove_count]
            self.endRemoveRows()
        row = len(self._items)
        self.beginInsertRows(QModelIndex(), row, row)
        self._items.append({role: item.get(role) for role in self._roles})
        self.endInsertRows()

    def append_batch(
        self,
        items: Iterable[Mapping[str, Any]],
        *,
        maximum: int | None = None,
    ) -> None:
        normalized = [
            {role: item.get(role) for role in self._roles}
            for item in items
        ]
        if not normalized:
            return
        if maximum is not None and maximum > 0 and len(normalized) > maximum:
            normalized = normalized[-maximum:]
        if maximum is not None and maximum > 0:
            total_after = len(self._items) + len(normalized)
            if total_after > maximum:
                remove_count = min(total_after - maximum, len(self._items))
                if remove_count > 0:
                    self.beginRemoveRows(QModelIndex(), 0, remove_count - 1)
                    del self._items[:remove_count]
                    self.endRemoveRows()
        row = len(self._items)
        self.beginInsertRows(QModelIndex(), row, row + len(normalized) - 1)
        self._items.extend(normalized)
        self.endInsertRows()

    def remove_at(self, row: int) -> bool:
        if row < 0 or row >= len(self._items):
            return False
        self.beginRemoveRows(QModelIndex(), row, row)
        del self._items[row]
        self.endRemoveRows()
        return True

    def update_at(self, row: int, item: Mapping[str, Any]) -> bool:
        """Update one row without resetting the consuming ``ListView``."""

        if row < 0 or row >= len(self._items):
            return False
        self._items[row] = {role: item.get(role) for role in self._roles}
        model_index = self.index(row, 0)
        self.dataChanged.emit(
            model_index,
            model_index,
            list(self._role_numbers),
        )
        return True

    def splice(
        self,
        row: int,
        remove_count: int = 0,
        items: Iterable[Mapping[str, Any]] = (),
    ) -> None:
        """Replace a contiguous range while preserving view scroll state."""

        start = max(0, min(int(row), len(self._items)))
        removable = max(0, min(int(remove_count), len(self._items) - start))
        if removable:
            self.beginRemoveRows(QModelIndex(), start, start + removable - 1)
            del self._items[start : start + removable]
            self.endRemoveRows()
        normalized = [
            {role: item.get(role) for role in self._roles}
            for item in items
        ]
        if normalized:
            self.beginInsertRows(QModelIndex(), start, start + len(normalized) - 1)
            self._items[start:start] = normalized
            self.endInsertRows()

    def clear(self) -> None:
        if not self._items:
            return
        self.beginResetModel()
        self._items.clear()
        self.endResetModel()

    def item_at(self, row: int) -> dict[str, Any] | None:
        if 0 <= row < len(self._items):
            return dict(self._items[row])
        return None

    def items(self) -> list[dict[str, Any]]:
        return [dict(item) for item in self._items]

    def __len__(self) -> int:
        return len(self._items)


class InputFileModel(ObjectListModel):
    def __init__(self, parent=None) -> None:
        super().__init__(("name", "path", "kind", "detail"), parent)


class LogModel(ObjectListModel):
    def __init__(self, parent=None) -> None:
        super().__init__(("time", "text", "level"), parent)


class WorkspaceModel(ObjectListModel):
    def __init__(self, parent=None) -> None:
        super().__init__(
            ("name", "path", "isDir", "depth", "expanded", "hasChildren", "detail"),
            parent,
        )


class HistoryModel(ObjectListModel):
    def __init__(self, parent=None) -> None:
        super().__init__(
            ("recordId", "time", "tool", "status", "inputs", "outputs", "detail"),
            parent,
        )


class TrashModel(ObjectListModel):
    def __init__(self, parent=None) -> None:
        super().__init__(
            (
                "batchId",
                "title",
                "tool",
                "status",
                "deletedAt",
                "counts",
                "restorePath",
                "size",
            ),
            parent,
        )
