"""Small PySide6/PySide2 compatibility surface used by the shared QML UI."""

from __future__ import annotations


try:
    import shiboken6 as _shiboken
    from PySide6.QtCore import (
        Property,
        QAbstractListModel,
        QCoreApplication,
        QModelIndex,
        QObject,
        QTimer,
        QUrl,
        Qt,
        Signal,
        Slot,
    )
    from PySide6.QtGui import QDesktopServices, QFontDatabase, QGuiApplication, QIcon, QPixmap
    from PySide6.QtQml import QQmlApplicationEngine
    from PySide6.QtWidgets import QApplication, QFileDialog

    QT_API = "PySide6"
    QT_MAJOR = 6
except ImportError:
    import shiboken2 as _shiboken  # type: ignore[no-redef]
    from PySide2.QtCore import (  # type: ignore[no-redef]
        Property,
        QAbstractListModel,
        QCoreApplication,
        QModelIndex,
        QObject,
        QTimer,
        QUrl,
        Qt,
        Signal,
        Slot,
    )
    from PySide2.QtGui import QDesktopServices, QFontDatabase, QGuiApplication, QIcon, QPixmap  # type: ignore[no-redef]
    from PySide2.QtQml import QQmlApplicationEngine  # type: ignore[no-redef]
    from PySide2.QtWidgets import QApplication, QFileDialog  # type: ignore[no-redef]

    QT_API = "PySide2"
    QT_MAJOR = 5


try:
    DISPLAY_ROLE = int(Qt.ItemDataRole.DisplayRole)
    USER_ROLE = int(Qt.ItemDataRole.UserRole)
except AttributeError:
    DISPLAY_ROLE = int(Qt.DisplayRole)
    USER_ROLE = int(Qt.UserRole)


def application_attribute(name: str):
    """Resolve a Qt application attribute across scoped/unscoped enums."""

    enum = getattr(Qt, "ApplicationAttribute", None)
    if enum is not None and hasattr(enum, name):
        return getattr(enum, name)
    return getattr(Qt, name, None)


def constant_property(value_type):
    """Create a read-only constant property without PySide2's copy bug.

    PySide2 5.15.2.1 rejects decorator-style ``Property(...,
    constant=True)`` when it copies the descriptor to attach the getter: its
    copy path passes ``Py_None`` as the setter and then mistakes that value for
    a WRITE method.  Constructing the property once with its getter preserves
    the intended Qt metadata on both PySide2 and PySide6.
    """

    def decorate(getter):
        return Property(value_type, getter, constant=True)

    return decorate


def delete_qobject(value) -> None:
    """Destroy a QML root while its context objects are still valid."""

    _shiboken.delete(value)


__all__ = [
    "Property",
    "QAbstractListModel",
    "QApplication",
    "QCoreApplication",
    "QDesktopServices",
    "QFileDialog",
    "QFontDatabase",
    "QGuiApplication",
    "QModelIndex",
    "QObject",
    "QQmlApplicationEngine",
    "QTimer",
    "QUrl",
    "Qt",
    "Signal",
    "Slot",
    "DISPLAY_ROLE",
    "USER_ROLE",
    "QT_API",
    "QT_MAJOR",
    "application_attribute",
    "constant_property",
    "delete_qobject",
]
