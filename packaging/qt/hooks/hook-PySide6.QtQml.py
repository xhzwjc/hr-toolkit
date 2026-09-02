"""Collect only the Qt Quick modules used by HRToolkit's QML scene."""

from PyInstaller.utils.hooks.qt import add_qt6_dependencies, pyside6_library_info


def _required_qml_entry(entry):
    destination = str(entry[1]).replace("\\", "/")
    marker = "/qml/"
    relative = destination.split(marker, 1)[-1] if marker in destination else destination
    exact = {
        "QtCore",
        "QtQml",
        "QtQml/Models",
        "QtQml/WorkerScript",
        "QtQuick",
        "QtQuick/Controls",
        "QtQuick/Controls/impl",
        "QtQuick/Layouts",
        "QtQuick/Templates",
        "QtQuick/Window",
    }
    return relative in exact or relative.startswith("QtQuick/Controls/Basic")


hiddenimports, binaries, datas = add_qt6_dependencies(__file__)
_qml_binaries, _qml_datas = pyside6_library_info.collect_qtqml_files()
binaries += [entry for entry in _qml_binaries if _required_qml_entry(entry)]
datas += [entry for entry in _qml_datas if _required_qml_entry(entry)]
