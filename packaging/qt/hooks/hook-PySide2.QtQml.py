"""Collect only the Qt 5 QML modules used by HRToolkit on Windows 7."""

from PyInstaller.utils.hooks.qt import add_qt5_dependencies, pyside2_library_info


def _required_qml_entry(entry):
    destination = str(entry[1]).replace("\\", "/")
    marker = "/qml/"
    relative = destination.split(marker, 1)[-1] if marker in destination else destination
    exact = {
        "QtQml",
        "QtQml/Models.2",
        "QtQml/WorkerScript.2",
        "QtQuick.2",
        "QtQuick/Controls.2",
        "QtQuick/Controls.2/impl",
        "QtQuick/Layouts",
        "QtQuick/Templates.2",
        "QtQuick/Window.2",
    }
    return relative in exact or relative.startswith("QtQuick/Controls.2/Default")


hiddenimports, binaries, datas = add_qt5_dependencies(__file__)
_qml_binaries, _qml_datas = pyside2_library_info.collect_qtqml_files()
binaries += [entry for entry in _qml_binaries if _required_qml_entry(entry)]
datas += [entry for entry in _qml_datas if _required_qml_entry(entry)]
hiddenimports += ["PySide2.QtGui"]
