import QtQuick 2.15
import QtQuick.Controls 2.15

TextField {
    id: control

    implicitHeight: 36
    leftPadding: 11
    rightPadding: 11
    topPadding: 7
    bottomPadding: 7
    selectByMouse: true
    color: enabled ? "#292825" : "#B3B0A6"
    placeholderTextColor: "#98958C"
    selectionColor: "#17715B"
    selectedTextColor: "#FFFFFF"
    font.pixelSize: 13

    background: Rectangle {
        radius: 0
        color: control.enabled ? "#FAF9F6" : "#F2F0EA"
        border.width: 1
        border.color: control.activeFocus ? "#17715B" : "#ECEAE4"
    }
}
