import QtQuick 2.15
import QtQuick.Controls 2.15

CheckBox {
    id: control

    spacing: 7
    implicitHeight: 28
    leftPadding: 0
    rightPadding: 0
    hoverEnabled: true
    focusPolicy: Qt.StrongFocus

    indicator: Rectangle {
        implicitWidth: 16
        implicitHeight: 16
        x: control.leftPadding
        y: (control.height - height) / 2
        color: !control.enabled ? "#F2F0EA" : (control.checked ? "#17715B" : "#FFFFFF")
        border.width: 1
        border.color: !control.enabled ? "#D8D5CB" : (control.checked ? "#17715B" : (control.activeFocus ? "#17715B" : "#B9B6AE"))

        Canvas {
            anchors.fill: parent
            visible: control.checked
            onPaint: {
                var context = getContext("2d")
                context.clearRect(0, 0, width, height)
                context.strokeStyle = "#FFFFFF"
                context.lineWidth = 1.7
                context.lineCap = "round"
                context.lineJoin = "round"
                context.beginPath()
                context.moveTo(3.4, 8.2)
                context.lineTo(6.5, 11.2)
                context.lineTo(12.6, 4.6)
                context.stroke()
            }
        }
    }

    contentItem: Text {
        leftPadding: control.indicator.width + control.spacing
        text: control.text
        color: control.enabled ? "#292825" : "#B3B0A6"
        font.pixelSize: 13
        verticalAlignment: Text.AlignVCenter
        wrapMode: Text.Wrap
    }
}
