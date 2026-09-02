import QtQuick 2.15
import QtQuick.Controls 2.15

ComboBox {
    id: control

    implicitHeight: 36
    implicitWidth: 220
    leftPadding: 11
    rightPadding: 34
    topPadding: 7
    bottomPadding: 7
    hoverEnabled: true
    font.pixelSize: 13

    contentItem: Text {
        leftPadding: 0
        rightPadding: 0
        text: control.displayText
        color: control.enabled ? "#292825" : "#B3B0A6"
        font: control.font
        verticalAlignment: Text.AlignVCenter
        elide: Text.ElideRight
    }

    indicator: Canvas {
        x: control.width - width - 11
        y: (control.height - height) / 2
        width: 12
        height: 8
        onPaint: {
            var context = getContext("2d")
            context.clearRect(0, 0, width, height)
            context.strokeStyle = control.enabled ? "#78766E" : "#B3B0A6"
            context.lineWidth = 1.4
            context.lineCap = "round"
            context.lineJoin = "round"
            context.beginPath()
            context.moveTo(1.5, 2)
            context.lineTo(6, 6)
            context.lineTo(10.5, 2)
            context.stroke()
        }
    }

    background: Rectangle {
        color: control.enabled ? "#FAF9F6" : "#F2F0EA"
        border.width: 1
        border.color: control.activeFocus ? "#17715B" : "#ECEAE4"
    }

    delegate: ItemDelegate {
        width: control.width
        height: 34
        highlighted: control.highlightedIndex === index
        contentItem: Text {
            text: control.textRole ? modelData[control.textRole] : modelData
            color: "#292825"
            font.pixelSize: 13
            verticalAlignment: Text.AlignVCenter
            elide: Text.ElideRight
        }
        background: Rectangle {
            color: parent.highlighted ? "#E4EFEA" : (parent.hovered ? "#F0EEE8" : "#FFFFFF")
        }
    }

    popup: Popup {
        y: control.height - 1
        width: control.width
        implicitHeight: Math.min(contentItem.implicitHeight + 2, 260)
        padding: 1
        contentItem: ListView {
            clip: true
            implicitHeight: contentHeight
            model: control.popup.visible ? control.delegateModel : null
            currentIndex: control.highlightedIndex
            ScrollIndicator.vertical: ScrollIndicator {}
        }
        background: Rectangle {
            color: "#FFFFFF"
            border.color: "#ECEAE4"
            border.width: 1
        }
    }
}
