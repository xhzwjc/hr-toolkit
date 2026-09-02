import QtQuick 2.15
import QtQuick.Controls 2.15
import QtQuick.Layouts 1.15

Dialog {
    id: control

    property string acceptText: ""
    property string rejectText: ""
    property string closeText: ""
    property bool acceptButtonEnabled: true

    modal: true
    padding: 18
    closePolicy: Popup.CloseOnEscape
    enter: Transition {}
    exit: Transition {}

    Overlay.modal: Rectangle { color: "#66000000" }

    background: Rectangle {
        color: "#FFFFFF"
        radius: 12
        border.color: "#ECEAE4"
        border.width: 1
    }

    header: Rectangle {
        implicitHeight: 54
        color: "#FFFFFF"
        radius: 12

        Text {
            anchors.left: parent.left
            anchors.right: parent.right
            anchors.leftMargin: 18
            anchors.rightMargin: 18
            anchors.verticalCenter: parent.verticalCenter
            text: control.title
            color: "#292825"
            font.pixelSize: 16
            font.weight: Font.DemiBold
            elide: Text.ElideRight
        }

        Rectangle {
            anchors.left: parent.left
            anchors.right: parent.right
            anchors.bottom: parent.bottom
            height: 1
            color: "#F1EFE9"
        }
    }

    footer: Rectangle {
        visible: control.acceptText || control.rejectText || control.closeText
        implicitHeight: visible ? 58 : 0
        color: "#FFFFFF"
        radius: 12

        Rectangle {
            anchors.left: parent.left
            anchors.right: parent.right
            anchors.top: parent.top
            height: 1
            color: "#F1EFE9"
        }

        RowLayout {
            anchors.fill: parent
            anchors.leftMargin: 18
            anchors.rightMargin: 18
            spacing: 8

            Item { Layout.fillWidth: true }
            AppButton {
                visible: control.closeText.length > 0
                text: control.closeText
                onClicked: control.close()
            }
            AppButton {
                visible: control.rejectText.length > 0
                text: control.rejectText
                onClicked: control.reject()
            }
            AppButton {
                visible: control.acceptText.length > 0
                text: control.acceptText
                variant: "primary"
                enabled: control.acceptButtonEnabled
                onClicked: control.accept()
            }
        }
    }
}
