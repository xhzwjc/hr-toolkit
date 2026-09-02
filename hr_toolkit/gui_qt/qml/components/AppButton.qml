import QtQuick 2.15
import QtQuick.Controls 2.15

Button {
    id: control
    property string variant: "secondary"
    property color primaryColor: "#17715B"
    property color primaryPressedColor: "#125E4B"
    property color textColor: variant === "primary" ? "#FFFFFF" : (variant === "link" ? primaryColor : "#292825")

    implicitHeight: variant === "link" ? 30 : 34
    implicitWidth: Math.max(variant === "link" ? 52 : 82, contentItem.implicitWidth + (variant === "link" ? 12 : 26))
    leftPadding: variant === "link" ? 6 : 13
    rightPadding: variant === "link" ? 6 : 13
    hoverEnabled: true
    focusPolicy: Qt.StrongFocus

    contentItem: Text {
        text: control.text
        color: control.enabled ? control.textColor : "#B3B0A6"
        font.pixelSize: 13
        font.weight: control.variant === "primary" ? Font.DemiBold : Font.Normal
        horizontalAlignment: Text.AlignHCenter
        verticalAlignment: Text.AlignVCenter
        elide: Text.ElideRight
    }

    background: Rectangle {
        radius: 8
        color: {
            if (!control.enabled)
                return control.variant === "link" ? "transparent" : "#F2F0EA"
            if (control.variant === "primary")
                return control.down ? control.primaryPressedColor : (control.hovered ? "#1B7B63" : control.primaryColor)
            if (control.variant === "tonal")
                return control.down ? "#DCE8E3" : (control.hovered ? "#E0ECE7" : "#E4EFEA")
            if (control.variant === "link")
                return control.down ? "#E4EFEA" : (control.hovered ? "#F0EEE8" : "transparent")
            return control.down ? "#EBE8E1" : (control.hovered ? "#F2F0EA" : "#FFFFFF")
        }
        border.color: control.variant === "secondary" ? "#ECEAE4" : "transparent"
        border.width: control.variant === "secondary" ? 1 : 0
    }
}
