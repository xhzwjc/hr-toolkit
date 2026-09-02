import QtQuick 2.15
import QtQuick.Controls 2.15

Button {
    id: control
    property string variant: "secondary"
    property color primaryColor: "#0F766E"
    property color primaryPressedColor: "#0A5C56"
    property color textColor: variant === "primary" ? "#FFFFFF" : (variant === "link" ? primaryColor : "#28312E")

    implicitHeight: 38
    implicitWidth: Math.max(82, contentItem.implicitWidth + 28)
    leftPadding: 14
    rightPadding: 14
    hoverEnabled: true

    contentItem: Text {
        text: control.text
        color: control.enabled ? control.textColor : "#9A9C99"
        font.pixelSize: 14
        font.weight: control.variant === "primary" ? Font.DemiBold : Font.Normal
        horizontalAlignment: Text.AlignHCenter
        verticalAlignment: Text.AlignVCenter
        elide: Text.ElideRight
    }

    background: Rectangle {
        radius: 9
        color: {
            if (!control.enabled)
                return "#F0EFEB"
            if (control.variant === "primary")
                return control.down ? control.primaryPressedColor : (control.hovered ? "#128277" : control.primaryColor)
            if (control.variant === "link")
                return control.down ? "#DCEBE7" : (control.hovered ? "#EAF3F0" : "transparent")
            return control.down ? "#E8E6E0" : (control.hovered ? "#F4F2ED" : "#FAF9F6")
        }
        border.color: control.variant === "secondary" ? "#DEDDD8" : "transparent"
        border.width: control.variant === "secondary" ? 1 : 0
    }
}
