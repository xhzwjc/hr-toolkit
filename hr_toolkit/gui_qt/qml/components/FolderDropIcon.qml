import QtQuick 2.15

Rectangle {
    id: root

    implicitWidth: 34
    implicitHeight: 34
    radius: 9
    color: "#E4EFEA"

    Canvas {
        anchors.centerIn: parent
        width: 18
        height: 18
        renderTarget: Canvas.Image
        renderStrategy: Canvas.Cooperative
        onPaint: {
            var context = getContext("2d")
            context.clearRect(0, 0, width, height)
            context.strokeStyle = "#17715B"
            context.lineWidth = 1.7
            context.lineCap = "round"
            context.lineJoin = "round"
            context.beginPath()
            context.moveTo(2.5, 14.5)
            context.lineTo(2.5, 5.5)
            context.lineTo(5.2, 5.5)
            context.lineTo(6.7, 3.5)
            context.lineTo(11.0, 3.5)
            context.lineTo(12.5, 5.5)
            context.lineTo(15.5, 5.5)
            context.lineTo(15.5, 14.5)
            context.closePath()
            context.stroke()
        }
    }
}
