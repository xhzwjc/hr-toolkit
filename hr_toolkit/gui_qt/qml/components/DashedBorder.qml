import QtQuick 2.15

Canvas {
    id: borderCanvas

    property color strokeColor: "#D0CBC0"
    property real cornerRadius: 12

    renderTarget: Canvas.Image
    renderStrategy: Canvas.Cooperative

    onStrokeColorChanged: requestPaint()
    onCornerRadiusChanged: requestPaint()
    onWidthChanged: requestPaint()
    onHeightChanged: requestPaint()

    onPaint: {
        var context = getContext("2d")
        context.clearRect(0, 0, width, height)
        context.strokeStyle = strokeColor
        context.lineWidth = 1
        context.setLineDash([5, 4])
        context.beginPath()
        context.roundedRect(0.5, 0.5, Math.max(0, width - 1), Math.max(0, height - 1), cornerRadius, cornerRadius)
        context.stroke()
    }
}
