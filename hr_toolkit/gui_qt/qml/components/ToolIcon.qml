import QtQuick 2.15

Canvas {
    id: icon

    property string iconId: ""
    property color strokeColor: "#55534C"
    property real lineWidth: 1.25

    implicitWidth: 16
    implicitHeight: 16
    renderTarget: Canvas.Image
    renderStrategy: Canvas.Cooperative

    onIconIdChanged: requestPaint()
    onStrokeColorChanged: requestPaint()
    onLineWidthChanged: requestPaint()
    onWidthChanged: requestPaint()
    onHeightChanged: requestPaint()

    onPaint: {
        var context = getContext("2d")
        context.clearRect(0, 0, width, height)
        var scale = Math.min(width, height) / 14.0
        var offsetX = (width - 14.0 * scale) / 2.0
        var offsetY = (height - 14.0 * scale) / 2.0
        function px(value) { return offsetX + value * scale }
        function py(value) { return offsetY + value * scale }
        function line(points) {
            context.beginPath()
            context.moveTo(px(points[0]), py(points[1]))
            for (var i = 2; i < points.length; i += 2)
                context.lineTo(px(points[i]), py(points[i + 1]))
            context.stroke()
        }
        function oval(left, top, right, bottom) {
            context.beginPath()
            context.ellipse(px(left), py(top), (right - left) * scale, (bottom - top) * scale)
            context.stroke()
        }
        context.strokeStyle = strokeColor
        context.lineWidth = lineWidth
        context.lineCap = "round"
        context.lineJoin = "round"
        if (iconId === "social_security") {
            context.strokeRect(px(2.5), py(1.5), 9 * scale, 11 * scale)
            line([5, 5, 9, 5]); line([5, 8, 9, 8])
        } else if (iconId === "insurance_ledger") {
            oval(1.5, 1.5, 12.5, 12.5); line([4.6, 7, 6.3, 8.7, 9.6, 5.4])
        } else if (iconId === "data_statistics") {
            line([2.5, 12, 2.5, 7]); line([7, 12, 7, 2.5]); line([11.5, 12, 11.5, 5])
        } else if (iconId === "salary_split") {
            line([7, 2, 7, 6]); line([7, 6, 3, 11]); line([7, 6, 11, 11])
        } else if (iconId === "salary_merge") {
            line([3, 3, 7, 8]); line([11, 3, 7, 8]); line([7, 8, 7, 12])
        } else if (iconId === "personnel_change_merge") {
            line([2, 4.5, 10, 4.5]); line([8, 2, 10.5, 4.5, 8, 7])
            line([12, 9.5, 4, 9.5]); line([6, 7, 3.5, 9.5, 6, 12])
        } else if (iconId === "archive_import") {
            context.strokeRect(px(2), py(4.5), 10 * scale, 7.5 * scale)
            line([2, 7, 12, 7]); line([7, 4.5, 7, 2.5])
        } else if (iconId === "material_collector") {
            context.strokeRect(px(2.5), py(3), 9 * scale, 8.5 * scale)
            line([2.5, 6, 11.5, 6]); line([7, 6, 7, 11.5])
        } else if (iconId === "folder_rename") {
            line([2, 10.5, 2, 4, 3.5, 2.5, 6, 2.5, 7.2, 4, 10.5, 4, 12, 5.5, 12, 10.5, 10.5, 12, 3.5, 12, 2, 10.5])
        } else if (iconId === "tutorial") {
            oval(1.5, 1.5, 12.5, 12.5); line([7, 6.5, 7, 10]); line([7, 4, 7, 4.45])
        } else if (iconId === "clock") {
            oval(1.5, 1.5, 12.5, 12.5); line([7, 4, 7, 7.2, 9, 8.6])
        } else if (iconId === "plus_circle") {
            oval(1.5, 1.5, 12.5, 12.5); line([7, 4.2, 7, 9.8]); line([4.2, 7, 9.8, 7])
        } else if (iconId === "folder_open") {
            line([1.7, 11.5, 2.4, 5.4, 5.3, 5.4, 6.5, 3.5, 11.9, 3.5, 12.4, 5.4])
            line([2.4, 6.2, 12.4, 6.2, 11, 11.5, 1.7, 11.5])
        } else if (iconId === "chevron_down") {
            line([3.2, 5.2, 7, 8.9, 10.8, 5.2])
        } else if (iconId === "run_log") {
            context.strokeRect(px(2.4), py(1.8), 9.2 * scale, 10.4 * scale)
            line([4.5, 5, 9.5, 5]); line([4.5, 7.4, 9.5, 7.4]); line([4.5, 9.8, 8, 9.8])
        } else {
            oval(3, 3, 11, 11)
        }
    }
}
