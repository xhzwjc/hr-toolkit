import QtQuick 2.15
import QtQuick.Layouts 1.15

Rectangle {
    id: progressPanel
    property int completed: 0
    property int total: 0
    property string message: ""
    property int elapsedSeconds: 0
    property int waitSeconds: 0
    property bool active: false
    property real fraction: total > 0 ? Math.max(0, Math.min(1, completed / total)) : 0
    implicitHeight: progressContent.implicitHeight + 28
    radius: 8
    color: "#F3F7F4"
    border.color: "#DCE6DF"

    ColumnLayout {
        id: progressContent
        anchors.left: parent.left
        anchors.right: parent.right
        anchors.top: parent.top
        anchors.margins: 14
        spacing: 8
        RowLayout {
            Layout.fillWidth: true
            Text { text: "当前阶段进度"; color: "#253B30"; font.pixelSize: 13; font.bold: true }
            Item { Layout.fillWidth: true }
            Text {
                objectName: "materialProgressCount"
                text: progressPanel.total > 0 ? progressPanel.completed + "/" + progressPanel.total + "（" + Math.floor(progressPanel.fraction * 100) + "%）" : "正在确定工作量"
                color: "#39624A"
                font.pixelSize: 12
            }
        }
        Rectangle {
            Layout.fillWidth: true
            height: 7
            radius: 3
            color: "#DEE7E0"
            Rectangle {
                objectName: "materialProgressFill"
                width: parent.width * progressPanel.fraction
                height: parent.height
                radius: 3
                color: "#287950"
                // 无计时器补进度或插值动画；宽度仅由业务完成数驱动。
            }
        }
        Text {
            Layout.fillWidth: true
            text: progressPanel.message
            textFormat: Text.PlainText
            wrapMode: Text.Wrap
            color: "#344C3D"
            font.pixelSize: 12
        }
        Text {
            Layout.fillWidth: true
            text: "已用 " + progressPanel.elapsedSeconds + " 秒" + (progressPanel.active ? " · 距上次进度更新 " + progressPanel.waitSeconds + " 秒；单份资料识别期间计数保持不变" : "")
            textFormat: Text.PlainText
            wrapMode: Text.Wrap
            color: "#708176"
            font.pixelSize: 11
        }
    }
}
