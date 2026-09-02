import QtQuick 2.15
import QtQuick.Controls 2.15
import QtQuick.Layouts 1.15
import "components"

ApplicationWindow {
    id: root
    objectName: "mainWindow"
    width: 1180
    height: 840
    minimumWidth: 760
    minimumHeight: 600
    visible: true
    color: "#F7F5F1"
    title: "HR Workbench v" + controller.appVersion

    readonly property color primary: "#0F766E"
    readonly property color textMain: "#222826"
    readonly property color textMuted: "#727875"
    readonly property color border: "#E6E3DD"
    // Keep the navigation geometry stable while the native window is dragged.
    // Hysteresis makes the responsive mode switch at most once in either
    // direction instead of repeatedly rebuilding both sides of the layout
    // around one breakpoint.
    property bool compactSidebar: false

    // During a live native resize on Windows the DWM compositor can stall if
    // QML re-evaluates every binding that depends on root.width/height on
    // each pixel change.  We therefore sample the size through a timer and
    // let the rest of the UI bind to the *settled* values instead of the raw
    // window geometry.  This mirrors what Electron (Codex / Claude / GitHub
    // Desktop) does internally: defer layout until WM_SIZE stops firing.
    property int settledWidth: width
    property int settledHeight: height
    Timer {
        id: settleTimer
        interval: 32   // ~2 frames; fast enough to feel instant
        running: false
        repeat: false
        onTriggered: {
            root.settledWidth = root.width
            root.settledHeight = root.height
        }
    }
    onWidthChanged: settleTimer.restart()
    onHeightChanged: settleTimer.restart()

    function updateResponsiveMode() {
        if (!compactSidebar && settledWidth <= 860)
            compactSidebar = true
        else if (compactSidebar && settledWidth >= 980)
            compactSidebar = false
    }

    onSettledWidthChanged: updateResponsiveMode()

    onClosing: function(closeEvent) {
        closeEvent.accepted = controller.requestClose()
    }
    Component.onCompleted: {
        updateResponsiveMode()
        controller.start()
    }

    RowLayout {
        anchors.fill: parent
        spacing: 0

        Rectangle {
            id: sidebar
            objectName: "sidebar"
            Layout.fillHeight: true
            Layout.preferredWidth: root.compactSidebar ? 76 : 248
            Layout.minimumWidth: Layout.preferredWidth
            color: "#F2F0EB"
            border.color: root.border

            ColumnLayout {
                anchors.fill: parent
                anchors.margins: root.compactSidebar ? 10 : 15
                spacing: 8

                Item {
                    Layout.fillWidth: true
                    Layout.preferredHeight: 76
                    RowLayout {
                        anchors.fill: parent
                        spacing: 10
                        Rectangle {
                            Layout.preferredWidth: 32
                            Layout.preferredHeight: 32
                            radius: 10
                            color: "#E5EFEA"
                            Text {
                                anchors.centerIn: parent
                                text: "✣"
                                color: root.primary
                                font.pixelSize: 20
                            }
                        }
                        ColumnLayout {
                            visible: !root.compactSidebar
                            Layout.fillWidth: true
                            spacing: 2
                            Text { text: "HR Workbench"; color: root.textMain; font.pixelSize: 15; font.weight: Font.DemiBold }
                            Text { text: "人资运营自动化"; color: root.textMuted; font.pixelSize: 12 }
                        }
                    }
                }

                Card {
                    Layout.fillWidth: true
                    Layout.preferredHeight: root.compactSidebar ? 54 : 104
                    color: "#FBFAF7"
                    ColumnLayout {
                        anchors.fill: parent
                        anchors.margins: root.compactSidebar ? 8 : 12
                        spacing: 5
                        Text {
                            Layout.fillWidth: true
                            text: root.compactSidebar ? "项目" : (controller.hasProject ? controller.projectName : "工作项目")
                            color: root.textMain
                            font.pixelSize: root.compactSidebar ? 11 : 14
                            font.weight: Font.DemiBold
                            horizontalAlignment: root.compactSidebar ? Text.AlignHCenter : Text.AlignLeft
                            elide: Text.ElideRight
                        }
                        Text {
                            visible: !root.compactSidebar
                            Layout.fillWidth: true
                            text: controller.hasProject ? (controller.projectWritable ? "当前项目 · 可写" : "当前项目 · 只读") : "新建或打开项目后开始处理"
                            color: controller.projectWritable ? root.primary : root.textMuted
                            font.pixelSize: 11
                            elide: Text.ElideRight
                        }
                        RowLayout {
                            visible: !root.compactSidebar
                            Layout.fillWidth: true
                            spacing: 6
                            AppButton { Layout.fillWidth: true; text: "新建项目"; implicitHeight: 30; onClicked: controller.requestCreateProject() }
                            AppButton { Layout.fillWidth: true; text: "打开项目"; implicitHeight: 30; onClicked: controller.openProjectDialog() }
                        }
                    }
                    MouseArea {
                        anchors.fill: parent
                        visible: root.compactSidebar
                        cursorShape: Qt.PointingHandCursor
                        onClicked: projectMenu.open()
                    }
                }

                AppButton {
                    visible: !root.compactSidebar
                    Layout.fillWidth: true
                    text: "最近项目"
                    variant: "link"
                    enabled: controller.recentProjects.length > 0
                    onClicked: projectMenu.open()
                }

                ScrollView {
                    id: navScroll
                    Layout.fillWidth: true
                    Layout.fillHeight: true
                    clip: true
                    ScrollBar.horizontal.policy: ScrollBar.AlwaysOff
                    contentWidth: availableWidth

                    Column {
                        width: navScroll.availableWidth
                        spacing: 9
                        Repeater {
                            model: controller.navGroups
                            delegate: Column {
                                width: parent.width
                                spacing: 3
                                Text {
                                    visible: !root.compactSidebar
                                    width: parent.width
                                    leftPadding: 7
                                    text: modelData.name
                                    color: "#979994"
                                    font.pixelSize: 11
                                    font.weight: Font.DemiBold
                                }
                                Repeater {
                                    model: modelData.items
                                    delegate: Rectangle {
                                        width: parent.width
                                        height: 34
                                        radius: 8
                                        color: controller.currentTool === modelData.id ? "#E3E6E0" : (navMouse.containsMouse ? "#EAE8E3" : "transparent")
                                        Row {
                                            anchors.fill: parent
                                            anchors.leftMargin: root.compactSidebar ? 0 : 9
                                            spacing: 8
                                            Text {
                                                width: root.compactSidebar ? parent.width : 18
                                                height: parent.height
                                                text: modelData.label.slice(0, 1)
                                                color: controller.currentTool === modelData.id ? root.primary : "#5F6562"
                                                font.pixelSize: 13
                                                font.weight: Font.DemiBold
                                                horizontalAlignment: Text.AlignHCenter
                                                verticalAlignment: Text.AlignVCenter
                                            }
                                            Text {
                                                visible: !root.compactSidebar
                                                width: parent.width - 34
                                                height: parent.height
                                                text: modelData.label
                                                color: controller.currentTool === modelData.id ? root.primary : "#464C49"
                                                font.pixelSize: 13
                                                font.weight: controller.currentTool === modelData.id ? Font.DemiBold : Font.Normal
                                                verticalAlignment: Text.AlignVCenter
                                                elide: Text.ElideRight
                                            }
                                        }
                                        MouseArea {
                                            id: navMouse
                                            anchors.fill: parent
                                            hoverEnabled: true
                                            cursorShape: Qt.PointingHandCursor
                                            onClicked: controller.selectTool(modelData.id)
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                Rectangle { Layout.fillWidth: true; Layout.preferredHeight: 1; color: root.border }
                AppButton {
                    Layout.fillWidth: true
                    text: root.compactSidebar ? "记" : "旧版记录"
                    variant: "link"
                    onClicked: {
                        controller.requestHistory()
                        historyDrawer.open()
                    }
                }
                AppButton {
                    Layout.fillWidth: true
                    text: root.compactSidebar ? "助" : "使用教程"
                    variant: "link"
                    onClicked: helpDialog.open()
                }
                AppButton {
                    Layout.fillWidth: true
                    text: root.compactSidebar ? "志" : "打开运行日志"
                    variant: "link"
                    onClicked: controller.openRunLog()
                }
                Text {
                    Layout.fillWidth: true
                    visible: !root.compactSidebar
                    text: "v" + controller.appVersion + "  ·  本地处理，不上传数据"
                    color: "#959894"
                    font.pixelSize: 10
                    horizontalAlignment: Text.AlignHCenter
                }
            }
        }

        Item {
            id: mainPane
            objectName: "mainPane"
            Layout.fillWidth: true
            Layout.fillHeight: true

            ColumnLayout {
                id: mainLayout
                objectName: "mainLayout"
                anchors.fill: parent
                anchors.leftMargin: 28
                anchors.rightMargin: 28
                anchors.topMargin: 24
                anchors.bottomMargin: 14
                spacing: 14

                RowLayout {
                    Layout.fillWidth: true
                    spacing: 10
                    ColumnLayout {
                        Layout.fillWidth: true
                        spacing: 4
                        Text { text: controller.toolGroup; color: root.primary; font.pixelSize: 12; font.weight: Font.DemiBold }
                        Text {
                            Layout.fillWidth: true
                            text: controller.toolTitle
                            color: root.textMain
                            font.pixelSize: 27
                            font.weight: Font.Bold
                            wrapMode: Text.Wrap
                        }
                        Text {
                            Layout.fillWidth: true
                            text: controller.toolDescription
                            color: root.textMuted
                            font.pixelSize: 13
                            wrapMode: Text.Wrap
                        }
                    }
                    AppButton {
                        id: workspaceButton
                        objectName: "workspaceButton"
                        text: "项目文件"
                        enabled: controller.hasProject
                        onClicked: {
                            controller.setWorkspaceExpanded(true)
                            workspaceDrawer.open()
                        }
                    }
                    AppButton { text: controller.updateBusy ? "检查中…" : "检查更新"; enabled: !controller.updateBusy; onClicked: controller.requestUpdateCheck() }
                }

                RowLayout {
                    Layout.fillWidth: true
                    visible: controller.variants.length > 0
                    spacing: 8
                    Repeater {
                        model: controller.variants
                        delegate: AppButton {
                            text: modelData.label
                            variant: controller.currentVariant === modelData.id ? "primary" : "secondary"
                            onClicked: controller.selectVariant(modelData.id)
                        }
                    }
                    Item { Layout.fillWidth: true }
                }

                ScrollView {
                    id: mainScroll
                    objectName: "mainScroll"
                    Layout.fillWidth: true
                    Layout.fillHeight: true
                    clip: true
                    contentWidth: availableWidth
                    ScrollBar.horizontal.policy: ScrollBar.AlwaysOff
                    ScrollBar.vertical.policy: ScrollBar.AsNeeded
                    ScrollBar.vertical.interactive: true

                    ColumnLayout {
                        width: mainScroll.availableWidth
                        spacing: 14

                        Card {
                            Layout.fillWidth: true
                            Layout.preferredHeight: uploadColumn.implicitHeight + 34
                            ColumnLayout {
                                id: uploadColumn
                                anchors.fill: parent
                                anchors.margins: 17
                                spacing: 10
                                RowLayout {
                                    Layout.fillWidth: true
                                    Text { text: controller.inputLabel; color: root.textMain; font.pixelSize: 15; font.weight: Font.DemiBold }
                                    Text { Layout.fillWidth: true; text: controller.inputHint; color: root.textMuted; font.pixelSize: 11; wrapMode: Text.Wrap }
                                    AppButton { visible: controller.inputAllowsFiles; text: controller.inputAllowsMultiple ? "选择文件" : "选择文件"; variant: "link"; onClicked: controller.chooseInputFiles() }
                                    AppButton { visible: controller.inputAllowsFolder; text: "选择文件夹"; variant: "link"; onClicked: controller.chooseInputFolder() }
                                }
                                Rectangle {
                                    Layout.fillWidth: true
                                    Layout.preferredHeight: inputList.count > 0 ? Math.min(260, Math.max(54, inputList.count * 46)) : 100
                                    radius: 12
                                    color: "#FCFBF8"
                                    border.width: 1
                                    border.color: "#D8D4CB"

                                    Column {
                                        anchors.centerIn: parent
                                        visible: inputList.count === 0
                                        spacing: 7
                                        Text { anchors.horizontalCenter: parent.horizontalCenter; text: "↑"; color: root.primary; font.pixelSize: 23 }
                                        Text { anchors.horizontalCenter: parent.horizontalCenter; text: controller.inputDropTitle; color: root.textMain; font.pixelSize: 13; font.weight: Font.DemiBold }
                                        Text { anchors.horizontalCenter: parent.horizontalCenter; text: "点击右上方按钮选择资料"; color: root.textMuted; font.pixelSize: 11 }
                                    }

                                    ListView {
                                        id: inputList
                                        objectName: "inputList"
                                        anchors.fill: parent
                                        anchors.margins: 7
                                        visible: count > 0
                                        clip: true
                                        model: controller.inputModel
                                        reuseItems: true
                                        cacheBuffer: 92
                                        property int activeDelegateCount: 0
                                        spacing: 2
                                        ScrollBar.vertical: ScrollBar { policy: inputList.contentHeight > inputList.height ? ScrollBar.AsNeeded : ScrollBar.AlwaysOff }
                                        delegate: Rectangle {
                                            Component.onCompleted: inputList.activeDelegateCount += 1
                                            Component.onDestruction: inputList.activeDelegateCount -= 1
                                            width: inputList.width
                                            height: 44
                                            radius: 8
                                            color: fileMouse.containsMouse ? "#F2F5F2" : "transparent"
                                            RowLayout {
                                                anchors.fill: parent
                                                anchors.leftMargin: 10
                                                anchors.rightMargin: 6
                                                spacing: 9
                                                Rectangle {
                                                    Layout.preferredWidth: 30; Layout.preferredHeight: 26; radius: 6
                                                    color: kind === "folder" ? "#E7EFEA" : "#EAF0F5"
                                                    Text { anchors.centerIn: parent; text: kind === "folder" ? "夹" : detail.slice(0, 3); color: kind === "folder" ? root.primary : "#557087"; font.pixelSize: 10; font.weight: Font.DemiBold }
                                                }
                                                ColumnLayout {
                                                    Layout.fillWidth: true; spacing: 1
                                                    Text { Layout.fillWidth: true; text: name; color: root.textMain; font.pixelSize: 12; elide: Text.ElideMiddle }
                                                    Text { Layout.fillWidth: true; text: path; color: root.textMuted; font.pixelSize: 10; elide: Text.ElideMiddle }
                                                }
                                                AppButton { text: "移除"; variant: "link"; implicitWidth: 54; implicitHeight: 30; onClicked: controller.removeInput(index) }
                                            }
                                            MouseArea { id: fileMouse; anchors.fill: parent; hoverEnabled: true; acceptedButtons: Qt.NoButton }
                                        }
                                    }

                                    MouseArea {
                                        anchors.fill: parent
                                        visible: inputList.count === 0
                                        cursorShape: Qt.PointingHandCursor
                                        onClicked: {
                                            if (controller.inputAllowsFiles && controller.inputAllowsFolder)
                                                addInputPopup.open()
                                            else if (controller.inputAllowsFolder)
                                                controller.chooseInputFolder()
                                            else
                                                controller.chooseInputFiles()
                                        }
                                    }

                                    Popup {
                                        id: addInputPopup
                                        x: Math.max(8, (parent.width - width) / 2)
                                        y: Math.max(8, (parent.height - height) / 2)
                                        width: 230
                                        padding: 8
                                        closePolicy: Popup.CloseOnEscape | Popup.CloseOnPressOutside
                                        background: Rectangle { radius: 10; color: "#FFFFFF"; border.color: root.border }
                                        contentItem: ColumnLayout {
                                            spacing: 4
                                            AppButton { Layout.fillWidth: true; text: "添加文件 / 压缩包"; onClicked: { addInputPopup.close(); controller.chooseInputFiles() } }
                                            AppButton { Layout.fillWidth: true; text: "添加文件夹"; onClicked: { addInputPopup.close(); controller.chooseInputFolder() } }
                                        }
                                    }
                                }
                            }
                        }

                        Card {
                            Layout.fillWidth: true
                            Layout.preferredHeight: formColumn.implicitHeight + 34
                            ColumnLayout {
                                id: formColumn
                                anchors.fill: parent
                                anchors.margins: 17
                                spacing: 12

                                RowLayout {
                                    Layout.fillWidth: true
                                    visible: controller.hasSupportField
                                    spacing: 10
                                    Text { Layout.preferredWidth: 190; text: controller.supportLabel; color: root.textMain; font.pixelSize: 13 }
                                    Rectangle {
                                        Layout.fillWidth: true
                                        Layout.preferredHeight: 38
                                        radius: 7
                                        color: "#FAF9F6"
                                        border.color: root.border
                                        Text { anchors.fill: parent; anchors.leftMargin: 10; anchors.rightMargin: 10; text: controller.supportPath || "未选择"; color: controller.supportPath ? root.textMain : "#A0A29F"; font.pixelSize: 12; verticalAlignment: Text.AlignVCenter; elide: Text.ElideMiddle }
                                    }
                                    AppButton { text: controller.supportButtonText; variant: "link"; onClicked: controller.chooseSupportFile() }
                                    AppButton { visible: controller.supportAllowsFolder; text: "选择文件夹"; variant: "link"; onClicked: controller.chooseSupportFolder() }
                                    AppButton { visible: !!controller.supportPath; text: "清除"; variant: "link"; onClicked: controller.clearSupport() }
                                }

                                Repeater {
                                    model: controller.formFields
                                    delegate: Loader {
                                        Layout.fillWidth: true
                                        visible: modelData.visible
                                        active: visible
                                        property var field: modelData
                                        sourceComponent: field.kind === "text" ? textFieldComponent
                                                       : field.kind === "choice" ? choiceFieldComponent
                                                       : field.kind === "check" ? checkFieldComponent
                                                       : field.kind === "date_range" ? dateRangeFieldComponent
                                                       : field.kind === "materials" ? materialsFieldComponent
                                                       : null
                                    }
                                }

                                Rectangle { Layout.fillWidth: true; Layout.preferredHeight: 1; color: root.border }
                                RowLayout {
                                    Layout.fillWidth: true
                                    visible: controller.currentTool !== "folder_rename"
                                    spacing: 10
                                    Text { text: "结果位置"; color: root.textMain; font.pixelSize: 13 }
                                    Text { Layout.fillWidth: true; text: controller.hasProject ? "当前项目 / 本次处理结果" : "请先新建或打开工作项目"; color: controller.hasProject ? root.primary : root.textMuted; font.pixelSize: 12; elide: Text.ElideMiddle }
                                    AppButton { text: "打开项目"; variant: "link"; enabled: controller.hasProject; onClicked: controller.openProjectFolder() }
                                }
                            }
                        }

                        RowLayout {
                            Layout.fillWidth: true
                            spacing: 12
                            AppButton {
                                text: controller.runButtonText
                                variant: "primary"
                                enabled: controller.hasProject && !controller.workspaceBusy
                                implicitWidth: 132
                                implicitHeight: 42
                                onClicked: controller.runOrCancel()
                            }
                            AppButton { text: "打开结果目录"; enabled: controller.canOpenLastResult; onClicked: controller.openLastResult() }
                            Text { visible: !!controller.lastRunText; text: controller.lastRunText; color: root.textMuted; font.pixelSize: 12 }
                            Item { Layout.fillWidth: true }
                        }

                        Card {
                            Layout.fillWidth: true
                            Layout.preferredHeight: 250
                            ColumnLayout {
                                anchors.fill: parent
                                anchors.margins: 17
                                spacing: 8
                                Text { text: "运行记录"; color: root.textMain; font.pixelSize: 15; font.weight: Font.DemiBold }
                                ListView {
                                    id: logList
                                    objectName: "logList"
                                    Layout.fillWidth: true
                                    Layout.fillHeight: true
                                    clip: true
                                    model: controller.logModel
                                    reuseItems: true
                                    cacheBuffer: 120
                                    spacing: 3
                                    onCountChanged: positionViewAtEnd()
                                    ScrollBar.vertical: ScrollBar { policy: ScrollBar.AsNeeded }
                                    delegate: RowLayout {
                                        width: logList.width
                                        height: Math.max(25, logText.implicitHeight + 4)
                                        spacing: 7
                                        Text { text: level === "muted" ? "" : "●"; color: level === "error" ? "#C83A3A" : level === "warning" ? "#C28112" : level === "success" ? "#1D8E68" : root.primary; font.pixelSize: 9 }
                                        Text { text: time; color: "#9A9D99"; font.pixelSize: 10; verticalAlignment: Text.AlignTop }
                                        Text { id: logText; Layout.fillWidth: true; text: model.text; color: level === "muted" ? root.textMuted : root.textMain; font.pixelSize: 12; wrapMode: Text.Wrap; textFormat: Text.PlainText }
                                    }
                                }
                            }
                        }
                        Item { Layout.fillWidth: true; Layout.preferredHeight: 4 }
                    }
                }
            }
        }
    }

    Popup {
        id: projectMenu
        x: Math.min(root.settledWidth - width - 12, sidebar.width + 8)
        y: 112
        width: 250
        padding: 9
        closePolicy: Popup.CloseOnEscape | Popup.CloseOnPressOutside
        background: Rectangle { radius: 11; color: "#FFFFFF"; border.color: root.border }
        contentItem: ColumnLayout {
            spacing: 4
            AppButton {
                Layout.fillWidth: true
                text: "新建工作项目"
                onClicked: { projectMenu.close(); controller.requestCreateProject() }
            }
            AppButton {
                Layout.fillWidth: true
                text: "打开已有项目"
                onClicked: { projectMenu.close(); controller.openProjectDialog() }
            }
            Rectangle {
                visible: controller.recentProjects.length > 0
                Layout.fillWidth: true
                Layout.preferredHeight: 1
                color: root.border
            }
            Text {
                visible: controller.recentProjects.length > 0
                Layout.fillWidth: true
                text: "最近项目"
                color: root.textMuted
                font.pixelSize: 11
                leftPadding: 8
            }
            Repeater {
                model: controller.recentProjects
                delegate: AppButton {
                    Layout.fillWidth: true
                    text: modelData.name
                    variant: "link"
                    onClicked: { projectMenu.close(); controller.openProject(modelData.path) }
                }
            }
        }
    }

    Component {
        id: textFieldComponent
        RowLayout {
            spacing: 10
            Text { Layout.preferredWidth: 190; text: field.label; color: root.textMain; font.pixelSize: 13 }
            TextField {
                Layout.fillWidth: true
                text: field.value === undefined || field.value === null ? "" : String(field.value)
                placeholderText: field.placeholder || ""
                selectByMouse: true
                font.pixelSize: 13
                onTextEdited: controller.setFieldValue(field.id, text)
                background: Rectangle { radius: 7; color: "#FAF9F6"; border.color: parent.activeFocus ? root.primary : root.border }
            }
        }
    }

    Component {
        id: choiceFieldComponent
        RowLayout {
            spacing: 10
            Text { Layout.preferredWidth: 190; text: field.label; color: root.textMain; font.pixelSize: 13 }
            ComboBox {
                id: combo
                Layout.preferredWidth: Math.min(360, Math.max(220, implicitWidth))
                model: field.options || []
                textRole: "label"
                currentIndex: {
                    var options = field.options || []
                    for (var i = 0; i < options.length; ++i) {
                        if (String(options[i].value) === String(field.value))
                            return i
                    }
                    return 0
                }
                onActivated: controller.setFieldValue(field.id, field.options[index].value)
            }
            Item { Layout.fillWidth: true }
        }
    }

    Component {
        id: checkFieldComponent
        RowLayout {
            spacing: 10
            Item { Layout.preferredWidth: 190; Layout.preferredHeight: 1 }
            CheckBox {
                text: field.label
                checked: !!field.value
                enabled: !(field.id === "use_ocr_cache" && controller.formFields.some(function(item) { return item.id === "library_mode" && item.value === "flat_ocr" }))
                onToggled: controller.setFieldValue(field.id, checked)
            }
            Item { Layout.fillWidth: true }
        }
    }

    Component {
        id: dateRangeFieldComponent
        RowLayout {
            spacing: 10
            Text { Layout.preferredWidth: 190; text: field.label; color: root.textMain; font.pixelSize: 13; Layout.alignment: Qt.AlignTop }
            ColumnLayout {
                Layout.fillWidth: true
                spacing: 6
                RowLayout {
                    Layout.fillWidth: true
                    spacing: 7
                    TextField {
                        Layout.preferredWidth: 138
                        text: field.startValue === undefined || field.startValue === null ? "" : String(field.startValue)
                        placeholderText: field.startPlaceholder || ""
                        selectByMouse: true
                        onTextEdited: controller.setFieldValue(field.startId, text)
                        background: Rectangle { radius: 7; color: "#FAF9F6"; border.color: parent.activeFocus ? root.primary : root.border }
                    }
                    Text { text: "至"; color: root.textMuted; font.pixelSize: 12 }
                    TextField {
                        Layout.preferredWidth: 138
                        text: field.endValue === undefined || field.endValue === null ? "" : String(field.endValue)
                        placeholderText: field.endPlaceholder || ""
                        selectByMouse: true
                        onTextEdited: controller.setFieldValue(field.endId, text)
                        background: Rectangle { radius: 7; color: "#FAF9F6"; border.color: parent.activeFocus ? root.primary : root.border }
                    }
                    Text { Layout.fillWidth: true; text: field.hint || ""; color: root.textMuted; font.pixelSize: 11; wrapMode: Text.Wrap }
                }
                Flow {
                    Layout.fillWidth: true
                    spacing: 5
                    Repeater {
                        model: field.presets || []
                        delegate: AppButton {
                            text: modelData.label
                            variant: "link"
                            implicitWidth: 58
                            implicitHeight: 28
                            onClicked: controller.applyDatePreset(field.presetGroup, modelData.value)
                        }
                    }
                }
            }
        }
    }

    Component {
        id: materialsFieldComponent
        ColumnLayout {
            spacing: 7
            RowLayout {
                Layout.fillWidth: true
                Text { Layout.preferredWidth: 190; text: "常用组合"; color: root.textMain; font.pixelSize: 13 }
                ComboBox {
                    id: presetCombo
                    Layout.preferredWidth: 190
                    model: controller.materialPresets
                    currentIndex: Math.max(0, controller.materialPresets.indexOf(controller.materialPresetName))
                    onActivated: controller.setMaterialPresetName(currentText)
                }
                AppButton { text: "应用"; variant: "link"; onClicked: controller.applyMaterialPreset(presetCombo.currentText) }
                Item { Layout.fillWidth: true }
            }
            Flow {
                Layout.fillWidth: true
                Layout.leftMargin: 190
                spacing: 4
                AppButton { text: "保存新预设"; variant: "link"; onClicked: controller.requestCreateMaterialPreset() }
                AppButton { text: "更新"; variant: "link"; onClicked: controller.updateMaterialPreset(presetCombo.currentText) }
                AppButton { text: "重命名"; variant: "link"; onClicked: controller.requestRenameMaterialPreset(presetCombo.currentText) }
                AppButton { text: "删除"; variant: "link"; onClicked: controller.requestDeleteMaterialPreset(presetCombo.currentText) }
            }
            RowLayout {
                Layout.fillWidth: true
                Text { Layout.preferredWidth: 190; text: field.label; color: root.textMain; font.pixelSize: 13 }
                AppButton { text: "全选"; variant: "link"; onClicked: controller.selectAllMaterials() }
                AppButton { text: "取消全选"; variant: "link"; onClicked: controller.clearMaterials() }
                Item { Layout.fillWidth: true }
            }
            RowLayout {
                Layout.fillWidth: true
                Item { Layout.preferredWidth: 190; Layout.preferredHeight: 1 }
                AppButton { text: "添加材料"; variant: "link"; onClicked: controller.requestAddCustomMaterial() }
                ComboBox {
                    id: customMaterialCombo
                    visible: controller.customMaterials.length > 0
                    Layout.preferredWidth: 150
                    model: controller.customMaterials
                }
                AppButton {
                    visible: controller.customMaterials.length > 0
                    text: "删除自定义材料"
                    variant: "link"
                    onClicked: controller.requestDeleteCustomMaterial(customMaterialCombo.currentText)
                }
                Item { Layout.fillWidth: true }
            }
            Flow {
                Layout.fillWidth: true
                spacing: 4
                Repeater {
                    model: field.options || []
                    delegate: CheckBox {
                        text: modelData.label
                        checked: modelData.selected
                        onToggled: controller.toggleMaterial(modelData.value, checked)
                    }
                }
            }
        }
    }

    Popup {
        id: workspaceDrawer
        objectName: "workspaceDrawer"
        // Bind position/size to the *settled* dimensions so the popup does
        // not chase the cursor pixel-by-pixel during a live native resize.
        x: root.settledWidth - width
        y: 0
        width: Math.min(440, root.settledWidth - 24)
        height: root.settledHeight
        modal: false
        dim: false
        padding: 0
        closePolicy: Popup.CloseOnEscape | Popup.CloseOnPressOutside
        onClosed: controller.setWorkspaceExpanded(false)
        enter: Transition {}
        exit: Transition {}

        background: Rectangle { color: "#FBFAF7"; border.color: root.border }
        ColumnLayout {
            anchors.fill: parent
            anchors.margins: 14
            spacing: 9
            RowLayout {
                Layout.fillWidth: true
                ColumnLayout {
                    Layout.fillWidth: true; spacing: 2
                    Text { text: "项目文件"; color: root.textMain; font.pixelSize: 18; font.weight: Font.DemiBold }
                    Text { Layout.fillWidth: true; text: controller.projectName; color: root.textMuted; font.pixelSize: 11; elide: Text.ElideMiddle }
                }
                AppButton { text: "关闭"; variant: "link"; onClicked: workspaceDrawer.close() }
            }
            RowLayout {
                Layout.fillWidth: true; spacing: 7
                AppButton { text: "全部文件"; variant: controller.workspaceScope === "all" ? "primary" : "secondary"; onClicked: controller.setWorkspaceScope("all") }
                AppButton { text: "当前功能"; variant: controller.workspaceScope === "tool" ? "primary" : "secondary"; onClicked: controller.setWorkspaceScope("tool") }
                Item { Layout.fillWidth: true }
                AppButton { text: "刷新"; variant: "link"; onClicked: controller.refreshWorkspace() }
            }
            TextField {
                Layout.fillWidth: true
                placeholderText: "按文件名查找"
                selectByMouse: true
                onTextEdited: controller.setWorkspaceSearch(text)
                background: Rectangle { radius: 8; color: "#FFFFFF"; border.color: parent.activeFocus ? root.primary : root.border }
            }
            Flow {
                Layout.fillWidth: true; spacing: 7
                AppButton { text: "导入文件"; enabled: controller.projectWritable && !controller.busy && !controller.workspaceBusy; onClicked: controller.importWorkspaceFiles() }
                AppButton { text: "导入文件夹"; enabled: controller.projectWritable && !controller.busy && !controller.workspaceBusy; onClicked: controller.importWorkspaceFolder() }
                AppButton { text: "移到回收站"; enabled: controller.projectWritable && !controller.busy && !controller.workspaceBusy; variant: "link"; onClicked: controller.requestMoveSelectedBatchToTrash() }
                AppButton { text: "回收站"; enabled: controller.hasProject; variant: "link"; onClicked: { controller.requestProjectTrash(); trashDialog.open() } }
                AppButton { visible: controller.workspaceBusy; text: "取消导入"; variant: "link"; onClicked: controller.cancelWorkspaceImport() }
            }
            ListView {
                id: workspaceList
                objectName: "workspaceList"
                Layout.fillWidth: true
                Layout.fillHeight: true
                clip: true
                model: controller.workspaceModel
                reuseItems: true
                cacheBuffer: 168
                currentIndex: -1
                property int activeDelegateCount: 0
                ScrollBar.vertical: ScrollBar { policy: ScrollBar.AsNeeded }
                delegate: Rectangle {
                    Component.onCompleted: workspaceList.activeDelegateCount += 1
                    Component.onDestruction: workspaceList.activeDelegateCount -= 1
                    width: workspaceList.width
                    height: 42
                    radius: 7
                    color: workspaceList.currentIndex === index ? "#E1ECE8" : (workspaceMouse.containsMouse ? "#F0EEE9" : "transparent")
                    RowLayout {
                        anchors.fill: parent
                        anchors.leftMargin: 7 + depth * 16
                        anchors.rightMargin: 7
                        spacing: 5
                        Text { Layout.preferredWidth: 14; text: isDir ? (expanded ? "▾" : "▸") : ""; color: root.textMuted; font.pixelSize: 11 }
                        Text { Layout.preferredWidth: 22; text: isDir ? "夹" : "文"; color: isDir ? root.primary : "#617381"; font.pixelSize: 11; font.weight: Font.DemiBold; horizontalAlignment: Text.AlignHCenter }
                        ColumnLayout {
                            Layout.fillWidth: true; spacing: 0
                            Text { Layout.fillWidth: true; text: name; color: root.textMain; font.pixelSize: 12; elide: Text.ElideMiddle }
                            Text { Layout.fillWidth: true; text: detail; color: root.textMuted; font.pixelSize: 9; elide: Text.ElideMiddle }
                        }
                        AppButton { z: 1; visible: workspaceList.currentIndex === index; text: "打开"; variant: "link"; implicitWidth: 46; implicitHeight: 28; onClicked: controller.openWorkspaceRow(index) }
                    }
                    MouseArea {
                        id: workspaceMouse
                        anchors.fill: parent
                        hoverEnabled: true
                        acceptedButtons: Qt.LeftButton
                        onClicked: {
                            workspaceList.currentIndex = index
                            controller.selectWorkspaceRow(index)
                            if (isDir)
                                controller.toggleWorkspaceRow(index)
                        }
                        onDoubleClicked: controller.openWorkspaceRow(index)
                    }
                }
            }
            Text {
                Layout.fillWidth: true
                visible: workspaceList.count === 0
                text: controller.hasProject ? "当前范围还没有项目文件" : "请先打开工作项目"
                color: root.textMuted
                font.pixelSize: 12
                horizontalAlignment: Text.AlignHCenter
            }
            AppButton { Layout.fillWidth: true; text: "在系统中打开项目文件夹"; enabled: controller.hasProject; onClicked: controller.openProjectFolder() }
        }
    }

    Drawer {
        id: historyDrawer
        edge: Qt.RightEdge
        width: Math.min(900, Math.max(650, root.settledWidth * 0.82))
        height: root.settledHeight
        modal: true
        interactive: true
        closePolicy: Popup.CloseOnEscape | Popup.CloseOnPressOutside
        background: Rectangle { color: "#FBFAF7"; border.color: root.border }

        ColumnLayout {
            anchors.fill: parent
            anchors.margins: 16
            spacing: 10
            RowLayout {
                Layout.fillWidth: true
                ColumnLayout {
                    Layout.fillWidth: true; spacing: 2
                    Text { text: "旧版记录"; color: root.textMain; font.pixelSize: 19; font.weight: Font.DemiBold }
                    Text { text: "查看升级前保存的上传资料和结果；新处理记录请在项目文件中查看。"; color: root.textMuted; font.pixelSize: 11; wrapMode: Text.Wrap; Layout.fillWidth: true }
                }
                AppButton { text: "关闭"; variant: "link"; onClicked: historyDrawer.close() }
            }
            RowLayout {
                Layout.fillWidth: true; spacing: 7
                TextField {
                    id: historySearch
                    Layout.fillWidth: true
                    placeholderText: "按功能或文件名查找"
                    selectByMouse: true
                    onAccepted: controller.refreshHistory(text, historyTool.currentValue, historyDate.currentText)
                    background: Rectangle { radius: 8; color: "#FFFFFF"; border.color: parent.activeFocus ? root.primary : root.border }
                }
                ComboBox {
                    id: historyTool
                    Layout.preferredWidth: 155
                    model: controller.historyToolOptions
                    textRole: "label"
                    valueRole: "value"
                }
                ComboBox { id: historyDate; Layout.preferredWidth: 125; model: controller.historyDateOptions }
                AppButton { text: "查找"; onClicked: controller.refreshHistory(historySearch.text, historyTool.currentValue, historyDate.currentText) }
            }
            Text { Layout.fillWidth: true; text: controller.historyMessage; color: root.textMuted; font.pixelSize: 11; wrapMode: Text.Wrap }
            RowLayout {
                Layout.fillWidth: true
                Layout.fillHeight: true
                spacing: 10
                Card {
                    Layout.preferredWidth: Math.max(300, historyDrawer.width * 0.47)
                    Layout.fillHeight: true
                    color: "#F8F7F3"
                    ListView {
                        id: historyList
                        anchors.fill: parent
                        anchors.margins: 8
                        clip: true
                        model: controller.historyModel
                        reuseItems: true
                        cacheBuffer: 168
                        currentIndex: count > 0 ? 0 : -1
                        ScrollBar.vertical: ScrollBar { policy: ScrollBar.AsNeeded }
                        delegate: Rectangle {
                            width: historyList.width
                            height: 76
                            radius: 8
                            color: historyList.currentIndex === index ? "#E1ECE8" : (historyMouse.containsMouse ? "#EFEDE8" : "transparent")
                            ColumnLayout {
                                anchors.fill: parent; anchors.margins: 8; spacing: 2
                                RowLayout {
                                    Layout.fillWidth: true
                                    Text { Layout.fillWidth: true; text: tool; color: root.textMain; font.pixelSize: 13; font.weight: Font.DemiBold; elide: Text.ElideRight }
                                    Text { text: status; color: status === "已完成" ? root.primary : "#A36D10"; font.pixelSize: 10 }
                                }
                                Text { Layout.fillWidth: true; text: time + " · " + inputs; color: root.textMuted; font.pixelSize: 10; elide: Text.ElideMiddle }
                                Text { Layout.fillWidth: true; text: "结果：" + outputs; color: root.textMuted; font.pixelSize: 10; elide: Text.ElideMiddle }
                            }
                            MouseArea {
                                id: historyMouse; anchors.fill: parent; hoverEnabled: true; cursorShape: Qt.PointingHandCursor
                                onClicked: { historyList.currentIndex = index; controller.selectHistoryRow(index) }
                            }
                        }
                    }
                }
                Card {
                    Layout.fillWidth: true
                    Layout.fillHeight: true
                    ColumnLayout {
                        anchors.fill: parent; anchors.margins: 16; spacing: 10
                        Text { Layout.fillWidth: true; text: controller.historyDetail.title || "选择一条记录查看详情"; color: root.textMain; font.pixelSize: 15; font.weight: Font.DemiBold; wrapMode: Text.Wrap }
                        Text { Layout.fillWidth: true; Layout.fillHeight: true; text: controller.historyDetail.body || "这里用于查看升级前由旧版本保存的处理记录。"; color: root.textMuted; font.pixelSize: 12; wrapMode: Text.Wrap; verticalAlignment: Text.AlignTop }
                        Flow {
                            Layout.fillWidth: true; spacing: 6
                            AppButton { text: "打开结果"; enabled: !!controller.historyDetail.canOpenOutput; onClicked: controller.openHistoryOutput() }
                            AppButton { text: "打开上传资料"; enabled: !!controller.historyDetail.canOpenInput; onClicked: controller.openHistoryInput() }
                            AppButton { text: "再次使用"; enabled: !!controller.historyDetail.canReuse; onClicked: controller.reuseHistory() }
                            AppButton { text: "移到回收站"; enabled: !!controller.historyDetail.canDelete; variant: "link"; onClicked: controller.requestMoveHistoryToTrash() }
                        }
                    }
                }
            }
            RowLayout {
                Layout.fillWidth: true
                AppButton { text: "打开归档资料"; variant: "link"; onClicked: controller.openHistoryRoot() }
                AppButton { text: "打开回收站"; variant: "link"; onClicked: controller.openHistoryTrash() }
                AppButton { text: "重新整理记录"; enabled: !controller.historyBusy; variant: "link"; onClicked: controller.rebuildHistoryIndex() }
                Item { Layout.fillWidth: true }
                AppButton { text: "上一页"; enabled: controller.historyHasPrevious && !controller.historyBusy; onClicked: controller.changeHistoryPage(-1) }
                Text { text: controller.historyPageText; color: root.textMuted; font.pixelSize: 11 }
                AppButton { text: "下一页"; enabled: controller.historyHasNext && !controller.historyBusy; onClicked: controller.changeHistoryPage(1) }
            }
        }
    }

    Dialog {
        id: trashDialog
        modal: true
        anchors.centerIn: Overlay.overlay
        width: Math.min(780, root.settledWidth - 42)
        height: Math.min(610, root.settledHeight - 42)
        title: "项目回收站"
        closePolicy: controller.trashBusy ? Popup.NoAutoClose : Popup.CloseOnEscape
        standardButtons: Dialog.Close
        contentItem: ColumnLayout {
            spacing: 9
            Text { Layout.fillWidth: true; text: "这里保存从当前项目移走的完整处理批次；恢复时不会覆盖已有资料。"; color: root.textMuted; font.pixelSize: 11; wrapMode: Text.Wrap }
            TextField {
                Layout.fillWidth: true; placeholderText: "查找已移除的批次"; selectByMouse: true
                onTextEdited: controller.setTrashSearch(text)
                background: Rectangle { radius: 8; color: "#FFFFFF"; border.color: parent.activeFocus ? root.primary : root.border }
            }
            RowLayout {
                Layout.fillWidth: true; Layout.fillHeight: true; spacing: 10
                Card {
                    Layout.preferredWidth: Math.max(290, trashDialog.availableWidth * 0.48)
                    Layout.fillHeight: true; color: "#F8F7F3"
                    ListView {
                        id: trashList
                        anchors.fill: parent; anchors.margins: 8; clip: true
                        model: controller.trashModel; reuseItems: true; cacheBuffer: 150
                        currentIndex: count > 0 ? 0 : -1
                        ScrollBar.vertical: ScrollBar { policy: ScrollBar.AsNeeded }
                        delegate: Rectangle {
                            width: trashList.width; height: 86; radius: 8
                            color: trashList.currentIndex === index ? "#E1ECE8" : (trashMouse.containsMouse ? "#EFEDE8" : "transparent")
                            ColumnLayout {
                                anchors.fill: parent; anchors.margins: 8; spacing: 2
                                Text { Layout.fillWidth: true; text: title; color: root.textMain; font.pixelSize: 12; font.weight: Font.DemiBold; elide: Text.ElideRight }
                                Text { Layout.fillWidth: true; text: tool + " · " + status; color: root.textMuted; font.pixelSize: 10; elide: Text.ElideRight }
                                Text { Layout.fillWidth: true; text: "移入：" + deletedAt; color: root.textMuted; font.pixelSize: 10 }
                                Text { Layout.fillWidth: true; text: counts + " · " + size; color: root.textMuted; font.pixelSize: 10; elide: Text.ElideRight }
                            }
                            MouseArea { id: trashMouse; anchors.fill: parent; hoverEnabled: true; cursorShape: Qt.PointingHandCursor; onClicked: { trashList.currentIndex = index; controller.selectTrashRow(index) } }
                        }
                    }
                }
                Card {
                    Layout.fillWidth: true; Layout.fillHeight: true
                    ColumnLayout {
                        anchors.fill: parent; anchors.margins: 16; spacing: 10
                        Text { text: controller.trashSelectedId ? "恢复到当前项目" : "请选择处理批次"; color: root.textMain; font.pixelSize: 15; font.weight: Font.DemiBold }
                        Text { Layout.fillWidth: true; text: controller.trashSelectedId ? "系统会核对完整清单并恢复到原业务目录；如有同名批次会自动使用新名称。" : "回收站为空，或当前筛选没有匹配结果。"; color: root.textMuted; font.pixelSize: 12; wrapMode: Text.Wrap }
                        Item { Layout.fillHeight: true }
                        AppButton { Layout.fillWidth: true; text: controller.trashBusy ? "正在处理…" : "恢复到项目"; enabled: !!controller.trashSelectedId && controller.projectWritable && !controller.trashBusy && !controller.busy && !controller.workspaceBusy; variant: "primary"; onClicked: controller.restoreSelectedTrash() }
                    }
                }
            }
        }
    }

    Dialog {
        id: helpDialog
        modal: true
        anchors.centerIn: Overlay.overlay
        width: Math.min(650, root.settledWidth - 48)
        height: Math.min(560, root.settledHeight - 48)
        title: "使用教程"
        standardButtons: Dialog.Close
        contentItem: ScrollView {
            clip: true
            contentWidth: availableWidth
            Text {
                width: parent.width
                text: "1. 先新建或打开一个工作项目。项目会完整保留每次处理的上传资料、结果和补充资料。\n\n2. 从左侧选择功能，再选择文件、压缩包或文件夹。工具只在本机处理，不会上传数据。\n\n3. 核对选项后开始处理。运行期间可以继续滚动页面、查看项目文件，也可以请求安全停止。\n\n4. 处理结果自动保存到当前项目。点击“项目文件”可查找、打开、导入或恢复资料。\n\n5. “旧版记录”只用于查看升级前版本保存的历史资料；新的处理记录统一进入工作项目。\n\n6. 关闭程序前，若仍有任务运行，请先停止并等待安全结束，避免外部 Excel 或压缩文件仍被占用。"
                color: root.textMain
                font.pixelSize: 13
                wrapMode: Text.Wrap
                textFormat: Text.PlainText
            }
        }
    }

    Dialog {
        id: textInputDialog
        modal: true
        anchors.centerIn: Overlay.overlay
        width: Math.min(500, root.settledWidth - 48)
        property string promptText: ""
        property string actionToken: ""
        title: "输入"
        standardButtons: Dialog.Ok | Dialog.Cancel
        onAccepted: controller.submitTextAction(actionToken, textInputField.text)
        contentItem: ColumnLayout {
            spacing: 9
            Text { Layout.fillWidth: true; text: textInputDialog.promptText; color: root.textMuted; font.pixelSize: 12; wrapMode: Text.Wrap }
            TextField { id: textInputField; Layout.fillWidth: true; selectByMouse: true }
        }
        function request(titleText, prompt, initialValue, token) {
            title = titleText
            promptText = prompt
            actionToken = token
            textInputField.text = initialValue
            open()
            textInputField.forceActiveFocus()
            textInputField.selectAll()
        }
    }

    Dialog {
        id: updateProgressDialog
        modal: true
        anchors.centerIn: Overlay.overlay
        width: Math.min(520, root.settledWidth - 48)
        title: "正在更新"
        closePolicy: Popup.NoAutoClose
        standardButtons: Dialog.NoButton
        contentItem: ColumnLayout {
            spacing: 12
            Text { Layout.fillWidth: true; text: controller.updateStatus; color: root.textMain; font.pixelSize: 13; wrapMode: Text.Wrap }
            ProgressBar { Layout.fillWidth: true; indeterminate: controller.updateProgress < 0; value: Math.max(0, controller.updateProgress) }
            RowLayout {
                Layout.fillWidth: true
                Item { Layout.fillWidth: true }
                AppButton { text: "取消下载"; visible: controller.updateBusy && controller.updateProgress >= 0; onClicked: controller.cancelUpdate() }
            }
        }
    }

    Dialog {
        id: notificationDialog
        modal: true
        anchors.centerIn: Overlay.overlay
        width: Math.min(520, root.settledWidth - 48)
        property string bodyText: ""
        property string level: "info"
        title: "提示"
        standardButtons: Dialog.Ok
        contentItem: Text { text: notificationDialog.bodyText; color: root.textMain; font.pixelSize: 13; wrapMode: Text.Wrap; textFormat: Text.PlainText; width: notificationDialog.availableWidth }
        function showMessage(titleText, messageText, levelText) {
            title = titleText
            bodyText = messageText
            level = levelText
            open()
        }
    }

    Dialog {
        id: confirmationDialog
        modal: true
        anchors.centerIn: Overlay.overlay
        width: Math.min(540, root.settledWidth - 48)
        property string bodyText: ""
        property string actionToken: ""
        title: "确认"
        standardButtons: Dialog.Yes | Dialog.No
        contentItem: Text { text: confirmationDialog.bodyText; color: root.textMain; font.pixelSize: 13; wrapMode: Text.Wrap; textFormat: Text.PlainText; width: confirmationDialog.availableWidth }
        onAccepted: controller.confirmAction(actionToken, true)
        onRejected: controller.confirmAction(actionToken, false)
    }

    Dialog {
        id: createProjectDialog
        modal: true
        anchors.centerIn: Overlay.overlay
        width: Math.min(590, root.settledWidth - 48)
        title: "新建工作项目"
        standardButtons: Dialog.Ok | Dialog.Cancel
        property alias projectName: projectNameField.text
        property alias projectParent: projectParentField.text
        onAccepted: controller.createProject(projectName, projectParent)
        contentItem: ColumnLayout {
            spacing: 10
            Text { text: "项目是一套可随时打开、完整留存资料的工作文件夹。"; color: root.textMuted; font.pixelSize: 12; wrapMode: Text.Wrap; Layout.fillWidth: true }
            Text { text: "项目名称"; color: root.textMain; font.pixelSize: 13; font.weight: Font.DemiBold }
            TextField { id: projectNameField; Layout.fillWidth: true; selectByMouse: true }
            Text { text: "保存位置"; color: root.textMain; font.pixelSize: 13; font.weight: Font.DemiBold }
            RowLayout {
                Layout.fillWidth: true
                TextField { id: projectParentField; Layout.fillWidth: true; selectByMouse: true }
                AppButton { text: "选择其他位置"; onClicked: { var chosen = controller.chooseProjectParent(); if (chosen) projectParentField.text = chosen } }
            }
            Text { Layout.fillWidth: true; text: projectParentField.text && projectNameField.text ? (projectParentField.text + "/" + projectNameField.text) : ""; color: root.primary; font.pixelSize: 11; elide: Text.ElideMiddle }
        }
    }

    Connections {
        target: controller
        function onNotificationRequested(title, message, level) { notificationDialog.showMessage(title, message, level) }
        function onConfirmationRequested(title, message, token) {
            confirmationDialog.title = title
            confirmationDialog.bodyText = message
            confirmationDialog.actionToken = token
            confirmationDialog.open()
        }
        function onProjectCreationRequested(name, parent) {
            createProjectDialog.projectName = name
            createProjectDialog.projectParent = parent
            createProjectDialog.open()
        }
        function onTextInputRequested(title, prompt, initialValue, token) {
            textInputDialog.request(title, prompt, initialValue, token)
        }
        function onUpdateChanged() {
            var downloading = controller.updateBusy && (
                controller.updateStatus.indexOf("正在准备") === 0 ||
                controller.updateStatus.indexOf("正在下载") === 0
            )
            if (downloading && !updateProgressDialog.opened)
                updateProgressDialog.open()
            else if (!controller.updateBusy && updateProgressDialog.opened)
                updateProgressDialog.close()
        }
    }
}
