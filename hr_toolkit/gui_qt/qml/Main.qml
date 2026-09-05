import QtQuick 2.15
import QtQuick.Controls 2.15
import QtQuick.Layouts 1.15
import QtQuick.Window 2.15
import "components"

ApplicationWindow {
    id: root
    objectName: "mainWindow"
    minimumWidth: 760
    minimumHeight: 600
    readonly property int preferredWindowWidth: 1400
    readonly property int preferredWindowHeight: 780
    readonly property int initialWindowMargin: 16
    readonly property int currentScreenAvailableWidth: Math.min(Screen.width, Screen.desktopAvailableWidth)
    readonly property int currentScreenAvailableHeight: Math.min(Screen.height, Screen.desktopAvailableHeight)
    width: Math.min(preferredWindowWidth,
                    Math.max(minimumWidth, currentScreenAvailableWidth - initialWindowMargin))
    height: Math.min(preferredWindowHeight,
                     Math.max(minimumHeight, currentScreenAvailableHeight - initialWindowMargin))
    visible: true
    color: "#F7F5F1"
    title: "HR Workbench v" + controller.appVersion

    readonly property color primary: "#17715B"
    readonly property color primaryActive: "#125E4B"
    readonly property color primarySoft: "#E4EFEA"
    readonly property color textMain: "#292825"
    readonly property color textMuted: "#78766E"
    readonly property color textFaint: "#98958C"
    readonly property color textDisabled: "#B3B0A6"
    readonly property color border: "#ECEAE4"
    readonly property color borderFaint: "#F1EFE9"
    readonly property color surface: "#FFFFFF"
    readonly property color surfaceAlt: "#FAF9F6"
    readonly property color navSelected: "#EBE8E1"
    readonly property color navHover: "#F0EEE8"
    readonly property int contentMaxWidth: 820
    readonly property bool showLegacyHistoryEntry: false
    // Keep the navigation geometry stable while the native window is dragged.
    // Hysteresis makes the responsive mode switch at most once in either
    // direction instead of repeatedly rebuilding both sides of the layout
    // around one breakpoint.
    property bool compactSidebar: false
    property bool wideContentInsets: false

    // Keep breakpoint-only geometry stable during a live native resize.  The
    // window and its main content still resize through Qt on every frame; only
    // expensive responsive-mode changes and modal-dialog geometry wait until
    // the size has been stable briefly.  The visible project-files panel is a
    // lightweight exception and follows the live window edge below.
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
        if (!wideContentInsets && settledWidth >= 1540)
            wideContentInsets = true
        else if (wideContentInsets && settledWidth <= 1460)
            wideContentInsets = false
    }

    function fieldById(fieldId) {
        // Reading the revision makes every dedicated binding update after the
        // controller changes a dependent field without mirroring business
        // state in QML.
        var revision = controller.formRevision
        var fields = controller.formFields
        for (var index = 0; index < fields.length; ++index) {
            if (String(fields[index].id) === fieldId)
                return fields[index]
        }
        return ({ "id": fieldId, "value": "", "options": [], "visible": false })
    }

    function choiceIndex(field) {
        var options = field.options || []
        for (var index = 0; index < options.length; ++index) {
            if (String(options[index].value) === String(field.value))
                return index
        }
        return 0
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
            color: "#F7F5F1"
            border.color: "#EBE9E4"

            ColumnLayout {
                anchors.fill: parent
                anchors.leftMargin: root.compactSidebar ? 10 : 12
                anchors.rightMargin: root.compactSidebar ? 10 : 12
                anchors.topMargin: 16
                anchors.bottomMargin: 14
                spacing: 0

                Item {
                    Layout.fillWidth: true
                    Layout.preferredHeight: 70
                    RowLayout {
                        anchors.fill: parent
                        anchors.leftMargin: root.compactSidebar ? 0 : 6
                        anchors.topMargin: 10
                        anchors.bottomMargin: 21
                        spacing: 9
                        BrandMark { Layout.preferredWidth: 26; Layout.preferredHeight: 26 }
                        ColumnLayout {
                            visible: !root.compactSidebar
                            Layout.fillWidth: true
                            spacing: 2
                            Text { text: "HR Workbench"; color: root.textMain; font.pixelSize: 15; font.weight: Font.DemiBold }
                            Text { text: "人资运营自动化"; color: root.textFaint; font.pixelSize: 12 }
                        }
                    }
                }

                Card {
                    id: sidebarProjectCard
                    objectName: "sidebarProjectCard"
                    Layout.fillWidth: true
                    Layout.leftMargin: root.compactSidebar ? 0 : 3
                    Layout.rightMargin: root.compactSidebar ? 0 : 3
                    Layout.bottomMargin: root.compactSidebar ? 8 : 12
                    Layout.preferredHeight: root.compactSidebar ? 54 : 154
                    color: root.surface

                    Item {
                        anchors.fill: parent
                        visible: root.compactSidebar
                        Button {
                            anchors.fill: parent
                            hoverEnabled: true
                            focusPolicy: Qt.StrongFocus
                            onClicked: projectMenu.open()
                            contentItem: Text {
                                text: "项目"
                                color: root.textMain
                                font.pixelSize: 11
                                font.weight: Font.DemiBold
                                horizontalAlignment: Text.AlignHCenter
                                verticalAlignment: Text.AlignVCenter
                            }
                            background: Rectangle { color: "transparent"; radius: 8 }
                        }
                    }

                    ColumnLayout {
                        visible: !root.compactSidebar
                        anchors.fill: parent
                        anchors.margins: 10
                        spacing: 0

                        Item {
                            Layout.fillWidth: true
                            Layout.preferredHeight: 25
                            RowLayout {
                                anchors.fill: parent
                                spacing: 4
                                Text {
                                    text: "工作项目"
                                    color: root.textFaint
                                    font.pixelSize: 11
                                    font.weight: Font.DemiBold
                                }
                                Item { Layout.fillWidth: true }
                                Button {
                                    Layout.preferredWidth: 76
                                    Layout.fillHeight: true
                                    opacity: controller.recentProjects.length > 0 ? 1.0 : 0.45
                                    enabled: controller.recentProjects.length > 0
                                    hoverEnabled: true
                                    focusPolicy: Qt.StrongFocus
                                    onClicked: projectMenu.open()
                                    contentItem: Item {
                                        Row {
                                            anchors.centerIn: parent
                                            spacing: 3
                                            Text { text: "最近项目"; color: root.textMuted; font.pixelSize: 11 }
                                            ToolIcon { width: 10; height: 10; iconId: "chevron_down"; strokeColor: root.textMuted; lineWidth: 1.2 }
                                        }
                                    }
                                    background: Rectangle { color: "transparent" }
                                }
                            }
                        }

                        Button {
                            objectName: "projectSelectorButton"
                            Layout.fillWidth: true
                            Layout.preferredHeight: 60
                            hoverEnabled: true
                            focusPolicy: Qt.StrongFocus
                            onClicked: projectMenu.open()
                            leftPadding: 10
                            rightPadding: 9
                            contentItem: RowLayout {
                                spacing: 7
                                ColumnLayout {
                                    Layout.fillWidth: true
                                    spacing: 4
                                    Text {
                                        Layout.fillWidth: true
                                        text: controller.hasProject ? controller.projectName : "尚未打开项目"
                                        color: root.textMain
                                        font.pixelSize: 13
                                        font.weight: Font.DemiBold
                                        elide: Text.ElideRight
                                    }
                                    Text {
                                        Layout.fillWidth: true
                                        text: controller.hasProject ? "当前项目 · 只读" : "新建或打开项目后开始处理"
                                        color: root.textFaint
                                        font.pixelSize: 10
                                        elide: Text.ElideRight
                                    }
                                }
                                ToolIcon { Layout.preferredWidth: 13; Layout.preferredHeight: 13; iconId: "chevron_down"; strokeColor: root.textMuted; lineWidth: 1.25 }
                            }
                            background: Rectangle {
                                radius: 8
                                color: parent.hovered ? "#FAF8F4" : "#FCFBF8"
                                border.color: parent.activeFocus ? root.primary : root.border
                            }
                        }

                        RowLayout {
                            Layout.fillWidth: true
                            Layout.fillHeight: true
                            Layout.topMargin: 5
                            spacing: 0
                            Button {
                                objectName: "newProjectAction"
                                Layout.fillWidth: true
                                Layout.fillHeight: true
                                hoverEnabled: true
                                focusPolicy: Qt.StrongFocus
                                onClicked: controller.requestCreateProject()
                                contentItem: Item {
                                    Row {
                                        anchors.centerIn: parent
                                        spacing: 7
                                        ToolIcon { width: 16; height: 16; iconId: "plus_circle"; strokeColor: "#55534C"; lineWidth: 1.25 }
                                        Text { text: "新建项目"; color: root.textMain; font.pixelSize: 12 }
                                    }
                                }
                                background: Rectangle { radius: 7; color: parent.down ? root.navSelected : (parent.hovered ? root.navHover : "transparent") }
                            }
                            Rectangle { Layout.preferredWidth: 1; Layout.preferredHeight: 18; color: root.border }
                            Button {
                                objectName: "openProjectAction"
                                Layout.fillWidth: true
                                Layout.fillHeight: true
                                hoverEnabled: true
                                focusPolicy: Qt.StrongFocus
                                onClicked: controller.openProjectDialog()
                                contentItem: Item {
                                    Row {
                                        anchors.centerIn: parent
                                        spacing: 7
                                        ToolIcon { width: 16; height: 16; iconId: "folder_open"; strokeColor: "#55534C"; lineWidth: 1.25 }
                                        Text { text: "打开项目"; color: root.textMain; font.pixelSize: 12 }
                                    }
                                }
                                background: Rectangle { radius: 7; color: parent.down ? root.navSelected : (parent.hovered ? root.navHover : "transparent") }
                            }
                        }
                    }
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
                        spacing: 10
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
                                    color: root.textDisabled
                                    font.pixelSize: 11
                                    font.weight: Font.DemiBold
                                }
                                Repeater {
                                    model: modelData.items
                                    delegate: Rectangle {
                                        width: parent.width
                                        height: 33
                                        radius: 8
                                        color: controller.currentTool === modelData.id ? root.navSelected : (navMouse.containsMouse ? root.navHover : "transparent")
                                        Row {
                                            anchors.fill: parent
                                            anchors.leftMargin: root.compactSidebar ? 0 : 9
                                            spacing: 8
                                            Item {
                                                width: root.compactSidebar ? parent.width : 18
                                                height: parent.height
                                                ToolIcon {
                                                    anchors.centerIn: parent
                                                    width: 16
                                                    height: 16
                                                    iconId: modelData.id
                                                    strokeColor: controller.currentTool === modelData.id ? root.primary : "#55534C"
                                                    lineWidth: 1.25
                                                }
                                            }
                                            Text {
                                                visible: !root.compactSidebar
                                                width: parent.width - 34
                                                height: parent.height
                                                text: modelData.label
                                                color: controller.currentTool === modelData.id ? root.primary : "#55534C"
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

                Rectangle { Layout.fillWidth: true; Layout.preferredHeight: 1; Layout.bottomMargin: 6; color: "#EBE9E4" }
                Rectangle {
                    visible: root.showLegacyHistoryEntry
                    Layout.fillWidth: true; Layout.preferredHeight: visible ? 32 : 0; radius: 8; color: historyNavMouse.containsMouse ? root.navHover : "transparent"
                    Row { anchors.fill: parent; anchors.leftMargin: root.compactSidebar ? 0 : 9; spacing: 8
                        Item { width: root.compactSidebar ? parent.width : 18; height: parent.height; ToolIcon { anchors.centerIn: parent; width: 16; height: 16; iconId: "clock"; strokeColor: "#55534C" } }
                        Text { visible: !root.compactSidebar; width: parent.width - 34; height: parent.height; text: "旧版记录"; color: "#55534C"; font.pixelSize: 13; verticalAlignment: Text.AlignVCenter }
                    }
                    MouseArea { id: historyNavMouse; anchors.fill: parent; hoverEnabled: true; cursorShape: Qt.PointingHandCursor; onClicked: { controller.requestHistory(); historyDrawer.open() } }
                }
                Rectangle {
                    Layout.fillWidth: true; Layout.preferredHeight: 32; radius: 8; color: helpNavMouse.containsMouse ? root.navHover : "transparent"
                    Row { anchors.fill: parent; anchors.leftMargin: root.compactSidebar ? 0 : 9; spacing: 8
                        Item { width: root.compactSidebar ? parent.width : 18; height: parent.height; ToolIcon { anchors.centerIn: parent; width: 16; height: 16; iconId: "tutorial"; strokeColor: "#78766E" } }
                        Text { visible: !root.compactSidebar; width: parent.width - 34; height: parent.height; text: "使用教程"; color: "#78766E"; font.pixelSize: 13; verticalAlignment: Text.AlignVCenter }
                    }
                    MouseArea { id: helpNavMouse; anchors.fill: parent; hoverEnabled: true; cursorShape: Qt.PointingHandCursor; onClicked: helpDialog.open() }
                }
                Rectangle { Layout.fillWidth: true; Layout.preferredHeight: 1; Layout.topMargin: 4; Layout.bottomMargin: 8; color: "#EBE9E4" }
                RowLayout {
                    Layout.fillWidth: true
                    Layout.preferredHeight: 24
                    spacing: 5
                    Text { visible: !root.compactSidebar; text: "v" + controller.appVersion; color: root.textDisabled; font.pixelSize: 10 }
                    Rectangle { visible: !root.compactSidebar; Layout.preferredWidth: 4; Layout.preferredHeight: 4; radius: 2; color: "#35A37B" }
                    Item { Layout.fillWidth: true }
                    Text { visible: !root.compactSidebar; text: "本地处理 · 不上传数据"; color: root.textDisabled; font.pixelSize: 10 }
                    Button {
                        id: runLogIconButton
                        objectName: "runLogIconButton"
                        Layout.preferredWidth: 24
                        Layout.preferredHeight: 24
                        hoverEnabled: true
                        focusPolicy: Qt.StrongFocus
                        onClicked: controller.openRunLog()
                        contentItem: Item {
                            ToolIcon { anchors.centerIn: parent; width: 14; height: 14; iconId: "run_log"; strokeColor: runLogIconButton.hovered ? root.textMuted : root.textDisabled; lineWidth: 1.15 }
                        }
                        background: Rectangle { radius: 6; color: parent.hovered ? root.navHover : "transparent" }
                        ToolTip.visible: hovered
                        ToolTip.text: "打开运行日志"
                    }
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
                anchors.leftMargin: root.compactSidebar ? 12 : 28
                anchors.rightMargin: root.compactSidebar ? 58 : (root.wideContentInsets ? 102 : 66)
                anchors.topMargin: 39
                anchors.bottomMargin: 14
                spacing: 14

                RowLayout {
                    Layout.preferredWidth: Math.min(root.contentMaxWidth, mainLayout.width)
                    Layout.maximumWidth: root.contentMaxWidth
                    Layout.alignment: Qt.AlignHCenter
                    spacing: 10
                    ColumnLayout {
                        Layout.fillWidth: true
                        spacing: 14
                        Text { text: controller.toolGroup; color: root.primary; font.pixelSize: 12; font.weight: Font.DemiBold }
                        Text {
                            Layout.fillWidth: true
                            text: controller.toolTitle
                            color: root.textMain
                            font.pixelSize: 24
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
                    AppButton { Layout.preferredWidth: 116; text: controller.updateBusy ? "检查中…" : "↻  检查更新"; enabled: !controller.updateBusy; onClicked: controller.requestUpdateCheck() }
                }

                RowLayout {
                    Layout.preferredWidth: Math.min(root.contentMaxWidth, mainLayout.width)
                    Layout.maximumWidth: root.contentMaxWidth
                    Layout.alignment: Qt.AlignHCenter
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
                    Layout.topMargin: 0
                    clip: true
                    contentWidth: availableWidth
                    contentHeight: contentColumn.implicitHeight
                    ScrollBar.horizontal.policy: ScrollBar.AlwaysOff
                    ScrollBar.vertical.policy: ScrollBar.AsNeeded
                    ScrollBar.vertical.interactive: true

                    ColumnLayout {
                        id: contentColumn
                        width: Math.min(root.contentMaxWidth, mainScroll.availableWidth)
                        height: implicitHeight
                        x: Math.max(0, (mainScroll.availableWidth - width) / 2)
                        spacing: 14

                        Card {
                            Layout.fillWidth: true
                            Layout.preferredHeight: uploadColumn.implicitHeight + 32
                            ColumnLayout {
                                id: uploadColumn
                                anchors.fill: parent
                                anchors.leftMargin: 20
                                anchors.rightMargin: 20
                                anchors.topMargin: 16
                                anchors.bottomMargin: 18
                                spacing: 10
                                RowLayout {
                                    Layout.fillWidth: true
                                    Layout.minimumHeight: 26
                                    Text { text: controller.inputLabel; color: root.textMain; font.pixelSize: 15; font.weight: Font.DemiBold }
                                    Text { Layout.fillWidth: true; text: controller.inputHint; color: root.textFaint; font.pixelSize: 11; wrapMode: Text.Wrap }
                                }
                                Rectangle {
                                    Layout.fillWidth: true
                                    Layout.preferredHeight: inputList.count > 0 ? Math.min(260, Math.max(54, inputList.count * 46)) : 118
                                    radius: 12
                                    color: "#FBFAF7"
                                    border.width: 0

                                    DashedBorder { anchors.fill: parent }

                                    Column {
                                        anchors.centerIn: parent
                                        visible: inputList.count === 0
                                        spacing: 6
                                        FolderDropIcon { anchors.horizontalCenter: parent.horizontalCenter }
                                        Text { anchors.horizontalCenter: parent.horizontalCenter; text: controller.inputDropTitle; color: root.textMain; font.pixelSize: 13; font.weight: Font.DemiBold }
                                        Text {
                                            anchors.horizontalCenter: parent.horizontalCenter
                                            text: controller.inputAllowsFolder && !controller.inputAllowsFiles
                                                ? "点击浏览文件夹路径"
                                                : (controller.inputAllowsFolder ? "浏览文件 · 选择文件夹" : "点击浏览文件")
                                            color: root.primary
                                            font.pixelSize: 11
                                            font.weight: Font.DemiBold
                                        }
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
                            Layout.preferredHeight: formColumn.implicitHeight + 54
                            ColumnLayout {
                                id: formColumn
                                anchors.fill: parent
                                anchors.leftMargin: 24
                                anchors.rightMargin: 24
                                anchors.topMargin: 32
                                anchors.bottomMargin: 22
                                spacing: 11

                                RowLayout {
                                    Layout.fillWidth: true
                                    visible: controller.hasSupportField
                                    spacing: 10
                                    Text { Layout.preferredWidth: 145; text: controller.supportLabel; color: root.textMain; font.pixelSize: 13 }
                                    Item {
                                        Layout.fillWidth: true
                                        Layout.preferredHeight: 30
                                        Text { anchors.fill: parent; text: controller.supportPath || "未选择"; color: controller.supportPath ? root.textMain : root.textFaint; font.pixelSize: 12; verticalAlignment: Text.AlignVCenter; elide: Text.ElideMiddle }
                                    }
                                    AppButton { text: controller.currentTool === "material_collector" ? "选择文件" : controller.supportButtonText; variant: "link"; onClicked: controller.chooseSupportFile() }
                                    AppButton { visible: controller.supportAllowsFolder; text: "选择文件夹"; variant: "link"; onClicked: controller.chooseSupportFolder() }
                                    AppButton { visible: !!controller.supportPath; text: "清除"; variant: "link"; onClicked: controller.clearSupport() }
                                }

                                RowLayout {
                                    Layout.fillWidth: true
                                    visible: controller.currentTool !== "folder_rename"
                                    spacing: 10
                                    Text { Layout.preferredWidth: 145; text: "结果位置"; color: root.textMain; font.pixelSize: 13 }
                                    Rectangle {
                                        Layout.fillWidth: true
                                        Layout.preferredHeight: 36
                                        color: root.surfaceAlt
                                        border.width: 1
                                        border.color: root.border
                                        Text { anchors.fill: parent; anchors.leftMargin: 11; anchors.rightMargin: 11; text: controller.hasProject ? "当前项目 / 本次处理结果" : "请先新建或打开工作项目"; color: controller.hasProject ? root.textFaint : root.textDisabled; font.pixelSize: 12; verticalAlignment: Text.AlignVCenter; elide: Text.ElideMiddle }
                                    }
                                    AppButton { text: "打开项目"; variant: "link"; enabled: controller.hasProject; onClicked: controller.openProjectFolder() }
                                }

                                ColumnLayout {
                                    id: materialOptions
                                    Layout.fillWidth: true
                                    Layout.topMargin: 9
                                    visible: controller.currentTool === "material_collector"
                                    spacing: 5
                                    readonly property var libraryField: root.fieldById("library_mode")
                                    readonly property var targetField: root.fieldById("target_input")
                                    readonly property var collectAllField: root.fieldById("collect_all")
                                    readonly property var zipField: root.fieldById("create_zip")
                                    readonly property var cacheField: root.fieldById("use_ocr_cache")
                                    readonly property bool flatOcr: String(libraryField.value) === "flat_ocr"

                                    Text {
                                        Layout.fillWidth: true
                                        text: "资料检索与打包设置"
                                        color: root.textMain
                                        font.pixelSize: 13
                                        font.weight: Font.DemiBold
                                    }

                                    Rectangle {
                                        Layout.fillWidth: true
                                        Layout.preferredHeight: materialOptionsColumn.implicitHeight + 28
                                        color: root.surface
                                        border.width: 1
                                        border.color: root.border

                                        ColumnLayout {
                                            id: materialOptionsColumn
                                            anchors.left: parent.left
                                            anchors.right: parent.right
                                            anchors.top: parent.top
                                            anchors.margins: 14
                                            spacing: 10

                                            RowLayout {
                                                Layout.fillWidth: true
                                                spacing: 8
                                                Text { Layout.preferredWidth: 66; text: "资料库形式"; color: root.textMain; font.pixelSize: 13 }
                                                AppComboBox {
                                                    id: materialLibraryMode
                                                    Layout.preferredWidth: 290
                                                    model: materialOptions.libraryField.options || []
                                                    textRole: "label"
                                                    currentIndex: root.choiceIndex(materialOptions.libraryField)
                                                    onActivated: controller.setFieldValue("library_mode", materialOptions.libraryField.options[index].value)
                                                }
                                                Text {
                                                    Layout.fillWidth: true
                                                    text: materialOptions.flatOcr ? "源文件不改；首次建立隐藏索引，未变化文件直接复用" : "原模式按姓名文件夹查找"
                                                    color: root.textFaint
                                                    font.pixelSize: 11
                                                    wrapMode: Text.Wrap
                                                }
                                            }

                                            RowLayout {
                                                Layout.fillWidth: true
                                                spacing: 8
                                                Text { Layout.preferredWidth: 66; text: "目标人员"; color: root.textMain; font.pixelSize: 13 }
                                                AppTextField {
                                                    id: materialTargetInput
                                                    Layout.fillWidth: true
                                                    text: materialOptions.targetField.value === undefined || materialOptions.targetField.value === null ? "" : String(materialOptions.targetField.value)
                                                    placeholderText: "姓名或身份证，多人用逗号隔开"
                                                    onTextEdited: controller.setFieldValue("target_input", text)
                                                }
                                                AppButton {
                                                    text: "✕ 清空"
                                                    variant: "link"
                                                    onClicked: {
                                                        materialTargetInput.text = ""
                                                        controller.setFieldValue("target_input", "")
                                                    }
                                                }
                                            }

                                            Text {
                                                Layout.fillWidth: true
                                                Layout.leftMargin: 74
                                                text: "输入姓名或身份证（多人用逗号隔开，如“张三, 李四”）；留空则按名单表格处理"
                                                color: root.textFaint
                                                font.pixelSize: 11
                                                wrapMode: Text.Wrap
                                            }

                                            RowLayout {
                                                Layout.fillWidth: true
                                                spacing: 8
                                                Text { Layout.preferredWidth: 66; text: "打包设置"; color: root.textMain; font.pixelSize: 13; Layout.alignment: Qt.AlignTop; topPadding: 5 }
                                                Flow {
                                                    Layout.fillWidth: true
                                                    spacing: 15
                                                    AppCheckBox {
                                                        text: materialOptions.flatOcr ? "全部（提取 OCR 识别到的该人员全部材料）" : "全部（直接拷贝匹配到的人员整个文件夹）"
                                                        checked: !!materialOptions.collectAllField.value
                                                        onToggled: controller.setFieldValue("collect_all", checked)
                                                    }
                                                    AppCheckBox {
                                                        text: "生成 ZIP 压缩包"
                                                        checked: !!materialOptions.zipField.value
                                                        onToggled: controller.setFieldValue("create_zip", checked)
                                                    }
                                                    AppCheckBox {
                                                        text: "启用缓存"
                                                        checked: !!materialOptions.cacheField.value
                                                        enabled: !materialOptions.flatOcr
                                                        onToggled: controller.setFieldValue("use_ocr_cache", checked)
                                                    }
                                                }
                                            }

                                            Text {
                                                Layout.fillWidth: true
                                                Layout.leftMargin: 74
                                                visible: !!materialOptions.collectAllField.value
                                                text: materialOptions.flatOcr ? "取消勾选「全部」后，可只提取指定材料；索引仍会覆盖整个资料库" : "取消勾选「全部」后可按需勾选材料类型（如身份证、劳动合同等）"
                                                color: root.textFaint
                                                font.pixelSize: 11
                                                wrapMode: Text.Wrap
                                            }

                                            Loader {
                                                Layout.fillWidth: true
                                                visible: !materialOptions.collectAllField.value
                                                active: visible
                                                property var field: root.fieldById("material_types")
                                                sourceComponent: materialCollectorTypesComponent
                                            }
                                        }
                                    }
                                }

                                Repeater {
                                    model: controller.currentTool === "material_collector" ? [] : controller.formFields
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

                            }
                        }

                        RowLayout {
                            Layout.fillWidth: true
                            spacing: 12
                            AppButton {
                                objectName: "runButton"
                                text: controller.runButtonText
                                variant: "primary"
                                // Match the Tk workflow: the primary action remains
                                // clickable before a project is open, then the
                                // controller explains the required next step.  A
                                // running action must also stay clickable so it can
                                // always be stopped safely.
                                enabled: controller.busy || !controller.workspaceBusy
                                implicitWidth: 132
                                implicitHeight: 40
                                onClicked: controller.runOrCancel()
                            }
                            AppButton { text: "打开结果目录"; enabled: controller.canOpenLastResult; implicitWidth: 138; implicitHeight: 40; onClicked: controller.openLastResult() }
                            Text { visible: !!controller.lastRunText; text: controller.lastRunText; color: root.textMuted; font.pixelSize: 12 }
                            Item { Layout.fillWidth: true }
                        }

                        MaterialRunProgress {
                            Layout.fillWidth: true
                            visible: controller.currentTool === "material_collector" && controller.runProgressVisible
                            completed: controller.runProgressCurrent
                            total: controller.runProgressTotal
                            message: controller.runProgressMessage
                            elapsedSeconds: controller.runProgressElapsed
                            waitSeconds: controller.runProgressWaitSeconds
                            active: controller.busy
                        }

                        Card {
                            Layout.fillWidth: true
                            Layout.preferredHeight: 220
                            ColumnLayout {
                                anchors.fill: parent
                                anchors.margins: 17
                                spacing: 8
                                RowLayout {
                                    Layout.fillWidth: true
                                    Text { text: "运行记录"; color: root.textMain; font.pixelSize: 15; font.weight: Font.DemiBold }
                                    Item { Layout.fillWidth: true }
                                    AppButton {
                                        objectName: "copyRunLogsButton"
                                        text: "复制全部"
                                        variant: "link"
                                        onClicked: controller.copyRunLogs()
                                    }
                                }
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
                                        TextEdit {
                                            id: logText
                                            objectName: "logText"
                                            Layout.fillWidth: true
                                            text: model.text
                                            color: level === "muted" ? root.textMuted : root.textMain
                                            font.pixelSize: 12
                                            wrapMode: TextEdit.Wrap
                                            textFormat: TextEdit.PlainText
                                            readOnly: true
                                            selectByMouse: true
                                            selectByKeyboard: true
                                            persistentSelection: true
                                            selectionColor: "#D8EAE3"
                                            selectedTextColor: root.textMain
                                            MouseArea {
                                                anchors.fill: parent
                                                acceptedButtons: Qt.RightButton
                                                onClicked: logContextMenu.popup()
                                            }
                                            Menu {
                                                id: logContextMenu
                                                objectName: "logContextMenu"
                                                MenuItem { text: "复制"; enabled: logText.selectedText.length > 0; onTriggered: logText.copy() }
                                                MenuItem { text: "选择本条"; onTriggered: { logText.forceActiveFocus(); logText.selectAll() } }
                                                MenuItem { text: "复制全部记录"; onTriggered: controller.copyRunLogs() }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        Item { Layout.fillWidth: true; Layout.preferredHeight: 4 }
                    }
                }
            }

            Rectangle {
                id: workspaceRail
                objectName: "workspaceButton"
                anchors.right: parent.right
                anchors.verticalCenter: parent.verticalCenter
                width: 46
                height: 104
                z: 8
                visible: !workspaceDrawer.opened
                color: workspaceRailMouse.containsMouse && controller.hasProject ? "#F0EEE8" : "#FAF9F6"
                border.color: "#ECEAE4"
                opacity: controller.hasProject ? 1.0 : 0.5

                Text {
                    anchors.centerIn: parent
                    text: "项\n目\n文\n件"
                    color: root.primary
                    font.pixelSize: 13
                    font.weight: Font.DemiBold
                    horizontalAlignment: Text.AlignHCenter
                    lineHeight: 1.15
                }
                MouseArea {
                    id: workspaceRailMouse
                    objectName: "workspaceButtonMouse"
                    anchors.fill: parent
                    hoverEnabled: true
                    enabled: controller.hasProject
                    cursorShape: enabled ? Qt.PointingHandCursor : Qt.ArrowCursor
                    onClicked: {
                        controller.setWorkspaceExpanded(true)
                        workspaceDrawer.open()
                    }
                }
            }
        }
    }

    Popup {
        id: projectMenu
        objectName: "projectMenu"
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
            Text { Layout.preferredWidth: 145; text: field.label; color: root.textMain; font.pixelSize: 13 }
            AppTextField {
                Layout.fillWidth: true
                text: field.value === undefined || field.value === null ? "" : String(field.value)
                placeholderText: field.placeholder || ""
                onTextEdited: controller.setFieldValue(field.id, text)
            }
        }
    }

    Component {
        id: choiceFieldComponent
        RowLayout {
            spacing: 10
            Text { Layout.preferredWidth: 145; text: field.label; color: root.textMain; font.pixelSize: 13 }
            AppComboBox {
                id: combo
                Layout.preferredWidth: Math.min(360, Math.max(220, implicitWidth))
                model: field.options || []
                textRole: "label"
                currentIndex: root.choiceIndex(field)
                onActivated: controller.setFieldValue(field.id, field.options[index].value)
            }
            Item { Layout.fillWidth: true }
        }
    }

    Component {
        id: checkFieldComponent
        RowLayout {
            spacing: 10
            Item { Layout.preferredWidth: 145; Layout.preferredHeight: 1 }
            AppCheckBox {
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
            Text { Layout.preferredWidth: 145; text: field.label; color: root.textMain; font.pixelSize: 13; Layout.alignment: Qt.AlignTop }
            ColumnLayout {
                Layout.fillWidth: true
                spacing: 6
                RowLayout {
                    Layout.fillWidth: true
                    spacing: 7
                    AppTextField {
                        Layout.preferredWidth: 138
                        text: field.startValue === undefined || field.startValue === null ? "" : String(field.startValue)
                        placeholderText: field.startPlaceholder || ""
                        onTextEdited: controller.setFieldValue(field.startId, text)
                    }
                    Text { text: "至"; color: root.textMuted; font.pixelSize: 12 }
                    AppTextField {
                        Layout.preferredWidth: 138
                        text: field.endValue === undefined || field.endValue === null ? "" : String(field.endValue)
                        placeholderText: field.endPlaceholder || ""
                        onTextEdited: controller.setFieldValue(field.endId, text)
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
        id: materialCollectorTypesComponent
        ColumnLayout {
            spacing: 5

            Rectangle { Layout.fillWidth: true; Layout.preferredHeight: 1; color: root.borderFaint }
            RowLayout {
                Layout.fillWidth: true
                spacing: 8
                Text { Layout.preferredWidth: 66; text: "指定材料"; color: root.textMain; font.pixelSize: 13 }
                AppButton { text: "全选"; variant: "link"; onClicked: controller.selectAllMaterials() }
                AppButton { text: "取消全选"; variant: "link"; onClicked: controller.clearMaterials() }
                Item { Layout.fillWidth: true }
            }
            RowLayout {
                Layout.fillWidth: true
                Layout.leftMargin: 74
                spacing: 6
                Text { text: "常用组合"; color: root.textFaint; font.pixelSize: 11 }
                AppComboBox {
                    id: materialCollectorPresetCombo
                    Layout.preferredWidth: 160
                    model: controller.materialPresets
                    currentIndex: Math.max(0, controller.materialPresets.indexOf(controller.materialPresetName))
                    onActivated: controller.setMaterialPresetName(currentText)
                }
                AppButton { text: "应用"; variant: "link"; onClicked: controller.applyMaterialPreset(materialCollectorPresetCombo.currentText) }
                Item { Layout.fillWidth: true }
            }
            Flow {
                Layout.fillWidth: true
                Layout.leftMargin: 74
                spacing: 4
                Text { text: "自定义预设"; color: root.textFaint; font.pixelSize: 11; height: 30; verticalAlignment: Text.AlignVCenter }
                AppButton { text: "保存当前为预设"; variant: "link"; onClicked: controller.requestCreateMaterialPreset() }
                AppButton { text: "更新"; variant: "link"; onClicked: controller.updateMaterialPreset(materialCollectorPresetCombo.currentText) }
                AppButton { text: "重命名"; variant: "link"; onClicked: controller.requestRenameMaterialPreset(materialCollectorPresetCombo.currentText) }
                AppButton { text: "删除"; variant: "link"; onClicked: controller.requestDeleteMaterialPreset(materialCollectorPresetCombo.currentText) }
            }
            Flow {
                Layout.fillWidth: true
                Layout.leftMargin: 74
                spacing: 10
                Repeater {
                    model: field.options || []
                    delegate: AppCheckBox {
                        text: modelData.label
                        checked: modelData.selected
                        onToggled: controller.toggleMaterial(modelData.value, checked)
                    }
                }
            }
            RowLayout {
                Layout.fillWidth: true
                Layout.leftMargin: 74
                spacing: 6
                Text { text: "自定义材料"; color: root.textFaint; font.pixelSize: 11 }
                AppComboBox {
                    id: materialCollectorCustomCombo
                    visible: controller.customMaterials.length > 0
                    Layout.preferredWidth: 160
                    model: controller.customMaterials
                }
                AppButton { text: "添加材料"; variant: "link"; onClicked: controller.requestAddCustomMaterial() }
                AppButton {
                    visible: controller.customMaterials.length > 0
                    text: "删除材料"
                    variant: "link"
                    onClicked: controller.requestDeleteCustomMaterial(materialCollectorCustomCombo.currentText)
                }
                Item { Layout.fillWidth: true }
            }
            Text {
                Layout.fillWidth: true
                Layout.leftMargin: 74
                text: "自定义材料和预设会保存在本机；应用组合后仍可继续增减勾选。"
                color: root.textFaint
                font.pixelSize: 11
                wrapMode: Text.Wrap
            }
        }
    }

    Component {
        id: materialsFieldComponent
        ColumnLayout {
            spacing: 7
            RowLayout {
                Layout.fillWidth: true
                Text { Layout.preferredWidth: 145; text: "常用组合"; color: root.textMain; font.pixelSize: 13 }
                AppComboBox {
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
                Layout.leftMargin: 145
                spacing: 4
                AppButton { text: "保存新预设"; variant: "link"; onClicked: controller.requestCreateMaterialPreset() }
                AppButton { text: "更新"; variant: "link"; onClicked: controller.updateMaterialPreset(presetCombo.currentText) }
                AppButton { text: "重命名"; variant: "link"; onClicked: controller.requestRenameMaterialPreset(presetCombo.currentText) }
                AppButton { text: "删除"; variant: "link"; onClicked: controller.requestDeleteMaterialPreset(presetCombo.currentText) }
            }
            RowLayout {
                Layout.fillWidth: true
                Text { Layout.preferredWidth: 145; text: field.label; color: root.textMain; font.pixelSize: 13 }
                AppButton { text: "全选"; variant: "link"; onClicked: controller.selectAllMaterials() }
                AppButton { text: "取消全选"; variant: "link"; onClicked: controller.clearMaterials() }
                Item { Layout.fillWidth: true }
            }
            RowLayout {
                Layout.fillWidth: true
                Item { Layout.preferredWidth: 145; Layout.preferredHeight: 1 }
                AppButton { text: "添加材料"; variant: "link"; onClicked: controller.requestAddCustomMaterial() }
                AppComboBox {
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
                    delegate: AppCheckBox {
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
        // The workspace is already virtualized, so moving this fixed-width
        // layer with the native window is cheap.  It must use live dimensions:
        // settledWidth/settledHeight intentionally stop changing during a
        // border drag and made the open panel appear frozen until mouse-up.
        x: root.width - width
        y: 0
        width: Math.min(340, root.width - 24)
        height: root.height
        modal: false
        dim: false
        padding: 0
        closePolicy: Popup.CloseOnEscape
        onClosed: {
            workspaceAddMenu.close()
            controller.setWorkspaceExpanded(false)
        }
        enter: Transition {}
        exit: Transition {}

        background: Rectangle { color: root.surface; border.color: root.border }
        ColumnLayout {
            anchors.fill: parent
            anchors.leftMargin: 16
            anchors.rightMargin: 16
            anchors.topMargin: 18
            anchors.bottomMargin: 14
            spacing: 8
            RowLayout {
                Layout.fillWidth: true
                Text { Layout.fillWidth: true; text: "项目文件"; color: root.textMain; font.pixelSize: 18; font.weight: Font.DemiBold }
                AppButton { text: "回收站"; variant: "link"; enabled: controller.hasProject; onClicked: { controller.requestProjectTrash(); trashDialog.open() } }
                AppButton { text: "收起"; variant: "link"; onClicked: workspaceDrawer.close() }
            }
            Text { Layout.fillWidth: true; Layout.topMargin: 3; text: controller.projectName; color: root.primary; font.pixelSize: 12; font.weight: Font.DemiBold; elide: Text.ElideRight }
            RowLayout {
                Layout.fillWidth: true
                Layout.bottomMargin: 4
                spacing: 7
                AppButton { text: "切换项目"; onClicked: { workspaceDrawer.close(); projectMenu.open() } }
                AppButton { text: "打开文件夹"; enabled: controller.hasProject; onClicked: controller.openProjectFolder() }
                Item { Layout.fillWidth: true }
            }
            RowLayout {
                Layout.fillWidth: true
                spacing: 6
                AppButton { Layout.fillWidth: true; text: "全部文件"; variant: controller.workspaceScope === "all" ? "tonal" : "secondary"; onClicked: controller.setWorkspaceScope("all") }
                AppButton { Layout.fillWidth: true; text: "当前功能"; variant: controller.workspaceScope === "tool" ? "tonal" : "secondary"; onClicked: controller.setWorkspaceScope("tool") }
            }
            Text { Layout.fillWidth: true; text: "按文件名查找"; color: root.textFaint; font.pixelSize: 11 }
            AppTextField {
                Layout.fillWidth: true
                placeholderText: "输入文件名"
                onTextEdited: controller.setWorkspaceSearch(text)
            }
            RowLayout {
                Layout.fillWidth: true
                Layout.topMargin: 2
                spacing: 7
                AppButton { text: "添加"; variant: "tonal"; enabled: controller.projectWritable && !controller.busy && !controller.workspaceBusy; onClicked: workspaceAddMenu.open() }
                AppButton { text: "刷新"; variant: "link"; onClicked: controller.refreshWorkspace() }
                AppButton { visible: controller.workspaceBusy; text: "取消导入"; variant: "link"; onClicked: controller.cancelWorkspaceImport() }
                Item { Layout.fillWidth: true }
            }
            Rectangle { Layout.fillWidth: true; Layout.preferredHeight: 1; color: root.borderFaint }
            Item {
                Layout.fillWidth: true
                Layout.fillHeight: true
                ListView {
                    id: workspaceList
                    objectName: "workspaceList"
                    anchors.fill: parent
                    clip: true
                    model: controller.workspaceModel
                    reuseItems: true
                    cacheBuffer: 128
                    currentIndex: -1
                    property int activeDelegateCount: 0
                    function keepRowVisible(row) {
                        Qt.callLater(function() {
                            if (workspaceList.count <= 0)
                                return
                            var safeRow = Math.max(0, Math.min(row, workspaceList.count - 1))
                            workspaceList.positionViewAtIndex(safeRow, ListView.Contain)
                        })
                    }
                    ScrollBar.vertical: ScrollBar { policy: ScrollBar.AsNeeded }
                    delegate: Rectangle {
                        Component.onCompleted: workspaceList.activeDelegateCount += 1
                        Component.onDestruction: workspaceList.activeDelegateCount -= 1
                        width: workspaceList.width
                        height: 32
                        radius: 6
                        color: workspaceList.currentIndex === index ? root.primarySoft : (workspaceMouse.containsMouse ? root.navHover : "transparent")
                        RowLayout {
                            anchors.fill: parent
                            anchors.leftMargin: 5 + depth * 15
                            anchors.rightMargin: 5
                            spacing: 4
                            Text { Layout.preferredWidth: 12; text: isDir ? (expanded ? "▾" : "▸") : ""; color: root.textMuted; font.pixelSize: 10 }
                            ToolIcon { Layout.preferredWidth: 16; Layout.preferredHeight: 16; iconId: isDir ? "folder_rename" : "social_security"; strokeColor: isDir ? root.primary : "#617381"; lineWidth: 1.15 }
                            Text { Layout.fillWidth: true; text: name; color: root.textMain; font.pixelSize: 12; elide: Text.ElideMiddle; verticalAlignment: Text.AlignVCenter }
                        }
                        MouseArea {
                            id: workspaceMouse
                            anchors.fill: parent
                            hoverEnabled: true
                            acceptedButtons: Qt.LeftButton
                            onClicked: {
                                workspaceList.currentIndex = index
                                controller.selectWorkspaceRow(index)
                                if (isDir) {
                                    controller.toggleWorkspaceRow(index)
                                    workspaceList.keepRowVisible(index)
                                }
                            }
                            onDoubleClicked: controller.openWorkspaceRow(index)
                        }
                    }
                }
                Text {
                    anchors.centerIn: parent
                    visible: workspaceList.count === 0
                    text: controller.hasProject ? "当前范围还没有项目文件" : "请先打开工作项目"
                    color: root.textFaint
                    font.pixelSize: 12
                    horizontalAlignment: Text.AlignHCenter
                }
            }
            Card {
                Layout.fillWidth: true
                Layout.preferredHeight: 94
                color: root.surface
                ColumnLayout {
                    anchors.fill: parent
                    anchors.margins: 10
                    spacing: 3
                    Text { Layout.fillWidth: true; text: controller.workspaceSelectionAvailable ? controller.workspaceSelectedName : "选择项目文件"; color: root.textMain; font.pixelSize: 13; font.weight: Font.DemiBold; elide: Text.ElideMiddle }
                    Text { Layout.fillWidth: true; text: controller.workspaceSelectedDetail; color: root.textMuted; font.pixelSize: 10; elide: Text.ElideRight }
                    RowLayout {
                        Layout.fillWidth: true
                        AppButton { text: "打开"; variant: "link"; enabled: controller.workspaceSelectionAvailable; onClicked: controller.launchWorkspaceSelection() }
                        AppButton { text: "定位"; variant: "link"; enabled: controller.workspaceSelectionAvailable; onClicked: controller.revealWorkspaceSelection() }
                        AppButton { text: "移到回收站"; variant: "link"; enabled: controller.workspaceSelectionAvailable && controller.projectWritable && !controller.busy && !controller.workspaceBusy; onClicked: controller.requestMoveSelectedBatchToTrash() }
                        Item { Layout.fillWidth: true }
                    }
                }
            }
        }
    }

    Popup {
        id: workspaceAddMenu
        x: Math.max(8, root.width - workspaceDrawer.width + 16)
        y: Math.min(root.height - height - 12, 260)
        width: 210
        padding: 8
        closePolicy: Popup.CloseOnEscape | Popup.CloseOnPressOutside
        enter: Transition {}
        exit: Transition {}
        background: Rectangle { color: root.surface; radius: 9; border.color: root.border }
        contentItem: ColumnLayout {
            spacing: 3
            AppButton { Layout.fillWidth: true; text: "导入文件"; onClicked: { workspaceAddMenu.close(); controller.importWorkspaceFiles() } }
            AppButton { Layout.fillWidth: true; text: "导入文件夹"; onClicked: { workspaceAddMenu.close(); controller.importWorkspaceFolder() } }
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
                AppTextField {
                    id: historySearch
                    Layout.fillWidth: true
                    placeholderText: "按功能或文件名查找"
                    selectByMouse: true
                    onAccepted: controller.refreshHistory(text, historyTool.currentValue, historyDate.currentText)
                }
                AppComboBox {
                    id: historyTool
                    Layout.preferredWidth: 155
                    model: controller.historyToolOptions
                    textRole: "label"
                    valueRole: "value"
                }
                AppComboBox { id: historyDate; Layout.preferredWidth: 125; model: controller.historyDateOptions }
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

    AppDialog {
        id: trashDialog
        modal: true
        anchors.centerIn: Overlay.overlay
        width: Math.min(780, root.settledWidth - 42)
        height: Math.min(610, root.settledHeight - 42)
        title: "项目回收站"
        closePolicy: controller.trashBusy ? Popup.NoAutoClose : Popup.CloseOnEscape
        closeText: "关闭"
        contentItem: ColumnLayout {
            spacing: 9
            Text { Layout.fillWidth: true; text: "这里保存从当前项目移走的完整处理批次；恢复时不会覆盖已有资料。"; color: root.textMuted; font.pixelSize: 11; wrapMode: Text.Wrap }
            AppTextField {
                Layout.fillWidth: true; placeholderText: "查找已移除的批次"; selectByMouse: true
                onTextEdited: controller.setTrashSearch(text)
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

    AppDialog {
        id: helpDialog
        objectName: "helpDialog"
        modal: true
        anchors.centerIn: Overlay.overlay
        width: Math.min(860, root.settledWidth - 48)
        height: Math.min(620, root.settledHeight - 48)
        title: "使用教程"
        closeText: "关闭"
        property var selectedItem: ({ "toolId": "", "mode": "", "label": "", "lines": [] })

        function selectCurrentTutorial() {
            var wantedTool = controller.currentTool
            var wantedMode = (wantedTool === "personnel_change_merge" || wantedTool === "archive_import") ? controller.currentVariant : ""
            var firstItem = null
            for (var groupIndex = 0; groupIndex < controller.tutorialGroups.length; ++groupIndex) {
                var items = controller.tutorialGroups[groupIndex].items || []
                for (var itemIndex = 0; itemIndex < items.length; ++itemIndex) {
                    var item = items[itemIndex]
                    if (firstItem === null)
                        firstItem = item
                    if (String(item.toolId) === wantedTool && String(item.mode || "") === wantedMode) {
                        selectedItem = item
                        return
                    }
                }
            }
            if (firstItem !== null)
                selectedItem = firstItem
        }

        function isSelected(item) {
            return String(selectedItem.toolId || "") === String(item.toolId || "")
                    && String(selectedItem.mode || "") === String(item.mode || "")
        }

        onOpened: selectCurrentTutorial()

        contentItem: RowLayout {
            spacing: 16

            Rectangle {
                Layout.preferredWidth: 190
                Layout.fillHeight: true
                radius: 9
                color: root.surfaceAlt
                border.color: root.borderFaint

                ScrollView {
                    id: tutorialNavigation
                    anchors.fill: parent
                    anchors.margins: 8
                    clip: true
                    contentWidth: availableWidth
                    ScrollBar.horizontal.policy: ScrollBar.AlwaysOff

                    Column {
                        width: tutorialNavigation.availableWidth
                        spacing: 8
                        Repeater {
                            model: controller.tutorialGroups
                            delegate: Column {
                                property var groupData: modelData
                                width: parent.width
                                spacing: 2
                                Text {
                                    width: parent.width
                                    leftPadding: 8
                                    topPadding: 5
                                    bottomPadding: 3
                                    text: groupData.name
                                    color: root.textDisabled
                                    font.pixelSize: 11
                                    font.weight: Font.DemiBold
                                }
                                Repeater {
                                    model: groupData.items
                                    delegate: Rectangle {
                                        width: parent.width
                                        height: 31
                                        radius: 7
                                        color: helpDialog.isSelected(modelData) ? root.navSelected : (tutorialItemMouse.containsMouse ? root.navHover : "transparent")
                                        Row {
                                            anchors.fill: parent
                                            anchors.leftMargin: 8
                                            spacing: 7
                                            Item {
                                                width: 17
                                                height: parent.height
                                                ToolIcon {
                                                    anchors.centerIn: parent
                                                    width: 15
                                                    height: 15
                                                    iconId: modelData.toolId
                                                    strokeColor: helpDialog.isSelected(modelData) ? root.primary : root.textMuted
                                                    lineWidth: 1.15
                                                }
                                            }
                                            Text {
                                                width: parent.width - 32
                                                height: parent.height
                                                text: modelData.label
                                                color: helpDialog.isSelected(modelData) ? root.primary : root.textMain
                                                font.pixelSize: 12
                                                font.weight: helpDialog.isSelected(modelData) ? Font.DemiBold : Font.Normal
                                                verticalAlignment: Text.AlignVCenter
                                                elide: Text.ElideRight
                                            }
                                        }
                                        MouseArea {
                                            id: tutorialItemMouse
                                            anchors.fill: parent
                                            hoverEnabled: true
                                            cursorShape: Qt.PointingHandCursor
                                            onClicked: helpDialog.selectedItem = modelData
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            Card {
                Layout.fillWidth: true
                Layout.fillHeight: true
                color: root.surface

                ScrollView {
                    id: tutorialContentScroll
                    anchors.fill: parent
                    anchors.margins: 18
                    clip: true
                    contentWidth: availableWidth
                    ScrollBar.horizontal.policy: ScrollBar.AlwaysOff

                    Column {
                        width: tutorialContentScroll.availableWidth
                        spacing: 11
                        Text {
                            width: parent.width
                            text: helpDialog.selectedItem.label || ""
                            color: root.textMain
                            font.pixelSize: 17
                            font.weight: Font.DemiBold
                            bottomPadding: 3
                        }
                        Repeater {
                            model: helpDialog.selectedItem.lines || []
                            delegate: Text {
                                width: parent.width
                                text: modelData.text
                                color: modelData.style === "warning" ? "#B06B13" : root.textMain
                                font.pixelSize: 13
                                font.weight: modelData.style === "strong" || modelData.style === "warning" ? Font.DemiBold : Font.Normal
                                wrapMode: Text.Wrap
                                textFormat: Text.PlainText
                                lineHeight: 1.22
                            }
                        }
                    }
                }
            }
        }
    }

    AppDialog {
        id: textInputDialog
        modal: true
        anchors.centerIn: Overlay.overlay
        width: Math.min(500, root.settledWidth - 48)
        property string promptText: ""
        property string actionToken: ""
        title: "输入"
        acceptText: "确定"
        rejectText: "取消"
        onAccepted: controller.submitTextAction(actionToken, textInputField.text)
        contentItem: ColumnLayout {
            spacing: 9
            Text { Layout.fillWidth: true; text: textInputDialog.promptText; color: root.textMuted; font.pixelSize: 12; wrapMode: Text.Wrap }
            AppTextField { id: textInputField; Layout.fillWidth: true; selectByMouse: true }
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

    AppDialog {
        id: updateProgressDialog
        modal: true
        anchors.centerIn: Overlay.overlay
        width: Math.min(520, root.settledWidth - 48)
        title: "正在更新"
        closePolicy: Popup.NoAutoClose
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

    AppDialog {
        id: notificationDialog
        modal: true
        anchors.centerIn: Overlay.overlay
        width: Math.min(520, root.settledWidth - 48)
        property string bodyText: ""
        property string level: "info"
        title: "提示"
        closeText: "知道了"
        contentItem: Text { text: notificationDialog.bodyText; color: root.textMain; font.pixelSize: 13; wrapMode: Text.Wrap; textFormat: Text.PlainText; width: notificationDialog.availableWidth }
        function showMessage(titleText, messageText, levelText) {
            title = titleText
            bodyText = messageText
            level = levelText
            open()
        }
    }

    AppDialog {
        id: confirmationDialog
        modal: true
        anchors.centerIn: Overlay.overlay
        width: Math.min(540, root.settledWidth - 48)
        property string bodyText: ""
        property string actionToken: ""
        title: "确认"
        acceptText: "确定"
        rejectText: "取消"
        contentItem: Text { text: confirmationDialog.bodyText; color: root.textMain; font.pixelSize: 13; wrapMode: Text.Wrap; textFormat: Text.PlainText; width: confirmationDialog.availableWidth }
        onAccepted: controller.confirmAction(actionToken, true)
        onRejected: controller.confirmAction(actionToken, false)
    }

    AppDialog {
        id: createProjectDialog
        modal: true
        anchors.centerIn: Overlay.overlay
        width: Math.min(590, root.settledWidth - 48)
        title: "新建工作项目"
        acceptText: "创建并打开"
        rejectText: "取消"
        property alias projectName: projectNameField.text
        property alias projectParent: projectParentField.text
        onAccepted: controller.createProject(projectName, projectParent)
        contentItem: ColumnLayout {
            spacing: 10
            Text { text: "项目是一套可随时打开、完整留存资料的工作文件夹。"; color: root.textMuted; font.pixelSize: 12; wrapMode: Text.Wrap; Layout.fillWidth: true }
            Text { text: "项目名称"; color: root.textMain; font.pixelSize: 13; font.weight: Font.DemiBold }
            AppTextField { id: projectNameField; Layout.fillWidth: true; selectByMouse: true }
            Text { text: "保存位置"; color: root.textMain; font.pixelSize: 13; font.weight: Font.DemiBold }
            RowLayout {
                Layout.fillWidth: true
                AppTextField { id: projectParentField; Layout.fillWidth: true; selectByMouse: true }
                AppButton { text: "选择其他位置"; onClicked: { var chosen = controller.chooseProjectParent(projectParentField.text); if (chosen) projectParentField.text = chosen } }
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
