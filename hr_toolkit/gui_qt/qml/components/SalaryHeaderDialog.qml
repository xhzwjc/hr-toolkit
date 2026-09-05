import QtQuick 2.15
import QtQuick.Controls 2.15
import QtQuick.Layouts 1.15

AppDialog {
    id: dialog
    property var backend
    property bool working: backend ? backend.busy : false
    property var groups: []
    property var issues: []
    property int selectedGroup: 0
    property int revision: 0
    property bool closingFromBackend: false
    property var currentGroup: groups.length ? groups[Math.min(selectedGroup, groups.length - 1)] : ({})
    objectName: "salaryHeaderDialog"
    title: "确认工资列头对应关系"
    width: Math.min(parent ? parent.width - 32 : 820, 820)
    height: Math.min(parent ? parent.height - 32 : 760, 760)
    x: parent ? (parent.width - width) / 2 : 0
    y: parent ? (parent.height - height) / 2 : 0
    closePolicy: Popup.CloseOnEscape

    function showData(data) {
        closingFromBackend = false
        groups = JSON.parse(JSON.stringify(data.groups || []))
        issues = JSON.parse(JSON.stringify(data.issues || []))
        selectedGroup = 0
        for (var i = 0; i < groups.length; i++) {
            if (!groups[i].ready && !groups[i].skip) {
                selectedGroup = i
                break
            }
        }
        revision++
        if (!opened) rememberCheck.checked = true
        open()
    }
    function dismiss() {
        if (opened) {
            closingFromBackend = true
            close()
        }
    }
    function choicePayload() {
        var chosen = []
        for (var i = 0; i < groups.length; i++)
            chosen.push({group_id: groups[i].group_id, selections: groups[i].selections, skip: !!groups[i].skip})
        var skipped = []
        for (var j = 0; j < issues.length; j++)
            if (issues[j].skip) skipped.push(issues[j].key)
        return JSON.stringify({groups: chosen, skipped_issues: skipped})
    }
    function selectedColumn(field) {
        var tick = revision
        return currentGroup.selections ? Number(currentGroup.selections[field] || 0) : 0
    }
    function chooseColumn(field, column) {
        currentGroup.selections[field] = column
        revision++
    }
    function columnOptions() {
        return [{column: 0, display: "请选择原表中的对应列", samples: ""}].concat(currentGroup.columns || [])
    }
    function columnIndex(field) {
        var value = selectedColumn(field)
        var options = columnOptions()
        for (var i = 0; i < options.length; i++)
            if (options[i].column === value) return i
        return 0
    }
    function sample(field) {
        var value = selectedColumn(field)
        var columns = currentGroup.columns || []
        for (var i = 0; i < columns.length; i++)
            if (columns[i].column === value) return columns[i].samples || "此列前几行没有可显示的样例"
        return "选择对应列后显示样例"
    }
    function problemText() {
        var tick = revision
        var problem = currentGroup.problem || ""
        if (problem.indexOf("请选择对应列：") !== 0) return problem
        var missing = []
        if (!selectedColumn("name")) missing.push("姓名")
        if (!selectedColumn("id_card")) missing.push("身份证号码")
        if (currentGroup.role === "detail" && !selectedColumn("amount")) missing.push("应发工资")
        return missing.length ? "请选择对应列：" + missing.join("、") : "对应列已选择，请核对样例后继续。"
    }
    function unresolvedCount() {
        var tick = revision
        var count = 0
        for (var i = 0; i < groups.length; i++) {
            var g = groups[i]
            if (g.skip) continue
            if (!g.selections.name || !g.selections.id_card || (g.role === "detail" && !g.selections.amount)) count++
        }
        for (var j = 0; j < issues.length; j++)
            if (!issues[j].skip) count++
        return count
    }
    onClosed: {
        if (!closingFromBackend && backend) backend.cancelSalaryMappings()
        closingFromBackend = false
    }

    contentItem: Flickable {
        clip: true
        contentWidth: width
        contentHeight: content.implicitHeight
        boundsBehavior: Flickable.StopAtBounds
        ScrollBar.vertical: ScrollBar {}
        ColumnLayout {
            id: content
            width: parent.width
            spacing: 13
            Text {
                Layout.fillWidth: true
                text: "可混合上传不同模板。已识别的文件无需重选；在下拉框分别确认其余模板，勾选“记住”后下次自动使用，各套设置互不替换。"
                color: "#55534D"; font.pixelSize: 13; wrapMode: Text.Wrap
            }
            AppComboBox {
                id: templatePicker
                objectName: "salaryTemplatePicker"
                Layout.fillWidth: true
                visible: dialog.groups.length > 0
                enabled: !dialog.working
                model: dialog.groups.map(function(g, i) {
                    return "模板 " + (i + 1) + " · " + (g.role === "summary" ? "已有汇总表" : "工资明细") + " · " + g.files.length + " 个文件 · " + (g.saved ? "已保存设置" : g.ready ? "已自动识别" : "需要确认")
                })
                currentIndex: dialog.selectedGroup
                onActivated: function(index) { dialog.selectedGroup = index }
            }
            Text {
                Layout.fillWidth: true
                visible: dialog.groups.length > 0
                text: (dialog.currentGroup.files || []).map(function(f) { return f.name }).slice(0, 8).join("、") + ((dialog.currentGroup.files || []).length > 8 ? " 等" : "")
                color: "#77746D"; font.pixelSize: 11; wrapMode: Text.Wrap; textFormat: Text.PlainText
            }
            Text {
                Layout.fillWidth: true
                visible: !!dialog.currentGroup.problem
                text: dialog.problemText()
                color: "#A26713"; font.pixelSize: 12; wrapMode: Text.Wrap; textFormat: Text.PlainText
            }
            ColumnLayout {
                Layout.fillWidth: true
                visible: dialog.groups.length > 0
                enabled: dialog.revision >= 0 && !dialog.working && !dialog.currentGroup.skip
                spacing: 12
                Repeater {
                    model: dialog.currentGroup.role === "summary"
                        ? [{key: "name", label: "姓名"}, {key: "id_card", label: "身份证号码"}]
                        : [{key: "name", label: "姓名"}, {key: "id_card", label: "身份证号码"}, {key: "amount", label: "应发工资"}]
                    delegate: ColumnLayout {
                        property string fieldKey: modelData.key
                        Layout.fillWidth: true
                        spacing: 4
                        RowLayout {
                            Layout.fillWidth: true
                            Text { Layout.preferredWidth: 100; text: modelData.label; color: "#292825"; font.pixelSize: 13 }
                            AppComboBox {
                                objectName: "salaryColumn_" + fieldKey
                                Layout.fillWidth: true
                                textRole: "display"
                                model: dialog.columnOptions()
                                currentIndex: dialog.columnIndex(fieldKey)
                                onActivated: function(index) { dialog.chooseColumn(fieldKey, model[index].column) }
                            }
                        }
                        Text {
                            Layout.fillWidth: true
                            Layout.leftMargin: 110
                            text: "样例：" + dialog.sample(fieldKey)
                            color: "#878279"; font.pixelSize: 11; wrapMode: Text.Wrap; textFormat: Text.PlainText
                            maximumLineCount: 2; elide: Text.ElideRight
                        }
                    }
                }
                Text {
                    Layout.fillWidth: true
                    text: "请核对本次要合并的应发金额，避免选择实发工资或累计金额。样例未缓存时显示原公式。"
                    visible: dialog.currentGroup.role !== "summary"
                    color: "#77746D"; font.pixelSize: 11; wrapMode: Text.Wrap
                }
                RowLayout {
                    Layout.fillWidth: true
                    AppButton { text: advanced.visible ? "收起工作表与表头位置" : "选择工作表与表头位置"; variant: "link"; onClicked: advanced.visible = !advanced.visible }
                    Item { Layout.fillWidth: true }
                    AppButton { objectName: "resetSalaryHeaders"; text: "恢复自动识别"; variant: "link"; onClicked: dialog.backend.resetSalaryHeader(dialog.currentGroup.group_id, dialog.choicePayload()) }
                }
                ColumnLayout {
                    id: advanced
                    Layout.fillWidth: true
                    visible: false
                    Text {
                        Layout.fillWidth: true
                        text: "可选择前 200 行内的表头，连续最多 6 行；列列表显示前 512 列中的非空列头。"
                        color: "#77746D"; font.pixelSize: 11; wrapMode: Text.Wrap
                    }
                    AppComboBox {
                        id: sheetPicker
                        objectName: "salarySheetPicker"
                        Layout.fillWidth: true
                        model: dialog.currentGroup.sheet_names || []
                        currentIndex: model.indexOf(dialog.currentGroup.sheet)
                    }
                    RowLayout {
                        Layout.fillWidth: true
                        Text { text: "表头第"; font.pixelSize: 12 }
                        SpinBox { id: firstRow; objectName: "salaryHeaderFirst"; from: 1; to: 200; value: dialog.currentGroup.header_row || 1; editable: true; Layout.preferredWidth: 105 }
                        Text { text: "行至"; font.pixelSize: 12 }
                        SpinBox { id: lastRow; objectName: "salaryHeaderLast"; from: 1; to: 200; value: dialog.currentGroup.header_bottom || 1; editable: true; Layout.preferredWidth: 105 }
                        Text { text: "行"; font.pixelSize: 12 }
                        Item { Layout.fillWidth: true }
                        AppButton {
                            text: "重新读取列头"
                            onClicked: dialog.backend.rescanSalaryHeader(dialog.currentGroup.group_id, sheetPicker.currentText, firstRow.value, lastRow.value, dialog.choicePayload())
                        }
                    }
                }
            }
            AppCheckBox {
                visible: dialog.groups.length > 0 && dialog.currentGroup.role === "detail"
                enabled: !dialog.working
                text: "本次不合并这一类文件（将在结果中列明）"
                checked: dialog.revision >= 0 && !!dialog.currentGroup.skip
                onToggled: { dialog.currentGroup.skip = checked; dialog.revision++ }
            }
            Repeater {
                model: dialog.issues
                delegate: ColumnLayout {
                    Layout.fillWidth: true
                    Text { Layout.fillWidth: true; text: modelData.name + "：" + modelData.message; textFormat: Text.PlainText; color: "#A26713"; font.pixelSize: 12; wrapMode: Text.Wrap }
                    AppCheckBox {
                        enabled: modelData.skippable && !dialog.working
                        text: modelData.skippable ? "本次不合并此文件（将在结果中列明）" : "此问题需要处理后重新选择资料"
                        checked: !!modelData.skip
                        onToggled: { dialog.issues[index].skip = checked; dialog.revision++ }
                    }
                }
            }
        }
    }
    footer: Rectangle {
        implicitHeight: 94
        color: "#FAF9F6"
        radius: dialog.background.radius
        Rectangle {
            anchors.left: parent.left
            anchors.right: parent.right
            anchors.top: parent.top
            height: parent.radius
            color: parent.color
        }
        ColumnLayout {
            anchors.fill: parent
            anchors.margins: 12
            spacing: 5
            RowLayout {
                Layout.fillWidth: true
                AppCheckBox { id: rememberCheck; objectName: "rememberSalaryHeaders"; text: "记住当前项目的对应关系"; checked: true; enabled: !dialog.working }
                Item { Layout.fillWidth: true }
                Text { text: dialog.working ? "正在读取列头…" : dialog.unresolvedCount() ? "还有 " + dialog.unresolvedCount() + " 项需要处理" : "确认后继续合并"; font.pixelSize: 11; color: "#77746D" }
            }
            RowLayout {
                Layout.fillWidth: true
                Item { Layout.fillWidth: true }
                AppButton { text: "取消"; onClicked: dialog.close() }
                AppButton {
                    objectName: "saveSalaryHeadersOnly"
                    text: "保存设置并关闭"
                    enabled: rememberCheck.checked && !dialog.working
                    onClicked: dialog.backend.applySalaryMappings(dialog.choicePayload(), true, false)
                }
                AppButton {
                    objectName: "applySalaryHeaders"
                    text: rememberCheck.checked ? "保存设置并继续合并" : "仅本次使用并继续合并"
                    variant: "primary"
                    enabled: !dialog.working && dialog.groups.length > 0 && dialog.unresolvedCount() === 0
                    onClicked: dialog.backend.applySalaryMappings(dialog.choicePayload(), rememberCheck.checked)
                }
            }
        }
    }
}
