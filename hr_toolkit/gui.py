from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, VERTICAL, Y, filedialog, messagebox
from tkinter import Tk, StringVar, Text
from tkinter import ttk

from hr_toolkit import __version__
from hr_toolkit.tools.folder_rename import (
    MODE_APPEND,
    MODE_REMOVE,
    MODE_REPLACE,
    rename_person_folders,
)
from hr_toolkit.tools.personnel_change_merge import merge_personnel_changes
from hr_toolkit.tools.salary_merge import merge_monthly_salary
from hr_toolkit.tools.salary_split import split_salary_by_company


RENAME_MODE_LABELS = {
    "追加文字": MODE_APPEND,
    "删除结尾文字": MODE_REMOVE,
    "修改单人名称": MODE_REPLACE,
}


class HRToolkitApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title(f"HR工具箱 v{__version__}")
        self.root.geometry("860x700")
        self.root.minsize(800, 620)

        self.current_tool = "salary_split"
        self.tool_title = StringVar()
        self.tool_description = StringVar()
        self.input_label = StringVar()
        self.choose_input_text = StringVar()
        self.run_button_text = StringVar()
        self.summary_label = StringVar()
        self.summary_button_text = StringVar()
        self.rename_mode = StringVar(value="追加文字")
        self.rename_target_label = StringVar(value="姓名（可不填）")
        self.rename_text_label = StringVar(value="要追加的文字")
        self.rename_replacement_label = StringVar(value="新名称")
        self.rename_target_name = StringVar()
        self.rename_text = StringVar()
        self.rename_replacement_name = StringVar()
        self.input_path = StringVar()
        self.summary_path = StringVar()
        self.output_dir = StringVar(value=str(default_output_parent_dir(self.current_tool)))
        self.output_dir_user_selected = False
        self.status_queue: queue.Queue[tuple[str, object | None]] = queue.Queue()
        self.last_output_dir: Path | None = None

        self._set_tool_texts()
        self._build_layout()
        self._poll_status_queue()

    def _build_layout(self) -> None:
        root_frame = ttk.Frame(self.root, padding=12)
        root_frame.pack(fill=BOTH, expand=True)

        left_frame = ttk.Frame(root_frame, width=220)
        left_frame.pack(side=LEFT, fill=Y, padx=(0, 12))
        left_frame.pack_propagate(False)

        ttk.Label(left_frame, text="工具列表", font=("", 11, "bold")).pack(anchor="w")
        self.tool_list = ttk.Treeview(left_frame, show="tree", height=14)
        self.tool_list.pack(fill=BOTH, expand=True, pady=(8, 0))
        self.tool_list.insert("", END, iid="salary_split", text="需求4 工资表拆分")
        self.tool_list.insert("", END, iid="salary_merge", text="需求5 工资表合并")
        self.tool_list.insert("", END, iid="personnel_change_merge", text="需求6 异动表汇总")
        self.tool_list.insert("", END, iid="folder_rename", text="需求8 文件夹改名")
        self.tool_list.insert("", END, iid="archive_import", text="需求7 档案入库（待实现）")
        self.tool_list.selection_set("salary_split")
        self.tool_list.bind("<<TreeviewSelect>>", self._on_tool_selected)

        right_frame = ttk.Frame(root_frame)
        right_frame.pack(side=RIGHT, fill=BOTH, expand=True)

        ttk.Label(
            right_frame,
            textvariable=self.tool_title,
            font=("", 13, "bold"),
        ).pack(anchor="w")
        ttk.Label(
            right_frame,
            textvariable=self.tool_description,
        ).pack(anchor="w", pady=(4, 14))

        self.tutorial_frame = ttk.LabelFrame(right_frame, text="使用教程", padding=8)
        self.tutorial_frame.pack(fill="x", pady=(0, 12))
        self.tutorial_text = Text(
            self.tutorial_frame,
            height=8,
            wrap="word",
            padx=8,
            pady=6,
            bg="#fff7d6",
            relief="solid",
            bd=1,
        )
        self.tutorial_text.pack(fill="x")
        self.tutorial_text.tag_configure("strong", font=("", 10, "bold"))
        self.tutorial_text.tag_configure("warning", foreground="#9f1d1d", font=("", 10, "bold"))
        self._set_tutorial_text()

        form = ttk.Frame(right_frame)
        form.pack(fill="x")

        ttk.Label(form, textvariable=self.input_label).grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(form, textvariable=self.input_path).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(form, textvariable=self.choose_input_text, command=self._choose_input).grid(row=0, column=2, sticky="e")

        self.summary_label_widget = ttk.Label(form, textvariable=self.summary_label)
        self.summary_entry_widget = ttk.Entry(form, textvariable=self.summary_path)
        self.summary_button_widget = ttk.Button(form, textvariable=self.summary_button_text, command=self._choose_summary)

        self.output_label_widget = ttk.Label(form, text="保存位置")
        self.output_entry_widget = ttk.Entry(form, textvariable=self.output_dir)
        self.output_button_widget = ttk.Button(form, text="选择目录", command=self._choose_output)
        self.output_label_widget.grid(row=2, column=0, sticky="w", pady=4)
        self.output_entry_widget.grid(row=2, column=1, sticky="ew", padx=8)
        self.output_button_widget.grid(row=2, column=2, sticky="e")

        self.rename_options_frame = ttk.LabelFrame(form, text="文件夹改名", padding=8)
        self.rename_options_frame.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        ttk.Label(self.rename_options_frame, text="操作").grid(row=0, column=0, sticky="w", pady=3)
        self.rename_mode_widget = ttk.Combobox(
            self.rename_options_frame,
            textvariable=self.rename_mode,
            values=list(RENAME_MODE_LABELS.keys()),
            state="readonly",
            width=16,
        )
        self.rename_mode_widget.grid(row=0, column=1, sticky="w", padx=8, pady=3)
        self.rename_mode_widget.bind("<<ComboboxSelected>>", self._on_rename_mode_changed)

        ttk.Label(self.rename_options_frame, textvariable=self.rename_target_label).grid(row=1, column=0, sticky="w", pady=3)
        self.rename_target_widget = ttk.Entry(self.rename_options_frame, textvariable=self.rename_target_name)
        self.rename_target_widget.grid(row=1, column=1, sticky="ew", padx=8, pady=3)

        self.rename_text_label_widget = ttk.Label(self.rename_options_frame, textvariable=self.rename_text_label)
        self.rename_text_label_widget.grid(row=2, column=0, sticky="w", pady=3)
        self.rename_text_widget = ttk.Entry(self.rename_options_frame, textvariable=self.rename_text)
        self.rename_text_widget.grid(row=2, column=1, sticky="ew", padx=8, pady=3)

        self.rename_replacement_label_widget = ttk.Label(self.rename_options_frame, textvariable=self.rename_replacement_label)
        self.rename_replacement_label_widget.grid(row=3, column=0, sticky="w", pady=3)
        self.rename_replacement_widget = ttk.Entry(self.rename_options_frame, textvariable=self.rename_replacement_name)
        self.rename_replacement_widget.grid(row=3, column=1, sticky="ew", padx=8, pady=3)
        self.rename_options_frame.columnconfigure(1, weight=1)
        form.columnconfigure(1, weight=1)
        self._update_summary_controls()
        self._update_output_controls()
        self._update_rename_controls()

        actions = ttk.Frame(right_frame)
        actions.pack(fill="x", pady=(14, 10))
        self.run_button = ttk.Button(actions, textvariable=self.run_button_text, command=self._run_current_tool)
        self.run_button.pack(side=LEFT)
        self.open_button = ttk.Button(actions, text="打开所在文件夹", command=self._open_output_dir)
        self.open_button.pack(side=LEFT, padx=(8, 0))

        ttk.Label(right_frame, text="执行结果").pack(anchor="w")
        log_frame = ttk.Frame(right_frame)
        log_frame.pack(fill=BOTH, expand=True, pady=(6, 0))
        scrollbar = ttk.Scrollbar(log_frame, orient=VERTICAL)
        self.log_text = Text(log_frame, height=12, wrap="word", yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.log_text.yview)
        self.log_text.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

        self._write_log(self._initial_log_text())

    def _on_tool_selected(self, _event=None) -> None:
        selection = self.tool_list.selection()
        if not selection:
            return
        self.current_tool = selection[0]
        self.last_output_dir = None
        self.input_path.set("")
        self.summary_path.set("")
        self.rename_target_name.set("")
        self.rename_text.set("")
        self.rename_replacement_name.set("")
        if not self.output_dir_user_selected:
            self.output_dir.set(str(default_output_parent_dir(self.current_tool)))
        self._set_tool_texts()
        self._clear_log()
        self._write_log(self._initial_log_text())

    def _set_tool_texts(self) -> None:
        if self.current_tool == "salary_merge":
            self.tool_title.set("需求5：多月工资合并个人薪资汇总")
            self.tool_description.set("选择工资表文件夹；如已有汇总表，可一并选择后追加新月份。")
            self.input_label.set("工资表文件夹")
            self.choose_input_text.set("选择文件夹")
            self.summary_label.set("已有汇总表（可选）")
            self.summary_button_text.set("选择汇总表")
            self.run_button_text.set("开始合并")
        elif self.current_tool == "personnel_change_merge":
            self.tool_title.set("需求6：异动表汇总")
            self.tool_description.set("选择包含各项目异动表的文件夹，工具会汇总增员、减员、转正、调动、奖罚扣补。")
            self.input_label.set("异动表文件夹")
            self.choose_input_text.set("选择文件夹")
            self.summary_label.set("异动表模板（可选）")
            self.summary_button_text.set("选择模板")
            self.run_button_text.set("开始汇总")
        elif self.current_tool == "folder_rename":
            self.tool_title.set("需求8：人员资料文件夹改名")
            self.tool_description.set("选择人员资料目录，先预览，再确认改名。")
            self.input_label.set("人员文件夹目录")
            self.choose_input_text.set("选择文件夹")
            self.summary_label.set("")
            self.summary_button_text.set("选择")
            self.run_button_text.set("预览")
        elif self.current_tool == "salary_split":
            self.tool_title.set("需求4：工资表按入职公司拆分")
            self.tool_description.set("选择一个包含“汇总表”和“明细表”的工资表，工具会按“入职公司”拆成多个公司文件。")
            self.input_label.set("工资表文件")
            self.choose_input_text.set("选择文件")
            self.summary_label.set("")
            self.summary_button_text.set("选择")
            self.run_button_text.set("开始拆分")
        else:
            self.tool_title.set("该工具暂未实现")
            self.tool_description.set("请选择左侧已经可用的工具。")
            self.input_label.set("输入")
            self.choose_input_text.set("选择")
            self.summary_label.set("")
            self.summary_button_text.set("选择")
            self.run_button_text.set("开始")
        if hasattr(self, "summary_label_widget"):
            self._update_summary_controls()
            self._update_output_controls()
            self._update_rename_controls()
        if hasattr(self, "tutorial_text"):
            self._set_tutorial_text()

    def _set_tutorial_text(self) -> None:
        self.tutorial_text.config(state="normal")
        self.tutorial_text.delete("1.0", END)
        for line, tag in self._tutorial_lines():
            if tag:
                self.tutorial_text.insert(END, line + "\n", tag)
            else:
                self.tutorial_text.insert(END, line + "\n")
        self.tutorial_text.config(state="disabled")

    def _tutorial_lines(self) -> list[tuple[str, str | None]]:
        if self.current_tool == "salary_merge":
            return [
                ("适用：把 1-12 个月工资表合成一张个人应发工资汇总表。", "strong"),
                ("步骤：把月度工资表放进同一个文件夹，选择该文件夹；如已有前几月汇总表，再选择“已有汇总表”。", None),
                ("点击“开始合并”后，结果会生成到保存位置下的新文件夹中。", None),
                ("结果：按姓名、身份证号、月份合并；没有工资的月份填 0；已存在的人员月份不会覆盖。", None),
                ("注意：工资表文件名或表内日期要能识别月份；重复人员或重复月份会在执行结果里提醒。", "warning"),
            ]
        if self.current_tool == "personnel_change_merge":
            return [
                ("适用：把多个项目异动表合成一份异动汇总表。", "strong"),
                ("步骤：把各项目异动表放进同一个文件夹，选择该文件夹；有固定模板时再选择“异动表模板”。", None),
                ("点击“开始汇总”后，结果会生成到保存位置下的新文件夹中。", None),
                ("结果：汇总增员、减员、转正、调动、奖罚扣补，并自动重新编号。", None),
                ("注意：需求6目前暂定，正式使用前要确认各项目异动表格式一致。", "warning"),
            ]
        if self.current_tool == "folder_rename":
            return [
                ("适用：批量修改所选目录下第一层人员文件夹名称。", "strong"),
                ("追加文字：姓名不填就是全部文件夹追加；填姓名就是只处理这个人。输入“劳动合同”会追加为“-劳动合同”。", None),
                ("删除结尾文字：输入“_劳动合同”，可删除“张三_劳动合同 / 张三-劳动合同 / 张三劳动合同”的结尾文字。", None),
                ("修改单人名称：填写原姓名和新名称，例如“张三”改为“章五”。", None),
                ("重要提醒：改名会直接改变真实文件夹名称。必须先看预览，确认无误后再点确认；建议操作前先备份。", "warning"),
            ]
        if self.current_tool == "salary_split":
            return [
                ("适用：一个完整工资表按“入职公司”拆成多个公司工资表。", "strong"),
                ("步骤：选择工资表文件，保存位置默认在桌面“工资表拆分结果”，点击“开始拆分”。", None),
                ("点击“打开所在文件夹”可直接查看本次生成的结果目录。", None),
                ("结果：每个入职公司生成一个 Excel，保留表头、格式、公式、小计和底部总计。", None),
                ("注意：源工资表不会被修改；如果模板列名或表结构变化，先发给开发确认。", "warning"),
            ]
        return [
            ("该工具暂未实现。", "strong"),
            ("请选择左侧已完成的工具：需求4、需求5、需求6、需求8。", None),
        ]

    def _update_summary_controls(self) -> None:
        if self.current_tool in {"salary_merge", "personnel_change_merge"}:
            self.summary_label_widget.grid(row=1, column=0, sticky="w", pady=4)
            self.summary_entry_widget.grid(row=1, column=1, sticky="ew", padx=8)
            self.summary_button_widget.grid(row=1, column=2, sticky="e")
            return
        self.summary_label_widget.grid_remove()
        self.summary_entry_widget.grid_remove()
        self.summary_button_widget.grid_remove()

    def _update_output_controls(self) -> None:
        if self.current_tool == "folder_rename":
            self.output_label_widget.grid_remove()
            self.output_entry_widget.grid_remove()
            self.output_button_widget.grid_remove()
            return
        self.output_label_widget.grid(row=2, column=0, sticky="w", pady=4)
        self.output_entry_widget.grid(row=2, column=1, sticky="ew", padx=8)
        self.output_button_widget.grid(row=2, column=2, sticky="e")

    def _update_rename_controls(self) -> None:
        if self.current_tool != "folder_rename":
            self.rename_options_frame.grid_remove()
            return
        self.rename_options_frame.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        self._update_rename_mode_controls()

    def _on_rename_mode_changed(self, _event=None) -> None:
        self._update_rename_mode_controls()

    def _update_rename_mode_controls(self) -> None:
        mode = RENAME_MODE_LABELS.get(self.rename_mode.get(), MODE_APPEND)
        if mode == MODE_APPEND:
            self.rename_target_label.set("姓名（可不填）")
            self.rename_text_label.set("要追加的文字")
            self.rename_text_widget.config(state="normal")
            self.rename_replacement_widget.config(state="disabled")
        elif mode == MODE_REMOVE:
            self.rename_target_label.set("姓名（可不填）")
            self.rename_text_label.set("要删除的结尾文字")
            self.rename_text_widget.config(state="normal")
            self.rename_replacement_widget.config(state="disabled")
        else:
            self.rename_target_label.set("原姓名")
            self.rename_text_label.set("不用填")
            self.rename_text_widget.config(state="disabled")
            self.rename_replacement_widget.config(state="normal")

    def _initial_log_text(self) -> str:
        if self.current_tool == "salary_merge":
            return "请选择工资表文件夹和保存位置，然后点击“开始合并”。已有汇总表是可选项，用于追加新月份。"
        if self.current_tool == "personnel_change_merge":
            return "请选择异动表文件夹和保存位置，然后点击“开始汇总”。模板是可选项，不选时使用文件夹中的第一份异动表作为模板。"
        if self.current_tool == "folder_rename":
            return "请选择人员文件夹目录，填写改名内容，然后点击“预览”。"
        if self.current_tool == "salary_split":
            return "请选择工资表文件和保存位置，然后点击“开始拆分”。"
        return "该工具暂未实现。"

    def _choose_input(self) -> None:
        if self.current_tool in {"salary_merge", "personnel_change_merge", "folder_rename"}:
            if self.current_tool == "personnel_change_merge":
                title = "选择异动表文件夹"
            elif self.current_tool == "folder_rename":
                title = "选择人员文件夹目录"
            else:
                title = "选择工资表文件夹"
            directory = filedialog.askdirectory(title=title)
            if directory:
                self.input_path.set(directory)
                if not self.output_dir_user_selected:
                    self.output_dir.set(str(default_output_parent_dir(self.current_tool)))
            return

        filename = filedialog.askopenfilename(
            title="选择工资表",
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
        )
        if filename:
            self.input_path.set(filename)
            if not self.output_dir_user_selected:
                self.output_dir.set(str(default_output_parent_dir(self.current_tool)))

    def _choose_output(self) -> None:
        directory = filedialog.askdirectory(title="选择保存位置")
        if directory:
            self.output_dir_user_selected = True
            self.output_dir.set(directory)

    def _choose_summary(self) -> None:
        title = "选择异动表模板" if self.current_tool == "personnel_change_merge" else "选择已有汇总表"
        filename = filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
        )
        if filename:
            self.summary_path.set(filename)

    def _run_current_tool(self) -> None:
        if self.current_tool == "archive_import":
            messagebox.showinfo("暂未实现", "该工具还在开发中。")
            return
        if self.current_tool == "folder_rename":
            self._run_folder_rename()
            return
        if self.current_tool == "personnel_change_merge":
            self._run_personnel_change_merge()
            return
        if self.current_tool == "salary_merge":
            self._run_salary_merge()
            return
        self._run_salary_split()

    def _run_salary_split(self) -> None:
        input_text = self.input_path.get().strip()
        output_text = self.output_dir.get().strip()
        if not input_text:
            messagebox.showwarning("缺少文件", "请先选择工资表文件。")
            return
        input_path = Path(input_text)
        if not input_path.exists():
            messagebox.showwarning("文件不存在", "选择的工资表文件不存在，请重新选择。")
            return
        if input_path.suffix.lower() != ".xlsx":
            messagebox.showwarning("格式不支持", "当前工具只支持 .xlsx 工资表。")
            return
        if not output_text:
            messagebox.showwarning("缺少目录", "请选择保存位置。")
            return
        output_parent_dir = Path(output_text)

        output_dir = make_result_output_dir(output_parent_dir)
        self.run_button.config(state="disabled")
        self._clear_log()
        self._write_log("开始拆分，请稍候...")

        worker = threading.Thread(
            target=self._salary_split_worker,
            args=(input_path, output_dir),
            daemon=True,
        )
        worker.start()

    def _run_salary_merge(self) -> None:
        input_text = self.input_path.get().strip()
        summary_text = self.summary_path.get().strip()
        summary_path = Path(summary_text) if summary_text else None
        output_text = self.output_dir.get().strip()
        if not input_text:
            messagebox.showwarning("缺少文件夹", "请先选择工资表文件夹。")
            return
        input_dir = Path(input_text)
        if not input_dir.exists() or not input_dir.is_dir():
            messagebox.showwarning("文件夹不存在", "选择的工资表文件夹不存在，请重新选择。")
            return
        if summary_path is not None and not summary_path.exists():
            messagebox.showwarning("汇总表不存在", "选择的已有汇总表不存在，请重新选择。")
            return
        if summary_path is not None and summary_path.suffix.lower() != ".xlsx":
            messagebox.showwarning("格式不支持", "已有汇总表只支持 .xlsx 文件。")
            return
        if not output_text:
            messagebox.showwarning("缺少目录", "请选择保存位置。")
            return
        output_parent_dir = Path(output_text)

        output_dir = make_result_output_dir(output_parent_dir)
        self.run_button.config(state="disabled")
        self._clear_log()
        self._write_log("开始合并，请稍候...")

        worker = threading.Thread(
            target=self._salary_merge_worker,
            args=(input_dir, output_dir, summary_path),
            daemon=True,
        )
        worker.start()

    def _run_personnel_change_merge(self) -> None:
        input_text = self.input_path.get().strip()
        template_text = self.summary_path.get().strip()
        template_path = Path(template_text) if template_text else None
        output_text = self.output_dir.get().strip()
        if not input_text:
            messagebox.showwarning("缺少文件夹", "请先选择异动表文件夹。")
            return
        input_dir = Path(input_text)
        if not input_dir.exists() or not input_dir.is_dir():
            messagebox.showwarning("文件夹不存在", "选择的异动表文件夹不存在，请重新选择。")
            return
        if template_path is not None and not template_path.exists():
            messagebox.showwarning("模板不存在", "选择的异动表模板不存在，请重新选择。")
            return
        if template_path is not None and template_path.suffix.lower() != ".xlsx":
            messagebox.showwarning("格式不支持", "异动表模板只支持 .xlsx 文件。")
            return
        if not output_text:
            messagebox.showwarning("缺少目录", "请选择保存位置。")
            return
        output_parent_dir = Path(output_text)

        output_dir = make_result_output_dir(output_parent_dir)
        self.run_button.config(state="disabled")
        self._clear_log()
        self._write_log("开始汇总，请稍候...")

        worker = threading.Thread(
            target=self._personnel_change_merge_worker,
            args=(input_dir, output_dir, template_path),
            daemon=True,
        )
        worker.start()

    def _run_folder_rename(self) -> None:
        input_text = self.input_path.get().strip()
        if not input_text:
            messagebox.showwarning("缺少文件夹", "请先选择人员文件夹目录。")
            return
        root_dir = Path(input_text)
        if not root_dir.exists() or not root_dir.is_dir():
            messagebox.showwarning("文件夹不存在", "选择的人员文件夹目录不存在，请重新选择。")
            return

        mode = RENAME_MODE_LABELS.get(self.rename_mode.get(), MODE_APPEND)
        try:
            preview = rename_person_folders(
                root_dir=root_dir,
                mode=mode,
                text=self.rename_text.get(),
                target_name=self.rename_target_name.get(),
                replacement_name=self.rename_replacement_name.get(),
                dry_run=True,
            )
        except Exception as exc:
            messagebox.showerror("预览失败", str(exc))
            return

        self._clear_log()
        self._write_log("预览结果：")
        self._write_folder_rename_preview(preview)
        if preview.operation_count == 0:
            messagebox.showinfo("没有可改名文件夹", "没有找到需要改名的文件夹，请检查输入内容。")
            return

        message = self._folder_rename_confirm_message(preview)
        if not messagebox.askyesno("确认改名", message):
            self._write_log("已取消执行。")
            return

        self.run_button.config(state="disabled")
        self._write_log("开始执行改名...")
        worker = threading.Thread(
            target=self._folder_rename_worker,
            args=(root_dir, mode, self.rename_text.get(), self.rename_target_name.get(), self.rename_replacement_name.get()),
            daemon=True,
        )
        worker.start()

    def _salary_split_worker(self, input_path: Path, output_dir: Path) -> None:
        try:
            result = split_salary_by_company(input_path, output_dir)
        except Exception as exc:
            self.status_queue.put(("error", exc))
            return
        self.status_queue.put(("success", result))

    def _salary_merge_worker(self, input_dir: Path, output_dir: Path, summary_path: Path | None) -> None:
        try:
            result = merge_monthly_salary(input_dir, output_dir, existing_summary_path=summary_path)
        except Exception as exc:
            self.status_queue.put(("error", exc))
            return
        self.status_queue.put(("success", result))

    def _personnel_change_merge_worker(self, input_dir: Path, output_dir: Path, template_path: Path | None) -> None:
        try:
            result = merge_personnel_changes(input_dir, output_dir, template_path=template_path)
        except Exception as exc:
            self.status_queue.put(("error", exc))
            return
        self.status_queue.put(("success", result))

    def _folder_rename_worker(
        self,
        root_dir: Path,
        mode: str,
        text: str,
        target_name: str,
        replacement_name: str,
    ) -> None:
        try:
            result = rename_person_folders(
                root_dir=root_dir,
                mode=mode,
                text=text,
                target_name=target_name,
                replacement_name=replacement_name,
            )
        except Exception as exc:
            self.status_queue.put(("error", exc))
            return
        self.status_queue.put(("success", result))

    def _poll_status_queue(self) -> None:
        try:
            while True:
                status, payload = self.status_queue.get_nowait()
                if status == "success":
                    self._handle_success(payload)
                elif status == "error":
                    self._handle_error(payload)
        except queue.Empty:
            pass
        self.root.after(150, self._poll_status_queue)

    def _handle_success(self, result) -> None:
        payload = result.to_dict()
        if self.current_tool == "folder_rename":
            self.last_output_dir = Path(payload["root_dir"])
            self._write_log("改名完成。")
            self._write_folder_rename_preview(result)
            message = f"已完成 {payload['operation_count']} 个文件夹改名。"
        else:
            self.last_output_dir = Path(payload["output_dir"])

            if self.current_tool == "salary_merge":
                self._write_log("合并完成。")
                if payload.get("existing_summary_path"):
                    self._write_log(f"已有汇总表：{payload['existing_summary_path']}")
                self._write_log(f"识别文件数：{payload['source_file_count']}")
                self._write_log(f"识别人员数：{payload['employee_count']}")
                self._write_log(f"工资记录数：{payload['record_count']}")
                self._write_log(f"本次写入记录数：{payload['applied_record_count']}")
                self._write_log(f"已存在未覆盖记录数：{payload['skipped_record_count']}")
                self._write_log(f"输出：{payload['output_file']}")
                for warning in payload["warnings"]:
                    self._write_log(f"提醒：{warning}")
                message = "工资表已合并完成，可以打开结果文件夹查看。"
            elif self.current_tool == "personnel_change_merge":
                self._write_log("汇总完成。")
                self._write_log(f"识别文件数：{payload['source_file_count']}")
                self._write_log(f"异动记录数：{payload['record_count']}")
                for sheet_name, count in payload["sheet_counts"].items():
                    self._write_log(f"- {sheet_name}：{count} 条")
                self._write_log(f"输出：{payload['output_file']}")
                for warning in payload["warnings"]:
                    self._write_log(f"提醒：{warning}")
                message = "异动表已汇总完成，可以打开结果文件夹查看。"
            elif self.current_tool == "salary_split":
                self._write_log("拆分完成。")
                self._write_log(f"识别公司数：{payload['company_count']}")
                self._write_log(f"识别人员数：{payload['employee_count']}")
                for item in payload["outputs"]:
                    self._write_log(f"- {item['company']}：{item['employee_count']} 人")
                    if item.get("file_path"):
                        self._write_log(f"  输出：{item['file_path']}")
                message = "工资表已拆分完成，可以打开结果文件夹查看。"
            else:
                message = "处理完成。"
        self.run_button.config(state="normal")
        messagebox.showinfo("处理完成", message)

    def _handle_error(self, exc: object | None) -> None:
        action = (
            "合并"
            if self.current_tool == "salary_merge"
            else "汇总"
            if self.current_tool == "personnel_change_merge"
            else "改名"
            if self.current_tool == "folder_rename"
            else "拆分"
        )
        self._write_log(f"{action}失败。")
        self._write_log(str(exc))
        self.run_button.config(state="normal")
        messagebox.showerror(f"{action}失败", str(exc))

    def _write_folder_rename_preview(self, result) -> None:
        payload = result.to_dict()
        self._write_log(f"目录：{payload['root_dir']}")
        self._write_log(f"数量：{payload['operation_count']}")
        for operation in payload["operations"][:30]:
            self._write_log(f"- {operation['source_name']} -> {operation['target_name']}")
        remaining = payload["operation_count"] - 30
        if remaining > 0:
            self._write_log(f"... 还有 {remaining} 条")
        for warning in payload["warnings"]:
            self._write_log(f"提醒：{warning}")

    def _folder_rename_confirm_message(self, result) -> str:
        payload = result.to_dict()
        lines = [f"确认改名 {payload['operation_count']} 个文件夹："]
        for operation in payload["operations"][:8]:
            lines.append(f"{operation['source_name']} -> {operation['target_name']}")
        remaining = payload["operation_count"] - 8
        if remaining > 0:
            lines.append(f"... 还有 {remaining} 条")
        lines.append("")
        lines.append("确认后会直接改名，是否继续？")
        return "\n".join(lines)

    def _open_output_dir(self) -> None:
        directory_text = self.output_dir.get().strip()
        directory = self.last_output_dir
        if self.current_tool == "folder_rename" and directory is None:
            input_text = self.input_path.get().strip()
            directory = Path(input_text) if input_text else None
            if directory is None:
                messagebox.showwarning("缺少目录", "请先选择人员文件夹目录。")
                return
        if directory is None and directory_text:
            directory = make_result_output_dir(Path(directory_text))
        if directory is None:
            messagebox.showwarning("缺少目录", "请先选择保存位置。")
            return
        if self.current_tool == "folder_rename" and not directory.exists():
            messagebox.showwarning("目录不存在", "选择的人员文件夹目录不存在。")
            return
        try:
            if self.current_tool != "folder_rename":
                directory.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            messagebox.showerror("无法创建目录", str(exc))
            return
        open_path(directory)

    def _write_log(self, text: str) -> None:
        self.log_text.insert(END, text + "\n")
        self.log_text.see(END)

    def _clear_log(self) -> None:
        self.log_text.delete("1.0", END)


def _default_result_dir_name() -> str:
    return "结果_" + datetime.now().strftime("%Y%m%d_%H%M%S")


def default_output_parent_dir(tool: str) -> Path:
    if tool == "salary_merge":
        folder_name = "工资合并结果"
    elif tool == "personnel_change_merge":
        folder_name = "异动表汇总结果"
    else:
        folder_name = "工资表拆分结果"
    return desktop_dir() / folder_name


def make_result_output_dir(parent_dir: Path) -> Path:
    return parent_dir / _default_result_dir_name()


def desktop_dir() -> Path:
    home = Path.home()
    desktop = home / "Desktop"
    if desktop.exists():
        return desktop
    return home / "桌面"


def open_path(path: Path) -> None:
    if sys.platform.startswith("win"):
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def main() -> None:
    root = Tk()
    HRToolkitApp(root)
    root.mainloop()
