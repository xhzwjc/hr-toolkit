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
from hr_toolkit.tools.salary_merge import merge_monthly_salary
from hr_toolkit.tools.salary_split import split_salary_by_company


class HRToolkitApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title(f"HR工具箱 v{__version__}")
        self.root.geometry("760x500")
        self.root.minsize(720, 460)

        self.current_tool = "salary_split"
        self.tool_title = StringVar()
        self.tool_description = StringVar()
        self.input_label = StringVar()
        self.choose_input_text = StringVar()
        self.run_button_text = StringVar()
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

        form = ttk.Frame(right_frame)
        form.pack(fill="x")

        ttk.Label(form, textvariable=self.input_label).grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(form, textvariable=self.input_path).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(form, textvariable=self.choose_input_text, command=self._choose_input).grid(row=0, column=2, sticky="e")

        self.summary_label_widget = ttk.Label(form, text="已有汇总表（可选）")
        self.summary_entry_widget = ttk.Entry(form, textvariable=self.summary_path)
        self.summary_button_widget = ttk.Button(form, text="选择汇总表", command=self._choose_summary)

        ttk.Label(form, text="保存位置").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(form, textvariable=self.output_dir).grid(row=2, column=1, sticky="ew", padx=8)
        ttk.Button(form, text="选择目录", command=self._choose_output).grid(row=2, column=2, sticky="e")
        form.columnconfigure(1, weight=1)
        self._update_summary_controls()

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
            self.run_button_text.set("开始合并")
        elif self.current_tool == "salary_split":
            self.tool_title.set("需求4：工资表按入职公司拆分")
            self.tool_description.set("选择一个包含“汇总表”和“明细表”的工资表，工具会按“入职公司”拆成多个公司文件。")
            self.input_label.set("工资表文件")
            self.choose_input_text.set("选择文件")
            self.run_button_text.set("开始拆分")
        else:
            self.tool_title.set("该工具暂未实现")
            self.tool_description.set("请选择左侧已经可用的工具。")
            self.input_label.set("输入")
            self.choose_input_text.set("选择")
            self.run_button_text.set("开始")
        if hasattr(self, "summary_label_widget"):
            self._update_summary_controls()

    def _update_summary_controls(self) -> None:
        if self.current_tool == "salary_merge":
            self.summary_label_widget.grid(row=1, column=0, sticky="w", pady=4)
            self.summary_entry_widget.grid(row=1, column=1, sticky="ew", padx=8)
            self.summary_button_widget.grid(row=1, column=2, sticky="e")
            return
        self.summary_label_widget.grid_remove()
        self.summary_entry_widget.grid_remove()
        self.summary_button_widget.grid_remove()

    def _initial_log_text(self) -> str:
        if self.current_tool == "salary_merge":
            return "请选择工资表文件夹和保存位置，然后点击“开始合并”。已有汇总表是可选项，用于追加新月份。"
        if self.current_tool == "salary_split":
            return "请选择工资表文件和保存位置，然后点击“开始拆分”。"
        return "该工具暂未实现。"

    def _choose_input(self) -> None:
        if self.current_tool == "salary_merge":
            directory = filedialog.askdirectory(title="选择工资表文件夹")
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
        filename = filedialog.askopenfilename(
            title="选择已有汇总表",
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
        )
        if filename:
            self.summary_path.set(filename)

    def _run_current_tool(self) -> None:
        if self.current_tool == "archive_import":
            messagebox.showinfo("暂未实现", "该工具还在开发中。")
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
        else:
            self._write_log("拆分完成。")
            self._write_log(f"识别公司数：{payload['company_count']}")
            self._write_log(f"识别人员数：{payload['employee_count']}")
            for item in payload["outputs"]:
                self._write_log(f"- {item['company']}：{item['employee_count']} 人")
                if item.get("file_path"):
                    self._write_log(f"  输出：{item['file_path']}")
            message = "工资表已拆分完成，可以打开结果文件夹查看。"
        self.run_button.config(state="normal")
        messagebox.showinfo("处理完成", message)

    def _handle_error(self, exc: object | None) -> None:
        action = "合并" if self.current_tool == "salary_merge" else "拆分"
        self._write_log(f"{action}失败。")
        self._write_log(str(exc))
        self.run_button.config(state="normal")
        messagebox.showerror(f"{action}失败", str(exc))

    def _open_output_dir(self) -> None:
        directory_text = self.output_dir.get().strip()
        directory = self.last_output_dir
        if directory is None and directory_text:
            directory = make_result_output_dir(Path(directory_text))
        if directory is None:
            messagebox.showwarning("缺少目录", "请先选择保存位置。")
            return
        try:
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
    folder_name = "工资合并结果" if tool == "salary_merge" else "工资表拆分结果"
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
