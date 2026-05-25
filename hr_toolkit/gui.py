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
from hr_toolkit.tools.salary_split import split_salary_by_company


class HRToolkitApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title(f"HR工具箱 v{__version__}")
        self.root.geometry("760x500")
        self.root.minsize(720, 460)

        self.input_path = StringVar()
        self.output_dir = StringVar(value=str(default_output_parent_dir()))
        self.output_dir_user_selected = False
        self.status_queue: queue.Queue[tuple[str, object | None]] = queue.Queue()
        self.last_output_dir: Path | None = None

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
        self.tool_list.insert("", END, iid="salary_merge", text="需求5 工资表合并（待实现）")
        self.tool_list.insert("", END, iid="archive_import", text="需求7 档案入库（待实现）")
        self.tool_list.selection_set("salary_split")

        right_frame = ttk.Frame(root_frame)
        right_frame.pack(side=RIGHT, fill=BOTH, expand=True)

        ttk.Label(
            right_frame,
            text="需求4：工资表按入职公司拆分",
            font=("", 13, "bold"),
        ).pack(anchor="w")
        ttk.Label(
            right_frame,
            text="选择一个包含“汇总表”和“明细表”的工资表，工具会按“入职公司”拆成多个公司文件。",
        ).pack(anchor="w", pady=(4, 14))

        form = ttk.Frame(right_frame)
        form.pack(fill="x")

        ttk.Label(form, text="工资表文件").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(form, textvariable=self.input_path).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(form, text="选择文件", command=self._choose_input).grid(row=0, column=2, sticky="e")

        ttk.Label(form, text="保存位置").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(form, textvariable=self.output_dir).grid(row=1, column=1, sticky="ew", padx=8)
        ttk.Button(form, text="选择目录", command=self._choose_output).grid(row=1, column=2, sticky="e")
        form.columnconfigure(1, weight=1)

        actions = ttk.Frame(right_frame)
        actions.pack(fill="x", pady=(14, 10))
        self.run_button = ttk.Button(actions, text="开始拆分", command=self._run_salary_split)
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

        self._write_log("请选择工资表文件和保存位置，然后点击“开始拆分”。")

    def _choose_input(self) -> None:
        filename = filedialog.askopenfilename(
            title="选择工资表",
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
        )
        if filename:
            self.input_path.set(filename)
            if not self.output_dir_user_selected:
                self.output_dir.set(str(default_output_parent_dir()))

    def _choose_output(self) -> None:
        directory = filedialog.askdirectory(title="选择保存位置")
        if directory:
            self.output_dir_user_selected = True
            self.output_dir.set(directory)

    def _run_salary_split(self) -> None:
        input_path = Path(self.input_path.get().strip())
        output_parent_dir = Path(self.output_dir.get().strip())
        if not input_path:
            messagebox.showwarning("缺少文件", "请先选择工资表文件。")
            return
        if not input_path.exists():
            messagebox.showwarning("文件不存在", "选择的工资表文件不存在，请重新选择。")
            return
        if input_path.suffix.lower() != ".xlsx":
            messagebox.showwarning("格式不支持", "当前工具只支持 .xlsx 工资表。")
            return
        if not output_parent_dir:
            messagebox.showwarning("缺少目录", "请选择保存位置。")
            return

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

    def _salary_split_worker(self, input_path: Path, output_dir: Path) -> None:
        try:
            result = split_salary_by_company(input_path, output_dir)
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
        self._write_log("拆分完成。")
        self._write_log(f"识别公司数：{payload['company_count']}")
        self._write_log(f"识别人员数：{payload['employee_count']}")
        for item in payload["outputs"]:
            self._write_log(f"- {item['company']}：{item['employee_count']} 人")
            if item.get("file_path"):
                self._write_log(f"  输出：{item['file_path']}")
        self.run_button.config(state="normal")
        messagebox.showinfo("拆分完成", "工资表已拆分完成，可以打开结果文件夹查看。")

    def _handle_error(self, exc: object | None) -> None:
        self._write_log("拆分失败。")
        self._write_log(str(exc))
        self.run_button.config(state="normal")
        messagebox.showerror("拆分失败", str(exc))

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


def default_output_parent_dir() -> Path:
    return desktop_dir() / "工资表拆分结果"


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
