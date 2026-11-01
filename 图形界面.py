# -*- coding: utf-8 -*-
"""成绩分析程序图形界面。仅管理考试总表并调用原有数据分析入口，不改动核心计算。"""
import ast
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from pprint import pformat


APP_TITLE = "统考数学成绩分析工具"
CONFIG_FILE = "数据分析.py"
OUTPUT_FILE = "数据分析.xlsx"
FLAG_FIELDS = ["录入", "分层", "对比", "展示", "排名"]
FLAG_LABELS = {
    "录入": "录入",
    "分层": "参与分层",
    "对比": "显示分层表",
    "展示": "控制页展示",
    "排名": "排名对比",
}


def read_exam_config(config_path):
    """从配置源码中安全提取考试总表字典，不执行配置文件。"""
    source = Path(config_path).read_text(encoding="utf-8-sig")
    tree = ast.parse(source, filename=str(config_path))
    for node in tree.body:
        if isinstance(node, ast.Assign):
            if any(isinstance(t, ast.Name) and t.id == "考试总表" for t in node.targets):
                value = ast.literal_eval(node.value)
                if not isinstance(value, dict):
                    raise ValueError("考试总表必须是字典。")
                return source, node, value
    raise ValueError("在数据分析.py 中找不到“考试总表”配置。")


def save_exam_config(config_path, source, node, exams):
    """只替换考试总表赋值区域，保留其余配置和程序代码。"""
    lines = source.splitlines(keepends=True)
    start = node.lineno - 1
    end = node.end_lineno
    indent = " " * node.col_offset
    replacement = indent + "考试总表 = " + pformat(exams, width=110, sort_dicts=False)
    lines[start:end] = [replacement + "\n"]
    Path(config_path).write_text("".join(lines), encoding="utf-8-sig")


class ScoreAnalysisGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1180x780")
        self.root.minsize(920, 620)
        self.project_dir = Path(__file__).resolve().parent
        self.source = ""
        self.exam_node = None
        self.exams = {}
        self.process = None
        self.log_queue = queue.Queue()
        self._build_ui()
        self.load_config()
        self.root.after(100, self._poll_log)

    def _build_ui(self):
        style = ttk.Style()
        try:
            style.theme_use("vista" if sys.platform == "win32" else "clam")
        except tk.TclError:
            pass
        style.configure("Treeview", rowheight=28, font=("Microsoft YaHei UI", 9))
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 9, "bold"))
        style.configure("TButton", padding=(9, 5))
        style.configure("Header.TLabel", font=("Microsoft YaHei UI", 15, "bold"))

        outer = ttk.Frame(self.root, padding=12)
        outer.pack(fill="both", expand=True)

        header = ttk.Frame(outer)
        header.pack(fill="x", pady=(0, 10))
        ttk.Label(header, text=APP_TITLE, style="Header.TLabel").pack(side="left")
        ttk.Label(header, text="图形化配置 · 一键运行 · 实时日志").pack(side="right")

        project = ttk.LabelFrame(outer, text="项目位置", padding=8)
        project.pack(fill="x", pady=(0, 9))
        self.project_var = tk.StringVar(value=str(self.project_dir))
        ttk.Entry(project, textvariable=self.project_var).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(project, text="选择文件夹", command=self.choose_project).pack(side="left")
        ttk.Button(project, text="打开文件夹", command=self.open_project).pack(side="left", padx=(6, 0))

        exam_frame = ttk.LabelFrame(outer, text="考试配置（双击开关列可切换；选中考试后可修改成绩文件）", padding=8)
        exam_frame.pack(fill="both", expand=True, pady=(0, 9))
        columns = ("name", "path", "录入", "分层", "对比", "展示", "排名")
        self.tree = ttk.Treeview(exam_frame, columns=columns, show="headings", selectmode="browse")
        headings = {
            "name": ("考试名称", 150, "w"),
            "path": ("成绩文件路径", 480, "w"),
            "录入": ("录入", 65, "center"),
            "分层": ("参与分层", 80, "center"),
            "对比": ("显示分层表", 90, "center"),
            "展示": ("控制页展示", 90, "center"),
            "排名": ("排名对比", 80, "center"),
        }
        for key, (title, width, anchor) in headings.items():
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, minwidth=55, anchor=anchor, stretch=(key == "path"))
        ybar = ttk.Scrollbar(exam_frame, orient="vertical", command=self.tree.yview)
        xbar = ttk.Scrollbar(exam_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        exam_frame.rowconfigure(0, weight=1)
        exam_frame.columnconfigure(0, weight=1)
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        self.tree.bind("<<TreeviewSelect>>", self.on_exam_select)

        tools = ttk.Frame(outer)
        tools.pack(fill="x", pady=(0, 9))
        ttk.Button(tools, text="选择成绩文件", command=self.choose_exam_file).pack(side="left")
        ttk.Button(tools, text="保存考试配置", command=self.save_config).pack(side="left", padx=6)
        ttk.Button(tools, text="重新载入配置", command=self.load_config).pack(side="left")
        ttk.Separator(tools, orient="vertical").pack(side="left", fill="y", padx=10)
        self.run_button = ttk.Button(tools, text="▶ 运行全部流程", command=self.run_program)
        self.run_button.pack(side="left")
        ttk.Button(tools, text="打开结果 Excel", command=self.open_output).pack(side="left", padx=6)
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(tools, textvariable=self.status_var).pack(side="right")

        log_frame = ttk.LabelFrame(outer, text="运行日志", padding=8)
        log_frame.pack(fill="both", expand=True)
        self.log = tk.Text(log_frame, height=12, wrap="word", font=("Consolas", 9), state="disabled")
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=log_scroll.set)
        self.log.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")

        footer = ttk.Label(
            outer,
            text="提示：修改开关或路径后请先保存配置。运行前请关闭 Excel 中打开的“数据分析.xlsx”，避免文件被占用。",
            foreground="#555555",
        )
        footer.pack(fill="x", pady=(7, 0))

    def _project_path(self):
        return Path(self.project_var.get()).expanduser().resolve()

    def choose_project(self):
        chosen = filedialog.askdirectory(initialdir=str(self._project_path()))
        if chosen:
            self.project_var.set(chosen)
            self.project_dir = Path(chosen)
            self.load_config()

    def open_project(self):
        path = self._project_path()
        if not path.exists():
            messagebox.showerror("路径不存在", str(path))
            return
        self._open_path(path)

    def _open_path(self, path):
        try:
            if sys.platform == "win32":
                os.startfile(str(path))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as exc:
            messagebox.showerror("打开失败", str(exc))

    def load_config(self):
        path = self._project_path() / CONFIG_FILE
        if not path.is_file():
            messagebox.showerror("找不到配置文件", f"当前项目文件夹中没有：\n{path}")
            return
        try:
            self.source, self.exam_node, self.exams = read_exam_config(path)
            self.project_dir = self._project_path()
            self._refresh_tree()
            self._append_log(f"已载入配置：{path}")
            self.status_var.set(f"已载入 {len(self.exams)} 场考试")
        except Exception as exc:
            messagebox.showerror("读取配置失败", str(exc))

    def _refresh_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for name, info in self.exams.items():
            values = [name, info.get("路径", "")]
            values.extend("☑" if info.get(key, False) else "☐" for key in FLAG_FIELDS)
            self.tree.insert("", "end", iid=name, values=values)

    def _selected_name(self):
        selected = self.tree.selection()
        return selected[0] if selected else None

    def on_exam_select(self, _event=None):
        return

    def on_tree_double_click(self, event):
        item = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        if not item or not column:
            return
        col_index = int(column[1:]) - 1
        if col_index < 2:
            return
        field = FLAG_FIELDS[col_index - 2]
        info = self.exams[item]
        info[field] = not bool(info.get(field, False))
        if field == "分层" and info[field]:
            info["录入"] = True
        elif field == "录入" and not info[field]:
            # 未录入的考试不能参与综合分层，避免保存后触发配置校验错误。
            info["分层"] = False
        self._refresh_tree()
        self.tree.selection_set(item)

    def choose_exam_file(self):
        name = self._selected_name()
        if not name:
            messagebox.showinfo("选择考试", "请先在考试列表中选中一场考试。")
            return
        current = self.exams[name].get("路径", "")
        initial = self._project_path() / current if current else self._project_path()
        chosen = filedialog.askopenfilename(
            title=f"选择{name}成绩文件",
            initialdir=str(initial.parent if initial.parent.exists() else self._project_path()),
            filetypes=[("Excel 文件", "*.xls *.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if chosen:
            try:
                relative = os.path.relpath(chosen, self._project_path())
                self.exams[name]["路径"] = relative
            except ValueError:
                self.exams[name]["路径"] = chosen
            self._refresh_tree()
            self.tree.selection_set(name)

    def save_config(self):
        path = self._project_path() / CONFIG_FILE
        try:
            current_source, node, _ = read_exam_config(path)
            save_exam_config(path, current_source, node, self.exams)
            self.source, self.exam_node, self.exams = read_exam_config(path)
            self._refresh_tree()
            self._append_log("考试配置已保存到 数据分析.py")
            self.status_var.set("配置已保存")
            messagebox.showinfo("保存成功", "考试配置已保存。其他分层权重、名单路径等高级参数仍可在数据分析.py 中调整。")
        except Exception as exc:
            messagebox.showerror("保存失败", str(exc))

    def run_program(self):
        if self.process and self.process.poll() is None:
            messagebox.showinfo("正在运行", "程序已经在运行，请等待完成。")
            return
        project = self._project_path()
        script = project / CONFIG_FILE
        if not script.is_file():
            messagebox.showerror("无法运行", f"找不到程序入口：\n{script}")
            return
        if not (project / "主程序.py").is_file():
            messagebox.showerror("无法运行", "当前文件夹中找不到 主程序.py。")
            return
        if not messagebox.askyesno(
            "确认运行",
            "即将按当前 数据分析.py 配置运行完整流程。\n\n请确认已关闭 Excel 中打开的 数据分析.xlsx。\n是否继续？",
        ):
            return
        self.run_button.configure(state="disabled")
        self.status_var.set("正在运行…")
        self._append_log("\n" + "=" * 65 + "\n开始运行\n" + "=" * 65)
        threading.Thread(target=self._run_worker, args=(project, script), daemon=True).start()

    def _run_worker(self, project, script):
        try:
            creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
            self.process = subprocess.Popen(
                [sys.executable, "-u", str(script)],
                cwd=str(project),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
                creationflags=creationflags,
            )
            for line in self.process.stdout:
                self.log_queue.put(("line", line))
            code = self.process.wait()
            self.log_queue.put(("done", code))
        except Exception as exc:
            self.log_queue.put(("error", str(exc)))

    def _poll_log(self):
        try:
            while True:
                kind, value = self.log_queue.get_nowait()
                if kind == "line":
                    self._append_log(value.rstrip("\n"))
                elif kind == "done":
                    self.run_button.configure(state="normal")
                    if value == 0:
                        self.status_var.set("运行完成")
                        self._append_log("\n✓ 全部流程执行结束。")
                    else:
                        self.status_var.set(f"运行失败（退出代码 {value}）")
                        self._append_log(f"\n✗ 程序异常退出，代码：{value}")
                elif kind == "error":
                    self.run_button.configure(state="normal")
                    self.status_var.set("启动失败")
                    self._append_log(f"\n✗ 启动失败：{value}")
        except queue.Empty:
            pass
        self.root.after(100, self._poll_log)

    def _append_log(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", str(text) + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def open_output(self):
        path = self._project_path() / OUTPUT_FILE
        if not path.exists():
            messagebox.showinfo("结果尚未生成", "还没有找到 数据分析.xlsx，请先运行程序。")
            return
        self._open_path(path)


def main():
    root = tk.Tk()
    ScoreAnalysisGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
