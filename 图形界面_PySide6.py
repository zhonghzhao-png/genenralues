# -*- coding: utf-8 -*-
"""简易 PySide6 桌面界面。只负责配置考试总表和启动原有分析流程。"""
import ast
import os
import sys
import subprocess
from pathlib import Path
from pprint import pformat

from PySide6.QtCore import Qt, QProcess
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QFileDialog, QMessageBox,
    QTableWidget, QTableWidgetItem, QTextEdit, QHeaderView, QCheckBox
)

CONFIG_FILE = "数据分析.py"
OUTPUT_FILE = "数据分析.xlsx"
FIELDS = ["录入", "分层", "对比", "展示", "排名"]
FIELD_LABELS = ["录入", "参与分层", "显示分层表", "控制页展示", "排名对比"]


def read_config(path):
    source = Path(path).read_text(encoding="utf-8-sig")
    tree = ast.parse(source, filename=str(path))
    for node in tree.body:
        if isinstance(node, ast.Assign) and any(
            isinstance(t, ast.Name) and t.id == "考试总表" for t in node.targets
        ):
            exams = ast.literal_eval(node.value)
            if not isinstance(exams, dict):
                raise ValueError("考试总表不是字典")
            return source, node, exams
    raise ValueError("找不到“考试总表”配置")


def write_config(path, exams):
    source = Path(path).read_text(encoding="utf-8-sig")
    tree = ast.parse(source, filename=str(path))
    for node in tree.body:
        if isinstance(node, ast.Assign) and any(
            isinstance(t, ast.Name) and t.id == "考试总表" for t in node.targets
        ):
            lines = source.splitlines(keepends=True)
            replacement = "考试总表 = " + pformat(exams, width=110, sort_dicts=False) + "\n"
            lines[node.lineno - 1:node.end_lineno] = [replacement]
            Path(path).write_text("".join(lines), encoding="utf-8-sig")
            return
    raise ValueError("找不到“考试总表”配置")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("统考数学成绩分析")
        self.resize(1050, 720)
        self.project = Path(__file__).resolve().parent
        self.exams = {}
        self.process = QProcess(self)
        self.process.setWorkingDirectory(str(self.project))
        self.process.readyReadStandardOutput.connect(self.read_output)
        self.process.readyReadStandardError.connect(self.read_error)
        self.process.finished.connect(self.finished)
        self.process.errorOccurred.connect(self.process_error)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        top = QHBoxLayout()
        top.addWidget(QLabel("项目文件夹："))
        self.folder = QLineEdit(str(self.project))
        top.addWidget(self.folder, 1)
        choose = QPushButton("浏览…")
        choose.clicked.connect(self.choose_folder)
        top.addWidget(choose)
        layout.addLayout(top)

        layout.addWidget(QLabel("考试配置：双击开关单元格切换；选中考试后可更换成绩文件。"))
        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels(["考试名称", "成绩文件路径"] + FIELD_LABELS)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        for col in range(2, 7):
            self.table.horizontalHeader().setSectionResizeMode(col, QHeaderView.ResizeToContents)
        self.table.cellDoubleClicked.connect(self.toggle_cell)
        layout.addWidget(self.table, 1)

        buttons = QHBoxLayout()
        for label, callback in [
            ("选择成绩文件", self.choose_exam_file),
            ("保存配置", self.save),
            ("重新载入", self.load),
            ("运行分析", self.run_analysis),
            ("打开结果 Excel", self.open_output),
        ]:
            button = QPushButton(label)
            button.clicked.connect(callback)
            buttons.addWidget(button)
        layout.addLayout(buttons)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setPlaceholderText("运行信息和错误会显示在这里")
        layout.addWidget(self.log, 1)

        self.statusBar().showMessage("就绪")
        self.load()

    def project_dir(self):
        return Path(self.folder.text()).expanduser().resolve()

    def log_line(self, text):
        self.log.append(str(text).rstrip())

    def choose_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择项目文件夹", self.folder.text())
        if folder:
            self.folder.setText(folder)
            self.project = Path(folder)
            self.load()

    def load(self):
        try:
            path = self.project_dir() / CONFIG_FILE
            _, _, self.exams = read_config(path)
            self.refresh_table()
            self.log_line(f"已载入配置：{path}")
            self.statusBar().showMessage(f"已载入 {len(self.exams)} 场考试")
        except Exception as exc:
            QMessageBox.critical(self, "读取配置失败", str(exc))

    def refresh_table(self):
        self.table.setRowCount(len(self.exams))
        for row, (name, info) in enumerate(self.exams.items()):
            values = [name, info.get("路径", "")]
            for col, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(row, col, item)
            for index, field in enumerate(FIELDS):
                item = QTableWidgetItem("✓" if info.get(field, False) else "")
                item.setTextAlignment(Qt.AlignCenter)
                item.setData(Qt.UserRole, field)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(row, index + 2, item)

    def toggle_cell(self, row, col):
        if col < 2:
            return
        name = self.table.item(row, 0).text()
        field = FIELDS[col - 2]
        info = self.exams[name]
        info[field] = not bool(info.get(field, False))
        if field == "分层" and info[field]:
            info["录入"] = True
        elif field == "录入" and not info[field]:
            info["分层"] = False
        self.refresh_table()
        self.table.selectRow(row)

    def selected_exam(self):
        row = self.table.currentRow()
        if row < 0:
            return None
        return self.table.item(row, 0).text()

    def choose_exam_file(self):
        name = self.selected_exam()
        if not name:
            QMessageBox.information(self, "提示", "请先选择一场考试。")
            return
        current = self.exams[name].get("路径", "")
        initial = self.project_dir() / current if current else self.project_dir()
        filename, _ = QFileDialog.getOpenFileName(
            self, "选择成绩文件", str(initial.parent),
            "Excel 文件 (*.xls *.xlsx *.xlsm);;所有文件 (*)"
        )
        if filename:
            try:
                self.exams[name]["路径"] = os.path.relpath(filename, self.project_dir())
            except ValueError:
                self.exams[name]["路径"] = filename
            self.refresh_table()
            self.table.selectRow(list(self.exams).index(name))

    def save(self):
        try:
            write_config(self.project_dir() / CONFIG_FILE, self.exams)
            self.log_line("考试配置已保存。")
            self.statusBar().showMessage("配置已保存")
            QMessageBox.information(self, "保存成功", "考试总表已写回 数据分析.py。其他高级参数仍可在该文件中修改。")
        except Exception as exc:
            QMessageBox.critical(self, "保存失败", str(exc))

    def run_analysis(self):
        if self.process.state() != QProcess.NotRunning:
            QMessageBox.information(self, "提示", "分析程序正在运行。")
            return
        project = self.project_dir()
        script = project / CONFIG_FILE
        if not script.is_file() or not (project / "主程序.py").is_file():
            QMessageBox.critical(self, "文件不完整", "项目文件夹中必须同时有 数据分析.py 和 主程序.py。")
            return
        answer = QMessageBox.question(
            self, "确认运行",
            "将运行原有成绩分析流程。请先关闭 Excel 中的 数据分析.xlsx。\n是否继续？"
        )
        if answer != QMessageBox.Yes:
            return
        try:
            self.process.setWorkingDirectory(str(project))
            self.process.start(sys.executable, ["-u", str(script)])
            self.statusBar().showMessage("正在运行…")
            self.log_line("\n========== 开始运行 ==========")
        except Exception as exc:
            QMessageBox.critical(self, "启动失败", str(exc))

    def read_output(self):
        data = bytes(self.process.readAllStandardOutput()).decode("utf-8", errors="replace")
        if data:
            self.log.insertPlainText(data)
            self.log.ensureCursorVisible()

    def read_error(self):
        data = bytes(self.process.readAllStandardError()).decode("utf-8", errors="replace")
        if data:
            self.log.insertPlainText(data)
            self.log.ensureCursorVisible()

    def finished(self, code, _status):
        self.statusBar().showMessage("运行完成" if code == 0 else f"运行失败，退出代码 {code}")
        self.log_line(f"\n========== 程序结束，退出代码：{code} ==========")

    def process_error(self, _error):
        self.log_line(f"启动进程失败：{self.process.errorString()}")
        self.statusBar().showMessage("启动失败")

    def open_output(self):
        path = self.project_dir() / OUTPUT_FILE
        if not path.exists():
            QMessageBox.information(self, "提示", "尚未生成 数据分析.xlsx。")
            return
        try:
            os.startfile(str(path))
        except Exception as exc:
            QMessageBox.critical(self, "打开失败", str(exc))


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
