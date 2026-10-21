"""Desktop tool for comparing two Excel workbooks."""

import sys
from pathlib import Path

import pandas as pd
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QApplication, QFileDialog, QHBoxLayout, QLabel, QListWidget,
    QListWidgetItem, QMainWindow, QMessageBox, QPushButton, QVBoxLayout, QWidget,
)

from excel_operations import compare_tables, save_result, validate_destination


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.files = []
        self.source_columns = []
        self.target_columns = []
        self.new_window = None
        self.setWindowTitle("Excel 交集与差集")
        self.resize(900, 500)
        widget = QWidget(self)
        self.setCentralWidget(widget)
        layout = QVBoxLayout(widget)
        layout.addWidget(QLabel("选择两个 Excel 文件：第一个为源表，第二个为目标表。"))
        self.qlist = QListWidget()
        layout.addWidget(self.qlist)
        buttons = QHBoxLayout()
        for text, callback in (
            ("选择文件", self.Select_multiple_files),
            ("确定", self.open_newwindow),
            ("清空列表", self.clear_list),
        ):
            button = QPushButton(text)
            button.clicked.connect(callback)
            buttons.addWidget(button)
        layout.addLayout(buttons)

    def _close_comparison(self):
        if self.new_window is not None:
            self.new_window.close()
            self.new_window.deleteLater()
            self.new_window = None

    def clear_list(self):
        self._close_comparison()
        self.files = []
        self.source_columns = []
        self.target_columns = []
        self.qlist.clear()

    def Select_multiple_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "选择文件", "", "Excel 文件 (*.xlsx *.xls)",
        )
        if not paths:
            return
        if len(paths) > 2:
            QMessageBox.warning(self, "文件数量错误", "请一次选择一个或两个文件。")
            return
        # Selecting a pair replaces the old pair; one file can complete a pair.
        candidates = self.files + paths if len(paths) == 1 and len(self.files) == 1 else paths
        try:
            headers = [pd.read_excel(path, nrows=0).columns.tolist() for path in candidates]
        except Exception as error:
            QMessageBox.critical(self, "无法读取文件", f"文件未更新，请检查工作簿。\n{error}")
            return
        # Commit only after every candidate has been read successfully.
        self._close_comparison()
        self.files = list(candidates)
        self.source_columns = headers[0]
        self.target_columns = headers[1] if len(headers) == 2 else []
        self.qlist.clear()
        for role, path in zip(("源表", "目标表"), self.files):
            self.qlist.addItem(f"{role}：{path}")

    def open_newwindow(self):
        if len(self.files) != 2:
            QMessageBox.warning(self, "文件不足", "请先选择两个 Excel 文件。")
            return
        if not set(self.source_columns).intersection(self.target_columns):
            QMessageBox.warning(self, "无法比较", "两个文件没有同名列，请先检查表头。")
            return
        self._close_comparison()
        self.new_window = NewWindow(
            self.files, self.source_columns, self.target_columns, self,
        )
        self.new_window.show()


class NewWindow(QMainWindow):
    def __init__(self, files, source_columns, target_columns, parent=None):
        super().__init__(parent)
        self.files = tuple(files)
        self.setWindowTitle("选择比较列并导出")
        self.resize(1050, 650)
        widget = QWidget(self)
        self.setCentralWidget(widget)
        layout = QVBoxLayout(widget)
        layout.addWidget(QLabel(
            "选择输出列和共同筛选列后直接计算。交集输出匹配的双方列；差集只输出源表列。"
        ))
        columns = QHBoxLayout()
        target_labels = set(target_columns)
        common = [column for column in source_columns if column in target_labels]
        self.listwidget = self._column_list(source_columns, select_all=True)
        self.listwidget_2 = self._column_list(target_columns, select_all=True)
        self.listwidget_3 = self._column_list(common)
        for label, view in (
            ("源表输出列", self.listwidget),
            ("目标表输出列（仅用于交集）", self.listwidget_2),
            ("共同筛选列（至少选择一列）", self.listwidget_3),
        ):
            box = QVBoxLayout()
            box.addWidget(QLabel(label))
            box.addWidget(view)
            columns.addLayout(box)
        layout.addLayout(columns)
        buttons = QHBoxLayout()
        for text, callback in (("求差集", self.filt), ("求交集", self.overlap)):
            button = QPushButton(text)
            button.clicked.connect(callback)
            buttons.addWidget(button)
        layout.addLayout(buttons)

    @staticmethod
    def _column_list(columns, select_all=False):
        view = QListWidget()
        view.setSelectionMode(QListWidget.MultiSelection)
        for column in columns:
            item = QListWidgetItem(str(column))
            # Preserve numeric/date labels even when display texts are identical.
            item.setData(Qt.UserRole, column)
            view.addItem(item)
            item.setSelected(select_all)
        return view

    @staticmethod
    def _selected(view):
        return [item.data(Qt.UserRole) for item in view.selectedItems()]

    def filt(self):
        self._export("difference")

    def overlap(self):
        self._export("intersection")

    def _export(self, operation):
        try:
            keys = self._selected(self.listwidget_3)
            if not keys:
                raise ValueError("请至少选择一列作为共同筛选列。")
            source, target = (pd.read_excel(path) for path in self.files)
            result = compare_tables(
                source, target, self._selected(self.listwidget),
                self._selected(self.listwidget_2), keys, operation,
            )
            title = "差集结果" if operation == "difference" else "交集结果"
            destination, _ = QFileDialog.getSaveFileName(
                self, "保存结果", str(Path(self.files[0]).with_name(f"{title}.xlsx")),
                "Excel 文件 (*.xlsx)", options=QFileDialog.DontConfirmOverwrite,
            )
            if not destination:
                return
            if not Path(destination).suffix:
                destination += ".xlsx"
            destination = validate_destination(destination, self.files)
            overwrite = False
            if destination.exists():
                answer = QMessageBox.question(
                    self, "确认覆盖", f"文件已存在，是否覆盖？\n{destination}",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
                )
                if answer != QMessageBox.Yes:
                    return
                overwrite = True
            saved = save_result(result, destination, self.files, overwrite=overwrite)
        except Exception as error:
            QMessageBox.critical(self, "操作失败", str(error))
            return
        QMessageBox.information(self, "导出完成", f"已导出 {len(result)} 行。\n{saved}")


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
