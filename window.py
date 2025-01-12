import sys
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QTableWidget,
    QTableWidgetItem,
    QPushButton,
    QComboBox,
    QHeaderView,
    QMessageBox,
    QProgressBar
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from visio_edit import VisioEdit


class WorkerThread(QThread):
    update_progress = pyqtSignal(int)  # 用于更新进度条的信号
    task_finished = pyqtSignal()  # 任务完成信号

    def __init__(self, table, parent=None, visible=False):
        super().__init__(parent)
        self.table = table
        self.visible = visible

    def run(self):
        """后台任务：处理表格数据并生成 Visio 图形"""
        row_count = self.table.rowCount()
        col_count = self.table.columnCount()

        with VisioEdit("output_data/new_file.vsd", visible=self.visible) as editor:
            for row in range(row_count):
                row_data = []

                for col in range(col_count):
                    if col == 2:  # 下拉框
                        widget = self.table.cellWidget(row, col)
                        if widget is not None and isinstance(widget, QComboBox):
                            row_data.append(True if widget.currentText().capitalize() == "True" else False)
                    else:  # 普通单元格
                        item = self.table.item(row, col)
                        row_data.append(item.text() if item else "")

                # 跳过无效数据
                if len(row_data) != 4:
                    continue
                try:
                    row_data[0] = int(row_data[0])
                    row_data[-1] = float(row_data[-1])
                except ValueError:
                    continue

                editor.add_action(*row_data)

                # 每处理一行更新进度条
                self.update_progress.emit(int((row + 1) / row_count * 100))

            # 完成绘图
            editor.paint()

        # 任务完成
        self.task_finished.emit()


class ProgressWindow(QWidget):
    """悬浮进度条窗口"""
    def __init__(self):
        super().__init__()
        self.progress_bar = None
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.progress_bar.setStyleSheet(
            """
            QProgressBar {
                text-align: center;
                border: 1px solid #bbb;
                background: #eee;
            }
            QProgressBar::chunk {
                background-color: #4CAF50; /* 进度条颜色 */
                width: 20px;
            }
            """
        )
        layout.addWidget(self.progress_bar)
        self.setLayout(layout)
        self.setWindowTitle("生成进度")
        self.resize(300, 100)


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.worker_thread = None
        self.progress_window = None
        self.table = None
        self.initUI()

    def initUI(self):
        # 设置主窗口布局
        main_layout = QHBoxLayout()

        # 左侧表格
        self.table = QTableWidget(0, 4)  # 初始0行，4列
        self.table.setHorizontalHeaderLabels(["指令时刻/s", "动作", "打开或关闭", "长度"])
        self.table.horizontalHeader().setSectionsClickable(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.verticalHeader().setVisible(False)

        # 添加功能按钮
        left_layout = QVBoxLayout()
        left_layout.addWidget(self.table)
        btn_add_row = QPushButton("添加行")
        btn_delete_row = QPushButton("删除行")
        left_layout.addWidget(btn_add_row)
        left_layout.addWidget(btn_delete_row)
        main_layout.addLayout(left_layout)

        # 绑定按钮事件
        btn_add_row.clicked.connect(self.add_row)
        btn_delete_row.clicked.connect(self.delete_row)

        # 右侧操作按钮
        right_layout = QVBoxLayout()
        btn_clear = QPushButton("清空")
        btn_generate = QPushButton("生成")
        btn_edit = QPushButton("打开visio编辑")
        right_layout.addWidget(btn_clear)
        right_layout.addWidget(btn_generate)
        # right_layout.addWidget(btn_edit)
        right_layout.addStretch()  # 添加弹性布局让按钮靠上
        main_layout.addLayout(right_layout)

        # 绑定右侧按钮事件
        btn_clear.clicked.connect(self.clear_table)
        btn_generate.clicked.connect(self.generate_output)
        btn_edit.clicked.connect(self.edit)

        # 设置主窗口布局
        self.setLayout(main_layout)
        self.setWindowTitle("指令操作 GUI")
        self.resize(800, 400)

    def add_row(self):
        """添加新行"""
        row_count = self.table.rowCount()
        self.table.insertRow(row_count)

        # 设置下拉框
        combo_box = QComboBox()
        combo_box.addItems(["", "True", "False"])
        combo_box.setStyleSheet(
            """
            QComboBox {
                border: none; /* 无边框 */
                background-color: white; /* 背景颜色 */
                padding-left: 4px; /* 左侧内边距 */
            }
            """
        )
        self.table.setCellWidget(row_count, 2, combo_box)  # 第三列是下拉框

        # 可编辑单元格
        for col in [0, 1, 3]:  # 指令时刻、动作、长度
            item = QTableWidgetItem()
            item.setFlags(item.flags() | Qt.ItemIsEditable)  # 设置可编辑
            self.table.setItem(row_count, col, item)

    def delete_row(self):
        """删除选中行"""
        current_row = self.table.currentRow()
        if current_row >= 0:
            self.table.removeRow(current_row)
        else:
            QMessageBox.warning(self, "删除行", "请先选中一行再删除！")

    def clear_table(self):
        """清空表格"""
        self.table.setRowCount(0)

    def edit(self):
        self.generate_output(visible=True)

    def generate_output(self, visible=False):
        """生成表格内容"""
        self.progress_window = ProgressWindow()  # 显示进度窗口
        self.progress_window.show()

        self.worker_thread = WorkerThread(self.table, visible=visible)
        self.worker_thread.update_progress.connect(self.progress_window.progress_bar.setValue)
        self.worker_thread.task_finished.connect(self.task_complete)
        self.worker_thread.task_finished.connect(self.progress_window.close)
        self.worker_thread.start()

    def task_complete(self):
        """任务完成"""
        QMessageBox.information(self, "任务完成", "生成操作已完成！")

    def keyPressEvent(self, event):
        """捕获键盘事件"""
        if event.modifiers() == Qt.ControlModifier and event.key() == Qt.Key_V:
            self.paste_clipboard_content()

    def paste_clipboard_content(self):
        """从剪贴板粘贴内容到表格"""
        clipboard = QApplication.clipboard()
        data = clipboard.text()
        if not data:
            QMessageBox.warning(self, "粘贴错误", "剪贴板没有内容！")
            return

        rows = data.split("\n")
        start_row = self.table.currentRow() if self.table.currentRow() >= 0 else 0  # 从当前选中的行开始
        start_col = 0

        for row_index, row in enumerate(rows):
            if not row.strip():
                continue  # 跳过空行
            cols = row.split("\t")  # 假设以制表符分隔
            for col_index, value in enumerate(cols):
                r = start_row + row_index
                c = start_col + col_index
                value = value.strip()

                # 确保表格有足够的行
                if r >= self.table.rowCount():
                    self.table.insertRow(self.table.rowCount())
                if c >= self.table.columnCount():
                    continue  # 忽略超出列范围的内容

                if c == 2:  # 第三列为下拉框
                    combo_box = QComboBox()
                    combo_box.setStyleSheet(
                        """
                        QComboBox {
                            border: none; /* 无边框 */
                            background-color: white; /* 背景颜色 */
                            padding-left: 4px; /* 左侧内边距 */
                        }
                        """
                    )
                    combo_box.addItems(["", "True", "False"])
                    combo_box.setCurrentText(value.capitalize())
                    self.table.setCellWidget(r, c, combo_box)
                else:
                    item = QTableWidgetItem(value)
                    self.table.setItem(r, c, item)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())
