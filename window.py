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
    QMessageBox
)
from PyQt5.QtCore import Qt


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
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
        self.table.horizontalHeader().setStyleSheet(
            """
            QHeaderView::section {
                border-bottom: 1px solid gray; /* 系统颜色的灰色边框 */
                background-color: #F0F0F0; /* 可选：表头背景颜色 */
            }
            """
        )
        self.table.setStyleSheet(
            """
            QTableWidget::item:selected {
                background: transparent; /* 移除选中背景色 */
                color: black; /* 确保文本颜色为黑色 */
            }
            QTableWidget::item:focus {
                border: 1px solid lightgray; /* 增加单元格的焦点边框 */
            }
            QTableWidget {
                gridline-color: lightgray; /* 设置表格线颜色 */
            }
            """
        )

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
        right_layout.addWidget(btn_clear)
        right_layout.addWidget(btn_generate)
        right_layout.addStretch()  # 添加弹性布局让按钮靠上

        main_layout.addLayout(right_layout)

        # 绑定右侧按钮事件
        btn_clear.clicked.connect(self.clear_table)
        btn_generate.clicked.connect(self.generate_output)

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
        combo_box.addItem("")  # 默认空值
        combo_box.addItems(["True", "False"])
        combo_box.setCurrentIndex(0)  # 设置默认选中空值
        # 设置下拉框样式
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

    def generate_output(self):
        """生成表格内容"""
        row_count = self.table.rowCount()
        col_count = self.table.columnCount()
        data = []

        for row in range(row_count):
            row_data = []
            for col in range(col_count):
                if col == 2:  # 下拉框
                    widget = self.table.cellWidget(row, col)
                    if widget is not None and isinstance(widget, QComboBox):
                        row_data.append(widget.currentText())
                else:  # 普通单元格
                    item = self.table.item(row, col)
                    row_data.append(item.text() if item else "")
            data.append(row_data)

        QMessageBox.information(self, "生成内容", f"表格内容:\n{data}")

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
        start_row = self.table.currentRow()  # 从当前选中的行开始
        start_col = self.table.currentColumn()  # 从当前选中的列开始

        for row_index, row in enumerate(rows):
            if not row.strip():
                continue  # 跳过空行
            cols = row.split("\t")  # 假设以制表符分隔
            for col_index, value in enumerate(cols):
                if col_index == 2:
                    value: str
                    value = value.capitalize()
                r = start_row + row_index
                c = start_col + col_index
                # 确保表格有足够的行和列
                if r >= self.table.rowCount():
                    self.table.insertRow(self.table.rowCount())
                if c >= self.table.columnCount():
                    continue  # 忽略超出列范围的内容
                # 填充表格单元格
                item = QTableWidgetItem(value.strip())
                self.table.setItem(r, c, item)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())
