import win32com.client
from collections import namedtuple

ActionPoint = namedtuple("ActionPoint", ["time", "action", "is_open", "length"])


class VisioEdit:

    def __init__(self, save_path, visible=True):
        self.visio = win32com.client.Dispatch("Visio.Application")
        self.visio.Visible = visible
        self.doc = self.visio.Documents.Add("")  # 新建一个文档
        self.page = self.visio.ActivePage  # 获取当前页
        self.width = 20
        self.heigt = 20
        self.page.PageSheet.CellsU("PrintPageOrientation").ResultIU = 1
        self.actions = {}

    def __enter__(self):
        return self  # 可返回任意对象，赋值给 `as` 子句的变量

    def __exit__(self, exc_type, exc_value, traceback):
        timeline = self.page.DrawLine(0.5, 10, 19.5, 10)  # 从(0, 5) 到 (10, 5)的直线
        # 加粗时间轴
        timeline.Cells("LineWeight").FormulaU = "4 pt"  # 加粗线条，单位为 pt
        # 添加结束箭头
        timeline.Cells("EndArrow").FormulaU = "1"  # 普通箭头

    def add_action(self, time, action, is_open, length):
        actions = self.actions.get(time, [])
        actions.append(ActionPoint(time, action, is_open, length))

    def reset_page_size(self):
        pass

    def paint_time_line(self):
        pass

    def paint_default(self):
        self.page.PageSheet.CellsU("PageWidth").ResultIU = 20  # 宽度
        self.page.PageSheet.CellsU("PageHeight").ResultIU = 20  # 高度

    def paint(self):
        if not self.actions:
            self.paint_default()
        self.reset_page_size()
        self.paint_time_line()


if __name__ == '__main__':
    with VisioEdit("") as editor:
        editor.add_action(0, "动作1", True, 0)
        editor.add_action(1, "这是一个动作2", True, 0.1)
        editor.add_action(2, "动作3", True, 0.2)
        editor.add_action(3, "动作4", False, 0.3)
        editor.add_action(4, "动作5", True, 0.4)
        editor.add_action(5, "打开阀门", True, 0.5)
        editor.add_action(5, "关闭电动气阀", False, 0.6)
        editor.paint()
