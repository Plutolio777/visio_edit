import logging
import os.path

import pythoncom
import win32com.client
from collections import namedtuple

ActionPoint = namedtuple("ActionPoint", ["time", "action", "is_open", "length"])


class VisioEdit:

    def __init__(self, save_path, visible=False, is_save=True):
        pythoncom.CoInitialize()
        self.visible = visible
        self.text_shape = None
        self.line_width = None
        self.timeline = None
        self.visio = win32com.client.Dispatch("Visio.Application")
        self.visio.Visible = visible
        self.doc = self.visio.Documents.Add("")  # 新建一个文档
        self.page = self.visio.ActivePage  # 获取当前页
        self.page_width = 20
        self.page_height = 20
        self.x_scale = 5
        self.action_line_height = 1
        self.page.PageSheet.CellsU("PrintPageOrientation").ResultIU = 1
        self.actions = {}
        self.column_text_width = 0.35
        self.save_path = save_path
        self.is_save = is_save

    def __enter__(self):
        return self  # 可返回任意对象，赋值给 `as` 子句的变量

    def __exit__(self, exc_type, exc_value, traceback):
        # timeline = self.page.DrawLine(0.5, 10, 19.5, 10)  # 从(0, 5) 到 (10, 5)的直线
        # 加粗时间轴
        try:
            output_dir = os.path.abspath("output_data")  # 转换为绝对路径
            if not os.path.exists(output_dir):
                os.mkdir(output_dir)
            self.save_path = os.path.join(output_dir, "new_file.vsd")
            if os.path.exists(self.save_path):
                os.remove(self.save_path)
            if self.is_save:
                self.doc.SaveAs(self.save_path)
        except Exception as e:
            logging.error(e, exc_info=True)
        finally:
            if not self.visible:
                self.doc.Close()
                self.visio.Quit()

    def add_action(self, time, action, is_open, length):
        if time not in self.actions:
            self.actions[time] = []
        actions = self.actions.get(time, [])
        actions.append(ActionPoint(time, action, is_open, length))

    def reset_page_size(self):
        self.line_width = max(max([max([j.length for j in i]) for i in self.actions.values()]), 1)
        print(self.line_width)
        self.line_width += 0.25
        self.line_width = self.x_scale * self.line_width

        self.page_width = self.line_width + 2
        self.page_height = self.page_width

        self.page.PageSheet.CellsU("PageWidth").ResultIU = self.page_width  # 宽度
        self.page.PageSheet.CellsU("PageHeight").ResultIU = self.page_height

    def paint_time_line(self):
        point = [0.5, self.page_height / 2, self.line_width, self.page_height / 2]
        print(point)
        self.timeline = self.page.DrawLine(*point)
        self.timeline.Cells("LineWeight").FormulaU = "1.5 pt"  # 加粗线条，单位为 pt
        # 添加结束箭头
        self.timeline.Cells("EndArrow").FormulaU = "2"  # 普通箭头

        text_x = self.line_width + 0.5
        text_y = self.page_height / 2
        self.text_shape = self.page.DrawRectangle((text_x - 0.3) - 0.25, text_y - 0.3, (text_x + 0.3) - 0.25,
                                                  text_y + 0.3)  # 添加一个透明的文本框
        self.text_shape.Text = "t(s)"
        self.text_shape.CellsU("Char.Size").FormulaU = "14 pt"
        self.text_shape.CellsU("Char.Style").FormulaU = "2"
        self.text_shape.Cells("FillForegnd").FormulaU = "RGB(255,255,255)"  # 白色或透明
        self.text_shape.Cells("FillPattern").FormulaU = "0"  # 填充模式为“无填充”
        self.text_shape.Cells("LineColor").FormulaU = "RGB(255,255,255)"  # 边框颜色设置为透明
        self.text_shape.Cells("LinePattern").FormulaU = "0"  # 边框模式为“无边框”

    def paint_default(self):
        pass

    @staticmethod
    def correction_characters(text):
        new_text = ""
        for i, char in enumerate(text):
            # 判断是否为中文字符（Unicode 范围）
            if '\u4e00' <= char <= '\u9fff' and i != len(text) - 1:  # 中文字符范围
                new_text = new_text + char + '\n'
            else:
                new_text += char
        return new_text

    @staticmethod
    def correction_characters_size(text_table):
        text = text_table.Text
        characters = text_table.Characters
        start_index = 0  # 起始索引
        for i, char in enumerate(text):
            # 判断是否为中文字符（Unicode 范围）
            if '\u4e00' <= char <= '\u9fff':  # 中文字符范围
                font_size = 9
            else:  # 其他字符（如标点）
                font_size = 7

            # 设置字符范围
            characters.Begin = start_index
            characters.End = i + 1
            characters.CharProps(7, font_size)  # 切换到字符属性
            # characters.Cells("Char.Size").FormulaU = font_size  # 设置字体大小

            # 更新起始索引
            start_index = i + 1

    def paint_actions(self):
        for row_index, (time, action_group) in enumerate(self.actions.items()):
            min_length = min([i.length for i in action_group])
            is_open = any([i.is_open for i in action_group])
            is_close = any([not i.is_open for i in action_group])
            x_start = (min_length * self.x_scale) + 0.5

            if is_open:
                line_point = [x_start, (self.page_height / 2) + 0.5, x_start, (self.page_height / 2)]
                print(action_group, line_point)
                line = self.page.DrawLine(*line_point)
                line.Cells("EndArrow").FormulaU = "2"
                line.Cells("LineWeight").FormulaU = "1.3 pt"
                open_action = []
                for action in action_group:
                    action: ActionPoint
                    if action.is_open:
                        open_action.append(action)
                table_width = self.column_text_width * len(open_action)
                should_draw_underline = len(open_action) > 1
                for index, action in enumerate(open_action):
                    action_text = VisioEdit.correction_characters(action.action)
                    point = [
                        line_point[0] - (table_width / 2) + self.column_text_width * index,
                        (self.page_height / 2) + 0.55,
                        line_point[0] - (table_width / 2) + self.column_text_width * (index + 1),
                        (self.page_height / 2) + 5
                    ]
                    text_table = self.page.DrawRectangle(*point)
                    text_table.Text = action_text
                    VisioEdit.correction_characters_size(text_table)

                    text_table.CellsU("VerticalAlign").FormulaU = "2"
                    text_table.Cells("FillForegnd").FormulaU = "RGB(255,255,255)"  # 白色或透明
                    text_table.Cells("FillPattern").FormulaU = "0"  # 填充模式为“无填充”
                    text_table.Cells("LineColor").FormulaU = "RGB(255,255,255)"  # 边框颜色设置为透明
                    text_table.Cells("LinePattern").FormulaU = "0"  # 边框模式为“无边框”
                    # text_table.CellsU("Char.Size").FormulaU = "9 pt"
                    if should_draw_underline:
                        underline_point = [
                            point[0],
                            line_point[1],
                            point[2],
                            line_point[1]
                        ]
                        text_table_underline = self.page.DrawLine(*underline_point)
                        text_table_underline.Cells("LineWeight").FormulaU = "1.3 pt"

            if is_close:
                line_point = [x_start, (self.page_height / 2) - 0.5, x_start, (self.page_height / 2)]
                line = self.page.DrawLine(*line_point)
                line.Cells("EndArrow").FormulaU = "2"
                line.Cells("LineWeight").FormulaU = "1.3 pt"

                close_action = []
                for action in action_group:
                    action: ActionPoint
                    if not action.is_open:
                        close_action.append(action)
                table_width = self.column_text_width * len(close_action)
                should_draw_underline = len(close_action) > 1
                for index, action in enumerate(close_action):
                    action_text = VisioEdit.correction_characters(action.action)
                    point = [
                        line_point[0] - (table_width / 2) + self.column_text_width * index,
                        (self.page_height / 2) - 0.55,
                        line_point[0] - (table_width / 2) + self.column_text_width * (index + 1),
                        (self.page_height / 2) - 5
                    ]
                    text_table = self.page.DrawRectangle(*point)
                    text_table.Text = action_text
                    VisioEdit.correction_characters_size(text_table)
                    text_table.CellsU("VerticalAlign").FormulaU = "0"
                    text_table.Cells("FillForegnd").FormulaU = "RGB(255,255,255)"  # 白色或透明
                    text_table.Cells("FillPattern").FormulaU = "0"  # 填充模式为“无填充”
                    text_table.Cells("LineColor").FormulaU = "RGB(255,255,255)"  # 边框颜色设置为透明
                    text_table.Cells("LinePattern").FormulaU = "0"  # 边框模式为“无边框”
                    # text_table.CellsU("Char.Size").FormulaU = "9 pt"
                    if should_draw_underline:
                        underline_point = [
                            point[0],
                            line_point[1],
                            point[2],
                            line_point[1]
                        ]
                        text_table_underline = self.page.DrawLine(*underline_point)
                        text_table_underline.Cells("LineWeight").FormulaU = "1.3 pt"

            # 绘制序号
            index_box_point = [x_start, (self.page_height / 2) + 0.03, x_start + 5, (self.page_height / 2) + 0.3]
            index_box = self.page.DrawRectangle(*index_box_point)
            index_box.Text = f'({row_index + 1})'
            index_box.Cells("Para.HorzAlign").FormulaU = "0"
            index_box.Cells("FillForegnd").FormulaU = "RGB(255,255,255)"  # 白色或透明
            index_box.Cells("FillPattern").FormulaU = "0"  # 填充模式为“无填充”
            index_box.Cells("LineColor").FormulaU = "RGB(255,255,255)"  # 边框颜色设置为透明
            index_box.Cells("LinePattern").FormulaU = "0"  # 边框模式为“无边框”
            index_box.CellsU("Char.Size").FormulaU = "13 pt"

            index_box_point = [x_start, (self.page_height / 2) - 0.03, x_start + 5, (self.page_height / 2) - 0.3]
            index_box = self.page.DrawRectangle(*index_box_point)
            index_box.Text = f'{time}s'
            index_box.Cells("Para.HorzAlign").FormulaU = "0"
            index_box.Cells("FillForegnd").FormulaU = "RGB(255,255,255)"  # 白色或透明
            index_box.Cells("FillPattern").FormulaU = "0"  # 填充模式为“无填充”
            index_box.Cells("LineColor").FormulaU = "RGB(255,255,255)"  # 边框颜色设置为透明
            index_box.Cells("LinePattern").FormulaU = "0"  # 边框模式为“无边框”
            index_box.CellsU("Char.Size").FormulaU = "10 pt"

    def paint(self):
        try:
            if not self.actions:
                self.paint_default()
            self.reset_page_size()
            self.paint_time_line()
            self.paint_actions()
        except Exception as e:
            logging.error(e, exc_info=True)


if __name__ == '__main__':
    with VisioEdit("output_data/new_file.vsd", visible=True) as editor:
        editor.add_action(0.1, "DQ11打开", True, 0)
        editor.add_action("t1", "这是一个动作2", True, 0.1)
        editor.add_action("t1", "这是一个动作3", True, 0.2)
        editor.add_action(3, "动作3", True, 0.3)
        editor.add_action(4, "动作4", False, 0.4)
        editor.add_action(5, "动作5", True, 0.5)
        editor.add_action(6, "打开阀门", True, 0.6)
        editor.add_action(6, "关闭电动气阀", False, 0.6)
        editor.add_action(8, "关闭电动气阀1", False, 0.8)
        editor.add_action(8, "关闭电动气阀2", False, 0.9)
        editor.paint()
