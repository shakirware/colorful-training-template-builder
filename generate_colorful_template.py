import openpyxl
from openpyxl.styles import PatternFill

class ColourfulTemplateGenerator:
    def __init__(self, workbook_name):
        self.workbook = openpyxl.Workbook()
        self.sheet = self.workbook.active
        self.workbook_name = workbook_name

    def create_box(self, start_row, start_col, end_row, end_col, fill_color):
        fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                cell = self.sheet.cell(row=row, column=col)
                cell.fill = fill

    def label_days(self, start_row, start_col, padding=1):
        days = ["Day 1", "Day 2", "Day 3", "Day 4", "Day 5", "Day 6", "Day 7"]
        for i, day in enumerate(days):
            self.sheet.cell(row=start_row + i * (padding + 1), column=start_col + padding, value=day)

    def create_consecutive_boxes(self, start_row, start_col, box_height, box_width, num_boxes, space_between, fill_color):
        for i in range(num_boxes):
            current_start_col = start_col + i * (box_width + space_between)
            self.create_box(start_row, current_start_col, start_row + box_height - 1, current_start_col + box_width - 1, fill_color)
            self.label_days(start_row, current_start_col)

    def save_workbook(self):
        self.workbook.save(self.workbook_name)
        print(f"Workbook saved as {self.workbook_name}")

generator = ColourfulTemplateGenerator('workout_plan_new.xlsx')

generator.create_consecutive_boxes(start_row=4, start_col=4, box_height=20, box_width=10, num_boxes=4, space_between=3, fill_color="AF125A")

generator.save_workbook()
