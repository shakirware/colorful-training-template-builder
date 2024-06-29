import openpyxl
from box_creator import BoxCreator
from label_creator import LabelCreator
from data_populator import DataPopulator

class WorkoutTemplateGenerator:
    def __init__(self, workbook_name):
        self.workbook = openpyxl.Workbook()
        self.sheet = self.workbook.active
        self.workbook_name = workbook_name
        self.box_creator = BoxCreator(self.sheet)
        self.label_creator = LabelCreator(self.sheet)
        self.data_populator = DataPopulator(self.sheet)

    def create_consecutive_boxes(self, start_row, start_col, num_boxes, space_between, fill_color, week_fill_color, set_fill_color, day_fill_color, exercise_fill_color, grid_fill_color, set_width, headers, workout_data):
        for i in range(num_boxes):
            box_height, box_width = self.calculate_box_dimensions(workout_data[i]['days'], set_width)
            current_start_col = start_col + i * (box_width + space_between)
            self.box_creator.create_box(start_row, current_start_col, start_row + box_height - 1, current_start_col + box_width - 1, fill_color)
            self.label_creator.label_weeks(start_row, current_start_col, box_width, i + 1, week_fill_color)
            set_start_row = start_row + 3
            self.label_creator.label_sets(set_start_row, current_start_col + 1, 3, set_fill_color, set_width, headers)
            day_start_row = set_start_row + 4
            self.label_creator.label_days(day_start_row, current_start_col, 6, day_fill_color, exercise_fill_color, 3, set_width, grid_fill_color)
            self.data_populator.populate_workout_data(day_start_row, current_start_col, workout_data[i]['days'], 3, set_width)

    def calculate_box_dimensions(self, days_data, set_width):
        min_height = 2 + 1 + 2 + 1 + 7 * (6 + 1) + 1
        min_width = 1 + 3 + 3 * (set_width + 1) - 1 + 2 + 1
        return min_height, min_width

    def save_workbook(self):
        self.workbook.save(self.workbook_name)
        print(f"Workbook saved as {self.workbook_name}")
