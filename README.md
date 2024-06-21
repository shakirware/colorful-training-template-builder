# Colorful Training Template Builder

This project is a Python-based tool for generating visually appealing workout plans in Excel. It utilizes the openpyxl library to create a structured, color-coded template for organizing weekly workout routines.

![Screenshot](https://i.imgur.com/ozbElcw.png)

## Features
- **Create Consecutive Boxes**: Creates multiple weekly workout plans in a single sheet with proper spacing and color coding.
- **Label Days**: Labels days of the week and merges cells for day names.
- **Label Weeks**: Labels weeks with the start and end date.
- **Label Sets**: Labels sets and headers within the workout plan.

## Dynamic Sizing
The spreadsheet dynamically sizes itself based on the parameters you pass in by calculating the necessary dimensions and positions for each element.

## Customizability
- **Number of Weeks**: Specify the number of weeks for the workout plan.
- **Number of Sets**: Customize the number of sets per exercise.
- **Number of Exercises**: Define the number of exercises per day.
- **Start Date**: Set the start date for the workout plan.
- **Color Schemes**: Customize the color schemes for various elements:
  - `fill_color`: Color for the entire box.
  - `week_fill_color`: Color for the week label.
  - `set_fill_color`: Color for the set labels.
  - `day_fill_color`: Color for the day labels.
  - `exercise_fill_color`: Color for the exercise cells.
  - `grid_fill_color`: Color for the grid cells.
- **Padding and Spacing**:
  - `space_between`: Space between each week box.
  - `week_padding_top`: Padding at the top of the week label.
  - `week_padding_side`: Padding at the sides of the week label.
  - `set_padding_top`: Padding at the top of the set labels.
  - `set_padding_between`: Padding between each set.
  - `set_padding_side`: Padding at the sides of the sets.
  - `day_padding_top`: Padding at the top of each day section.
  - `day_padding_between`: Padding between each day.
  - `day_padding_side`: Padding at the sides of each day section.
- **Box Dimensions**:
  - `box_height`: Height of each week box.
  - `box_width`: Width of each week box.
  - `day_height`: Height of each day row.
  - `set_width`: Width of each set column.
- **Headers**: Customize the headers for the sets.

## Example Usage

```python
from datetime import datetime
from workout_plan_generator import ColourfulTemplateGenerator

generator = ColourfulTemplateGenerator('workout_plan_new.xlsx')

start_date = datetime.strptime('2024-06-24', '%Y-%m-%d')

generator.create_consecutive_boxes(
    start_row=2,
    start_col=2,
    box_height=42,
    box_width=21, 
    num_boxes=6,
    space_between=3,
    week_padding_top=1,
    week_padding_side=1,
    set_padding_top=1,
    set_padding_between=1,
    set_padding_side=1,
    num_sets=3,
    day_padding_top=1,
    day_padding_between=1,
    day_padding_side=1,
    day_height=4, 
    fill_color="2B7EA1",  
    week_fill_color="3397C1",  
    set_fill_color="4EA9D0",  
    day_fill_color="6EB8D8", 
    exercise_fill_color="8EC8E1", 
    grid_fill_color="AED8EA",
    set_width=4,
    headers=["Reps", "Sets", "%1RM", "Notes"],
    start_date=start_date
)

generator.save_workbook()
