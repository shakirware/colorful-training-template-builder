from workout_template_generator import WorkoutTemplateGenerator
from workout_data_loader import load_workout_data

def main():
    workout_data = load_workout_data('workout_data.yaml')
    generator = WorkoutTemplateGenerator('workout_plan_new.xlsx')
    generator.create_consecutive_boxes(
        start_row=2,
        start_col=2,
        num_boxes=6,
        space_between=3,
        set_width=4,
        headers=["Reps", "Weights", "%1RM", "Notes"],
        workout_data=workout_data,
        fill_color="4C40F2",
        week_fill_color="5E53F3",
        set_fill_color="7066F4",
        day_fill_color="8279F6",
        exercise_fill_color="948DF7",
        grid_fill_color="A6A0F8"
    )
    generator.save_workbook()

if __name__ == "__main__":
    main()
