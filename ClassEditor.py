import os
import tkinter as tk
from test1 import main as step1_main
from test2 import step2_main
from test3 import remove_teams_and_plus_lines
from test4 import add_newline_after_uhr
from test5 import remove_word_from_file
from test6 import remove_lines_with_keywords
from test7 import clean_schedule
from test8 import merge_LEK_with_next_line
from test9 import merge_teacher_names
from test10 import add_timestamp_to_schedule
from test11Wdays import extract_schedules
from test12 import process_all_files_in_days_folder
from test13 import merge_class_durations
from test14 import remove_mittagspause_lines
from test15 import calculate_day_hours_in_folder
from test16_fix_praxis import remove_zero_duration_lines
from test17 import create_schedule_document
from testGUI_New import ThemeEditorApp

def process_schedule_steps():
    steps = [
        ('schedule2.txt', 'schedule3.txt', remove_teams_and_plus_lines),
        ('schedule3.txt', 'schedule4.txt', add_newline_after_uhr),
        ('schedule4.txt', 'schedule5.txt', lambda i, o: remove_word_from_file(i, o, 'Uhr')),
        ('schedule5.txt', 'schedule6.txt', lambda i, o: remove_lines_with_keywords(i, o, ["Theorieunterricht", "Konsultation"])),
        ('schedule6.txt', 'schedule7.txt', clean_schedule),
        ('schedule7.txt', 'schedule8.txt', merge_LEK_with_next_line),
        ('schedule8.txt', 'schedule9.txt', merge_teacher_names),
        ('schedule9.txt', 'schedule10.txt', add_timestamp_to_schedule),
    ]

    for input_file, output_file, func in steps:
        func(input_file, output_file)
        os.remove(input_file)


def finalize_schedule():
    extract_schedules('schedule10.txt')
    os.remove('schedule10.txt')
    process_all_files_in_days_folder()
    merge_class_durations('days')
    remove_mittagspause_lines('days')
    calculate_day_hours_in_folder('days')
    remove_zero_duration_lines('days')
    create_schedule_document(input_file='schedule.txt', folder_path='days', output_file='Weekly_Class_Schedules.docx')

if __name__ == "__main__":
    step1_main()
    step2_main()
    process_schedule_steps()
    finalize_schedule()

    root = tk.Tk()
    app = ThemeEditorApp(root)
    root.mainloop()
