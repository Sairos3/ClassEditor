import os
import math

def parse_duration(duration_str):
    hours, minutes = map(int, duration_str.replace('h', '').replace('min', '').split())
    return hours + math.ceil(minutes / 60)

def calculate_day_hours_in_folder(folder_path, max_hours=8):
    for filename in os.listdir(folder_path):
        if filename.endswith('.txt'):
            file_path = os.path.join(folder_path, filename)
            calculate_day_hours(file_path, max_hours)

def calculate_day_hours(file_path, max_hours=8):
    with open(file_path, 'r') as file:
        lines = file.readlines()
    
    schedule = []
    praxis_index = None
    total_hours = 0
    
    for i, line in enumerate(lines):
        if "Class:" in line and "Duration:" in line:
            class_info, duration_str = line.split("| Duration:")
            duration = parse_duration(duration_str.strip())
            total_hours += duration
            
            if "Praxisunterricht" in class_info:
                praxis_index = i
            
            schedule.append((class_info.strip(), duration))
    
    if total_hours > max_hours and praxis_index is not None:
        overdone_hours = total_hours - max_hours
        praxis_class, praxis_duration = schedule[praxis_index]
        new_praxis_duration = max(0, praxis_duration - overdone_hours)
        schedule[praxis_index] = (praxis_class, new_praxis_duration)
        total_hours = max_hours
    
    updated_lines = [
        f"{class_info} | Duration: {duration}\n"
        for class_info, duration in schedule
    ]
    
    with open(file_path, 'w') as file:
        file.writelines(updated_lines)
