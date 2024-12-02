import os
from datetime import timedelta

def parse_duration(duration_str):
    duration_str = duration_str.replace('Duration:', '').strip()
    hours, minutes = 0, 0
    if 'h' in duration_str:
        hours = int(duration_str.split('h')[0].strip())
    if 'min' in duration_str:
        minutes = int(duration_str.split('h')[1].split('min')[0].strip()) if 'h' in duration_str else int(duration_str.split('min')[0].strip())
    return timedelta(hours=hours, minutes=minutes)

def format_duration(td):
    hours = td.seconds // 3600
    minutes = (td.seconds % 3600) // 60
    return f"{hours}h {minutes}min"

def merge_class_durations(directory_path):
    for filename in os.listdir(directory_path):
        file_path = os.path.join(directory_path, filename)

        if filename.endswith('.txt'):
            with open(file_path, 'r') as f:
                lines = f.readlines()

            class_durations = {}

            for line in lines:
                parts = line.split('|')
                if len(parts) < 2:
                    continue  # Skip malformed lines
                class_name, duration_str = parts[0].strip(), parts[1].strip()

                try:
                    duration = parse_duration(duration_str)
                except ValueError:
                    print(f"Skipping invalid duration format in line: {line}")
                    continue

                class_durations[class_name] = class_durations.get(class_name, timedelta()) + duration

            output_path = os.path.join(directory_path, filename)
            with open(output_path, 'w') as f:
                for class_name, total_duration in class_durations.items():
                    f.write(f"{class_name} | Duration: {format_duration(total_duration)}\n")
