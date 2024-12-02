import os

def time_to_minutes(time_str):
    return int(time_str[:2]) * 60 + int(time_str[3:5])

def minutes_to_hm(minutes):
    return f"{minutes // 60}h {minutes % 60}min"

def process_schedule(input_file, output_file):
    try:
        with open(input_file, 'r') as file:
            lines = file.read().strip().splitlines()

        schedule = []
        for i in range(0, len(lines) - 1, 2):
            class_name = lines[i].strip()
            time_range = lines[i + 1].strip()

            start_time, end_time = time_range.split('-')
            duration_minutes = time_to_minutes(end_time) - time_to_minutes(start_time)
            duration_str = minutes_to_hm(duration_minutes)

            schedule.append(f"Class: {class_name} | Duration: {duration_str}")

        with open(output_file, 'w') as file:
            file.write('\n'.join(schedule))

    except FileNotFoundError:
        print(f"Error: The file '{input_file}' was not found.")
    except Exception as e:
        print(f"An error occurred: {e}")

def process_all_files_in_days_folder():
    input_folder = 'days'
    output_folder = 'days'

    for filename in os.listdir(input_folder):
        if filename.endswith('.txt'):
            input_file = os.path.join(input_folder, filename)
            output_file = os.path.join(output_folder, filename)

            process_schedule(input_file, output_file)
