import re
from datetime import datetime

def extract_dates(text):
    return re.findall(r'\b(\d{2})\.(\d{2})\.(\d{4})\b', text)

def convert_to_date(date_tuple):
    return datetime.strptime(f"{date_tuple[0]}.{date_tuple[1]}.{date_tuple[2]}", "%d.%m.%Y")

def remove_lines_before_newest_date(text):
    dates = extract_dates(text)
    date_objects = [convert_to_date(date) for date in dates]
    newest_date = max(date_objects)
    newest_date_str = newest_date.strftime("%d.%m.%Y")
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if newest_date_str in line:
            return "\n".join(lines[i + 1:])
    return text

def move_timestamps_to_new_line(text):
    return re.sub(r'(\d{2}:\d{2}-\d{2}:\d{2})', r'\n\1', text)

def step2_main(input_file='schedule.txt', output_file='schedule2.txt'):
    with open(input_file, 'r', encoding='utf-8') as file:
        content = file.read()

    modified_content = remove_lines_before_newest_date(content)
    final_content = move_timestamps_to_new_line(modified_content)

    with open(output_file, 'w', encoding='utf-8') as file:
        file.write(final_content)
