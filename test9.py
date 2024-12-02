import re

def merge_teacher_names(file_input, file_output):
    with open(file_input, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    merged_lines = []
    i = 0
    while i < len(lines):
        current_line = lines[i].strip()
        next_line = lines[i + 1].strip() if i + 1 < len(lines) else ""

        if '/' in current_line and not re.match(r'\d{2}:\d{2}-\d{2}:\d{2}', next_line):
            merged_lines.append(current_line + " " + next_line)
            i += 2
        else:
            merged_lines.append(current_line)
            i += 1

    with open(file_output, 'w', encoding='utf-8') as file:
        file.write("\n".join(merged_lines))
