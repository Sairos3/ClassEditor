import re

def add_timestamp_to_schedule(input_file, output_file):
    timestamp_pattern = re.compile(r'^\d{2}:\d{2}-\d{2}:\d{2}$')
    
    with open(input_file, 'r') as file:
        lines = file.readlines()
    
    updated_lines = []
    
    for i, line in enumerate(lines):
        stripped_line = line.strip()
        updated_lines.append(stripped_line)
        
        if (not timestamp_pattern.match(stripped_line) and
                (i + 1 >= len(lines) or not timestamp_pattern.match(lines[i + 1].strip()))):
            updated_lines.append("08:00-16:00")
    
    with open(output_file, 'w') as file:
        file.write("\n".join(updated_lines))
