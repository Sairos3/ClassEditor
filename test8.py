import re

def merge_LEK_with_next_line(input_file, output_file):
    try:
        with open(input_file, 'r', encoding='utf-8') as file:
            lines = file.readlines()

        merged_lines = []
        skip_next = False

        for i in range(len(lines) - 1):
            if skip_next:
                skip_next = False
                continue

            current_line = lines[i].strip()
            next_line = lines[i + 1].strip()

            if current_line.startswith("LEK"):
                merged_line = current_line + " " + next_line
                merged_lines.append(merged_line)
                skip_next = True
            else:
                merged_lines.append(current_line)

        if not skip_next:
            merged_lines.append(lines[-1].strip())

        with open(output_file, 'w', encoding='utf-8') as file:
            file.write("\n".join(merged_lines))

    except FileNotFoundError:
        print(f"Error: The file {input_file} was not found.")
    except Exception as e:
        print(f"An error occurred: {e}")
