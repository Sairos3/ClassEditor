def remove_teams_and_plus_lines(input_file, output_file):
    try:
        with open(input_file, 'r', encoding='utf-8') as file:
            filtered_lines = [
                line for line in file 
                if not line.startswith("Teams") and "(+" not in line
            ]
        with open(output_file, 'w', encoding='utf-8') as file:
            file.writelines(filtered_lines)
    except FileNotFoundError:
        print(f"Error: File '{input_file}' not found.")
    except Exception as e:
        print(f"An error occurred: {e}")
