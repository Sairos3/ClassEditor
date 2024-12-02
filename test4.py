import re

def add_newline_after_uhr(input_file, output_file):
    try:
        with open(input_file, 'r', encoding='utf-8') as file:
            schedule_text = file.read()

        transformed_text = re.sub(r"Uhr([A-Za-z0-9 ]+)", r"Uhr\n\1", schedule_text)

        with open(output_file, 'w', encoding='utf-8') as file:
            file.write(transformed_text)

    except FileNotFoundError:
        print(f"Error: The file {input_file} was not found.")
    except Exception as e:
        print(f"An error occurred: {e}")
