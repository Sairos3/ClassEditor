import re

def clean_schedule(input_file, output_file):
    try:
        with open(input_file, 'r', encoding='utf-8') as file:
            schedule_text = file.read()

        cleaned_text = re.sub(r'(Unterrichtsfrei|Sonderveranstaltung|Praxisunterricht|Mittagspause)\d*\s*min?', r'\1', schedule_text)

        with open(output_file, 'w', encoding='utf-8') as file:
            file.write(cleaned_text)

    except FileNotFoundError:
        print(f"Error: The file '{input_file}' was not found.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
