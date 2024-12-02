def remove_word_from_file(input_file, output_file, word_to_remove):
    try:
        with open(input_file, 'r', encoding='utf-8') as file:
            schedule_text = file.read()

        with open(output_file, 'w', encoding='utf-8') as file:
            file.write(schedule_text.replace(word_to_remove, ''))

    except FileNotFoundError:
        print(f"Error: The file {input_file} was not found.")
    except Exception as e:
        print(f"An error occurred: {e}")
