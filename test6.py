def remove_lines_with_keywords(input_file, output_file, keywords_to_remove):
    try:
        with open(input_file, 'r', encoding='utf-8') as file:
            lines = file.readlines()

        with open(output_file, 'w', encoding='utf-8') as file:
            for line in lines:
                if not any(keyword in line for keyword in keywords_to_remove):
                    file.write(line)

    except FileNotFoundError:
        print(f"Error: The file {input_file} was not found.")
    except Exception as e:
        print(f"An error occurred: {e}")
