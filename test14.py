import os

def remove_mittagspause_lines(folder_path):
    for filename in os.listdir(folder_path):
        if filename.endswith('.txt'):
            file_path = os.path.join(folder_path, filename)

            with open(file_path, 'r') as file:
                lines = file.readlines()

            filtered_lines = [line for line in lines if 'Mittagspause' not in line]

            with open(file_path, 'w') as file:
                file.writelines(filtered_lines)
