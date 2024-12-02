import os

def remove_zero_duration_lines(directory_path):
    for filename in os.listdir(directory_path):
        if filename.endswith(".txt"):
            file_path = os.path.join(directory_path, filename)
            
            with open(file_path, "r") as file:
                lines = file.readlines()
            
            filtered_lines = [line for line in lines if "Duration: 0" not in line]
            
            with open(file_path, "w") as file:
                file.writelines(filtered_lines)
