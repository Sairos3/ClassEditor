import os

def extract_schedules(input_file):
    weekdays = ["1_Montag", "2_Dienstag", "3_Mittwoch", "4_Donnerstag", "5_Freitag"]
    
    if not os.path.exists("days"):
        os.makedirs("days")
    
    with open(input_file, 'r') as f:
        lines = f.readlines()
    
    schedule_chunks = []
    current_schedule = []

    for line in lines:
        current_schedule.append(line)
        if '-16:00' in line:
            schedule_chunks.append(current_schedule)
            current_schedule = []
    
    for i, schedule in enumerate(schedule_chunks):
        weekday_name = weekdays[i % len(weekdays)]
        with open(f"days/{weekday_name}_schedule.txt", 'w') as f:
            f.writelines(schedule)
