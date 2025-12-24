
import csv
import random
from datetime import datetime, timedelta

def generate_fingerprints():
    # Read employees data
    employees = {}
    with open('employees.csv', 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            employees[row['Employee_ID']] = {'name': row['Employee_Name'], 'dept': row['Department']}

    # Required punch times
    required_times = [
        (8, 0), (12, 0), (15, 0), (20, 0), (23, 0), (8, 0) # last one is next day
    ]

    with open('fingerprints.csv', 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['Employee_ID', 'Employee_Name', 'Department', 'Fingerprint_Date', 'Fingerprint_Time'])

        with open('shift_schedule.csv', 'r', encoding='utf-8') as schedule_file:
            reader = csv.DictReader(schedule_file)
            for row in reader:
                emp_id = row['Employee_ID']
                shift_date = datetime.strptime(row['Shift_Date'], '%Y-%m-%d').date()
                emp_data = employees[emp_id]

                # Decide the type of day
                day_type = random.choices(
                    ['perfect', 'incomplete', 'late', 'absent'], 
                    weights=[0.70, 0.15, 0.10, 0.05], 
                    k=1
                )[0]

                if day_type == 'absent':
                    continue # No punches for this day

                punches_to_make = list(range(len(required_times)))
                if day_type == 'incomplete':
                    # Randomly remove 1 or 2 punches
                    missing_count = random.randint(1, 2)
                    for _ in range(missing_count):
                        if punches_to_make:
                            punches_to_make.pop(random.randrange(len(punches_to_make)))
                
                late_punch_index = -1
                if day_type == 'late':
                    late_punch_index = random.choice(punches_to_make)

                for i, (h, m) in enumerate(required_times):
                    if i not in punches_to_make:
                        continue

                    punch_day = shift_date
                    if i == 5: # Next day's 8:00
                        punch_day += timedelta(days=1)
                    
                    base_time = datetime.combine(punch_day, datetime.min.time()).replace(hour=h, minute=m)
                    
                    if i == late_punch_index:
                        # Late punch, 10-25 minutes late
                        offset = random.randint(10, 25)
                    else:
                        # On-time punch, +/- 5 minutes
                        offset = random.randint(-5, 5)
                        
                    punch_time = base_time + timedelta(minutes=offset)
                    
                    writer.writerow([
                        emp_id,
                        emp_data['name'],
                        emp_data['dept'],
                        punch_time.strftime('%Y-%m-%d'),
                        punch_time.strftime('%H:%M')
                    ])

if __name__ == "__main__":
    generate_fingerprints()
    print("fingerprints.csv has been generated successfully.")
