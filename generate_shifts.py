import csv
from datetime import date, timedelta

def generate_shift_schedule():
    start_date = date(2025, 1, 1)
    end_date = date(2025, 12, 31)
    employee_ids = [f"E{i:03}" for i in range(1, 21)]
    
    with open('shift_schedule.csv', 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['Employee_ID', 'Shift_Date'])
        
        current_date = start_date
        employee_index = 0
        while current_date <= end_date:
            emp1_id = employee_ids[employee_index % 20]
            emp2_id = employee_ids[(employee_index + 1) % 20]
            
            writer.writerow([emp1_id, current_date.strftime('%Y-%m-%d')])
            writer.writerow([emp2_id, current_date.strftime('%Y-%m-%d')])
            
            employee_index = (employee_index + 2) % 20
            current_date += timedelta(days=1)

if __name__ == "__main__":
    generate_shift_schedule()