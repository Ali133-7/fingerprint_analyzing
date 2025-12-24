"""
Debug script to understand the absence logic behavior
"""

import pandas as pd
from datetime import datetime, timedelta
from attendance_calculator import AttendanceCalculator

def debug_absence_logic():
    # Test with a simpler scenario
    fingerprint_data = pd.DataFrame({
        'Employee ID': [1, 1, 1, 1, 1, 1],  # 6 required times, all late
        'Name': ['Test Employee', 'Test Employee', 'Test Employee', 'Test Employee', 'Test Employee', 'Test Employee'],
        'Department': ['Test Dept', 'Test Dept', 'Test Dept', 'Test Dept', 'Test Dept', 'Test Dept'],
        'Date': ['2025-12-12', '2025-12-12', '2025-12-12', '2025-12-12', '2025-12-12', '2025-12-12'],
        'Time': ['08:45', '12:45', '15:45', '20:45', '23:45', '08:45']  # All 45 minutes late
    })

    shift_data = pd.DataFrame({
        'Employee ID': [1],
        'Shift Date': ['2025-12-12']
    })

    print("Debugging absence logic...")
    print("Fingerprint Data:")
    print(fingerprint_data)
    print("\nShift Data:")
    print(shift_data)

    # Test with late threshold = 3
    print("\n" + "="*50)
    print("TEST: Late threshold = 3, tolerance = 30 minutes")
    calculator = AttendanceCalculator(
        required_times=["08:00", "12:00", "15:00", "20:00", "23:00", "08:00"],
        tolerance_minutes=30,
        late_threshold=3,
        time_normalization_map={}
    )

    # Manually test the day calculation
    grouped_fingerprint = fingerprint_data.groupby(['Employee ID', 'Date'])
    
    print("\nManual day calculation for 2025-12-12:")
    emp_date_fingerprints = calculator._get_fingerprints_for_date_range(
        grouped_fingerprint, 1, pd.to_datetime('2025-12-12').date()
    )
    print("Fingerprints with DateTime:")
    print(emp_date_fingerprints[['Date', 'Time', 'DateTime']])
    
    # Get required times for the date
    required_times = calculator._get_required_times_for_date(pd.to_datetime('2025-12-12').date())
    print("\nRequired times:")
    for req_time in required_times:
        print(f"  {req_time['time_str']} -> {req_time['time']} ({req_time['description']})")
    
    # Match fingerprints to required times
    attendance_status = calculator._match_fingerprints_to_required_times(emp_date_fingerprints, required_times)
    print("\nAttendance status:")
    for i, status in enumerate(attendance_status):
        print(f"  Required: {status['required_time_str']} at {status['required_time']}")
        print(f"    Matched: {status['matched']}")
        print(f"    Actual: {status['actual_time']}")
        print(f"    Tolerance: {status['tolerance_start']} to {status['tolerance_end']}")
        if status['matched'] and status['actual_time'] and status['required_time']:
            delay = (status['actual_time'] - status['required_time']).total_seconds() / 60
            print(f"    Delay: {delay} minutes")
        print()
    
    # Calculate late arrivals in _determine_day_status logic
    matched_count = sum(1 for status in attendance_status if status['matched'])
    total_required = len(attendance_status)
    late_arrivals = 0
    for status in attendance_status:
        if status['matched'] and status['actual_time'] and status['required_time']:
            delay = (status['actual_time'] - status['required_time']).total_seconds() / 60  # Convert to minutes
            if delay > 0:  # Only count actual late arrivals (positive delay)
                late_arrivals += 1
                print(f"Late arrival detected: {delay} minutes")
    
    print(f"\nMatched count: {matched_count}")
    print(f"Total required: {total_required}")
    print(f"Late arrivals: {late_arrivals}")
    print(f"Late threshold: {calculator.late_threshold}")
    
    # Determine day status
    day_punches = emp_date_fingerprints[emp_date_fingerprints['Date'] == '2025-12-12']
    has_punches_on_day = len(day_punches) > 0
    print(f"Has punches on day: {has_punches_on_day}")
    
    day_status = calculator._determine_day_status(attendance_status, has_punches_on_day, calculator.late_threshold)
    print(f"Day status: {day_status}")
    
    # Calculate metrics
    attendance_metrics = calculator._calculate_attendance_metrics(attendance_status)
    print(f"\nAttendance metrics: {attendance_metrics}")
    
    # Complete day calculation
    day_result = calculator._calculate_day_attendance(1, emp_date_fingerprints, pd.to_datetime('2025-12-12').date())
    print(f"\nDay result: {day_result}")

if __name__ == "__main__":
    debug_absence_logic()