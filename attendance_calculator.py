"""
Attendance Calculator Module
Implements the attendance calculation logic based on specified rules
"""

import pandas as pd
from datetime import datetime, timedelta
from collections import defaultdict


class AttendanceCalculator:
    def __init__(self, required_times=None, tolerance_minutes=30, time_normalization_map=None, late_threshold=3, absence_threshold=3, debug_mode=False):
        """
        Initialize the attendance calculator
        
        Args:
            required_times: List of required check-in times (default: [08:00, 12:00, 15:00, 20:00, 23:00, 08:00 next day])
            tolerance_minutes: Tolerance in minutes for check-in times (default: 30)
            time_normalization_map: Dictionary mapping input times to normalized times (default: {'03:00': '15:00', '03:00:00': '15:00'})
            late_threshold: Number of late arrivals that constitute an absence (default: 3)
            absence_threshold: Number of missing punches that constitute 1 absence day (default: 3)
            debug_mode: Enable debug logging (default: False)
        """
        if required_times is None:
            # Default required times according to requirements
            self.required_times = [
                "08:00",  # Start of work
                "12:00",  # End of morning
                "15:00",  # Start of shift
                "20:00",  # End of shift
                "23:00",  # End of night
                "08:00"   # End of shift (next day)
            ]
        else:
            self.required_times = required_times
            
        self.tolerance_minutes = tolerance_minutes
        self.late_threshold = late_threshold
        self.absence_threshold = absence_threshold  # Number of missing punches = 1 absence day
        
        # Store time descriptions from settings (will be updated when settings are provided)
        self.time_descriptions = {}  # Maps time_str to description
        
        # Default time normalization mapping based on your specification
        if time_normalization_map is None:
            self.time_normalization_map = {
                '03:00': '15:00',
                '03:00:00': '15:00'
                # Additional mappings can be added here as needed
            }
        else:
            self.time_normalization_map = time_normalization_map
        
        # Set debug mode
        self.debug_mode = debug_mode
        
    def calculate_attendance(self, fingerprint_data, shift_data, settings=None):
        """
        Calculate attendance based on fingerprint data and shift schedule
        
        Args:
            fingerprint_data: DataFrame with fingerprint records (Name, Department, Date, Time)
            shift_data: DataFrame with shift schedule (Name, Shift Date)
            settings: Optional settings dict with custom times and tolerance
        
        Returns:
            DataFrame with attendance results
        """
        # Clean column names by stripping whitespace (including tabs)
        if fingerprint_data is not None:
            fingerprint_data.columns = fingerprint_data.columns.str.strip()
            # Remove duplicate rows (exact duplicates based on all columns)
            initial_count = len(fingerprint_data)
            fingerprint_data = fingerprint_data.drop_duplicates()
            if len(fingerprint_data) < initial_count:
                if self.debug_mode:
                    print(f"DEBUG: Removed {initial_count - len(fingerprint_data)} duplicate fingerprint records (exact duplicates)")
            
            # Remove duplicates based on Name, Date, Time (same employee, same date, same time)
            initial_count = len(fingerprint_data)
            fingerprint_data = fingerprint_data.drop_duplicates(subset=['Name', 'Date', 'Time'], keep='first')
            if len(fingerprint_data) < initial_count and self.debug_mode:
                removed = initial_count - len(fingerprint_data)
                print(f"DEBUG: Removed {removed} duplicate fingerprint records (same Name, Date, Time)")
            
            # Remove fingerprints that are within 10 minutes of each other for the same employee
            # Create DateTime column for comparison
            if not fingerprint_data.empty:
                try:
                    fingerprint_data['DateTime'] = pd.to_datetime(
                        fingerprint_data['Date'].astype(str) + ' ' + fingerprint_data['Time'].astype(str),
                        format='%Y-%m-%d %H:%M',
                        errors='coerce'
                    )
                    
                    # Sort by Name and DateTime
                    fingerprint_data = fingerprint_data.sort_values(['Name', 'DateTime']).reset_index(drop=True)
                    
                    # Remove fingerprints within 10 minutes
                    indices_to_remove = set()
                    for i in range(len(fingerprint_data)):
                        if i in indices_to_remove:
                            continue
                        
                        current_name = fingerprint_data.iloc[i]['Name']
                        current_datetime = fingerprint_data.iloc[i]['DateTime']
                        
                        if pd.isna(current_datetime):
                            continue
                        
                        # Check next fingerprints for the same employee
                        for j in range(i + 1, len(fingerprint_data)):
                            if j in indices_to_remove:
                                continue
                            
                            next_name = fingerprint_data.iloc[j]['Name']
                            next_datetime = fingerprint_data.iloc[j]['DateTime']
                            
                            if pd.isna(next_datetime):
                                continue
                            
                            # If different employee, stop checking
                            if current_name != next_name:
                                break
                            
                            # Calculate time difference
                            time_diff = (next_datetime - current_datetime).total_seconds() / 60  # in minutes
                            
                            # If within 10 minutes, mark for removal
                            if time_diff < 10:
                                indices_to_remove.add(j)
                            else:
                                # If more than 10 minutes, stop checking for this fingerprint
                                break
                    
                    # Remove marked indices
                    if indices_to_remove:
                        fingerprint_data = fingerprint_data.drop(fingerprint_data.index[list(indices_to_remove)]).reset_index(drop=True)
                        if self.debug_mode:
                            print(f"DEBUG: Removed {len(indices_to_remove)} duplicate fingerprint records (within 10 minutes)")
                    
                    # Drop DateTime column as it's temporary
                    fingerprint_data = fingerprint_data.drop('DateTime', axis=1)
                    
                except Exception as e:
                    if self.debug_mode:
                        print(f"DEBUG: Error removing fingerprints within 10 minutes: {e}")
        
        if shift_data is not None:
            shift_data.columns = shift_data.columns.str.strip()
            # Remove duplicate rows (exact duplicates)
            initial_count = len(shift_data)
            shift_data = shift_data.drop_duplicates()
            if len(shift_data) < initial_count:
                if self.debug_mode:
                    print(f"DEBUG: Removed {initial_count - len(shift_data)} duplicate shift records (exact duplicates)")
            
            # Remove duplicates based on Name, Shift Date (same employee, same shift date)
            initial_count = len(shift_data)
            shift_data = shift_data.drop_duplicates(subset=['Name', 'Shift Date'], keep='first')
            if len(shift_data) < initial_count and self.debug_mode:
                removed = initial_count - len(shift_data)
                print(f"DEBUG: Removed {removed} duplicate shift records (same Name, Shift Date)")
        
        if settings:
            # Extract time strings and descriptions from settings['times']
            # settings['times'] contains tuples: (time_str, description)
            times_from_settings = settings.get('times', self.required_times)
            
            if times_from_settings and isinstance(times_from_settings[0], tuple):
                # If times are tuples (time_str, description), extract both
                self.required_times = [time[0] if isinstance(time, tuple) else time for time in times_from_settings]
                # Store descriptions mapping time_str -> description
                self.time_descriptions = {}
                for time_tuple in times_from_settings:
                    if isinstance(time_tuple, tuple) and len(time_tuple) >= 2:
                        time_str = str(time_tuple[0])
                        description = time_tuple[1]
                        self.time_descriptions[time_str] = description
            else:
                self.required_times = times_from_settings
                # If no descriptions provided, use default descriptions
                self.time_descriptions = {}
            
            self.tolerance_minutes = settings.get('tolerance', self.tolerance_minutes)
            self.late_threshold = settings.get('late_threshold', self.late_threshold)
            self.absence_threshold = settings.get('absence_threshold', self.absence_threshold)
        
        # Validate input data
        if fingerprint_data is None or fingerprint_data.empty:
            print("Fingerprint data is empty or None")
            return pd.DataFrame()
        
        if shift_data is None or shift_data.empty:
            print("Shift data is empty or None")
            return pd.DataFrame()
        
        # Ensure required columns exist in fingerprint data
        required_fingerprint_cols = ['Name', 'Department', 'Date', 'Time']
        missing_fingerprint_cols = [col for col in required_fingerprint_cols if col not in fingerprint_data.columns]
        if missing_fingerprint_cols:
            print(f"Missing required column(s) in fingerprint data: {', '.join(missing_fingerprint_cols)}")
            return pd.DataFrame()

        # Ensure required columns exist in shift data
        required_shift_cols = ['Name', 'Shift Date']
        missing_shift_cols = [col for col in required_shift_cols if col not in shift_data.columns]
        if missing_shift_cols:
            print(f"Missing required column(s) in shift data: {', '.join(missing_shift_cols)}")
            return pd.DataFrame()
        
        # Process the data
        results = []
        
        if self.debug_mode:
            print(f"DEBUG: Starting attendance calculation for {len(shift_data)} shift records")
        
        try:
            # Group fingerprint data by employee name and date
            grouped_fingerprint = fingerprint_data.groupby(['Name', 'Date'])
        except KeyError as e:
            print(f"Error grouping fingerprint data: {e}")
            return pd.DataFrame()
        
        try:
            # Get unique employees from both datasets (using Name)
            all_employees = set(fingerprint_data['Name'].unique()) | set(shift_data['Name'].unique())
        except KeyError as e:
            print(f"Error accessing Name column: {e}")
            return pd.DataFrame()
        
        if self.debug_mode:
            print(f"DEBUG: Processing {len(all_employees)} unique employees")
        
        # Create a mapping of employee name to their total shift days for Total Working Days calculation
        employee_shift_counts = shift_data.groupby('Name').size().to_dict()
        
        for emp_name in all_employees:
            try:
                # Get shift schedule for this employee
                emp_shifts = shift_data[shift_data['Name'] == emp_name]
                
                if self.debug_mode:
                    print(f"\nDEBUG: Processing employee {emp_name} with {len(emp_shifts)} shift days")
                
                for _, shift_row in emp_shifts.iterrows():
                    try:
                        shift_date = shift_row['Shift Date']
                        
                        # Convert to datetime if it's not already
                        if isinstance(shift_date, str):
                            shift_date = pd.to_datetime(shift_date).date()
                        elif isinstance(shift_date, pd.Timestamp):
                            shift_date = shift_date.date()
                        else:
                            shift_date = shift_date
                            
                        # Only process days that are in the shift schedule
                        # Get fingerprint records for this employee on this date and next day
                        emp_date_fingerprints = self._get_fingerprints_for_date_range(
                            grouped_fingerprint, emp_name, shift_date
                        )
                        
                        if self.debug_mode:
                            print(f"DEBUG: Found {len(emp_date_fingerprints)} fingerprint records for employee {emp_name} on {shift_date}")
                        
                        # Calculate attendance for this day
                        day_result = self._calculate_day_attendance(
                            emp_name, emp_date_fingerprints, shift_date
                        )
                        
                        results.append(day_result)
                    except KeyError as e:
                        print(f"Error accessing shift data: {e}")
                        continue
                    except Exception as e:
                        print(f"Error processing shift for employee {emp_name}: {e}")
                        continue
            except Exception as e:
                print(f"Error processing employee {emp_name}: {e}")
                continue
        
        # Store the shift counts to be used in aggregation
        self.employee_shift_counts = employee_shift_counts
        
        if not results:
            print("No results generated - check your data")
            return pd.DataFrame()
        
        # Convert results to DataFrame
        results_df = pd.DataFrame(results)
        
        # Store daily results for detailed view
        self.daily_results = results_df
        
        # Aggregate results by employee
        aggregated_results = self._aggregate_employee_results(results_df)
        
        return aggregated_results
        
    def _get_fingerprints_for_date_range(self, grouped_fingerprint, emp_name, shift_date):
        """
        Get fingerprint records for an employee on a specific date and the next day
        (for handling 08:00 next day check-out)
        """
        try:
            # Get records for the shift date
            day_records = grouped_fingerprint.get_group((emp_name, shift_date.strftime('%Y-%m-%d')))
        except KeyError:
            # If no records for this date, create empty DataFrame with same columns
            day_records = pd.DataFrame(columns=['Name', 'Department', 'Date', 'Time'])
        
        # Get records for the next day (for 08:00 next day check-out)
        next_day = shift_date + timedelta(days=1)
        try:
            next_day_records = grouped_fingerprint.get_group((emp_name, next_day.strftime('%Y-%m-%d')))
        except KeyError:
            # If no records for next day, create empty DataFrame with same columns
            next_day_records = pd.DataFrame(columns=['Name', 'Department', 'Date', 'Time'])
        
        # Combine both sets of records, but only if both are non-empty or handle empty DataFrames properly
        if day_records.empty and next_day_records.empty:
            all_records = pd.DataFrame(columns=['Name', 'Department', 'Date', 'Time'])
        elif day_records.empty:
            all_records = next_day_records.copy()
        elif next_day_records.empty:
            all_records = day_records.copy()
        else:
            all_records = pd.concat([day_records, next_day_records], ignore_index=True)
        
        # Convert time to datetime for proper sorting and comparison
        # Apply time normalization using the instance's mapping
        if not all_records.empty:
            # Normalize time values before creating DateTime
            normalized_times = []
            for _, record in all_records.iterrows():
                time_str = str(record['Time'])
                # Apply time normalization mapping from instance variable
                normalized_time = self.time_normalization_map.get(time_str, time_str)
                normalized_times.append(normalized_time)
            
            all_records = all_records.copy()  # To avoid SettingWithCopyWarning
            all_records['Time'] = normalized_times
            
            # Validate time format before creating DateTime
            try:
                all_records['DateTime'] = pd.to_datetime(
                    all_records['Date'].astype(str) + ' ' + all_records['Time'].astype(str),
                    format='%Y-%m-%d %H:%M',
                    errors='coerce'
                )
                
                # Remove rows with invalid DateTime (NaT)
                invalid_count = all_records['DateTime'].isna().sum()
                if invalid_count > 0:
                    if self.debug_mode:
                        print(f"DEBUG: Removed {invalid_count} records with invalid DateTime")
                    all_records = all_records[all_records['DateTime'].notna()].copy()
                    
            except Exception as e:
                if self.debug_mode:
                    print(f"DEBUG: Error creating DateTime: {e}")
                # Fallback: create empty DataFrame with DateTime column
                all_records['DateTime'] = pd.Series([], dtype='datetime64[ns]')
        else:
            # If no records, create empty DataFrame with DateTime column
            all_records['DateTime'] = pd.Series([], dtype='datetime64[ns]')
        
        return all_records.sort_values('DateTime')
        
    def _calculate_day_attendance(self, emp_name, fingerprints, shift_date):
        """
        Calculate attendance for a specific employee on a specific date
        """
        # Get employee info from the first fingerprint record
        if len(fingerprints) > 0:
            emp_name_from_data = fingerprints['Name'].iloc[0]
            emp_dept = fingerprints['Department'].iloc[0]
        else:
            # If no fingerprints, use the provided name and try to get department from shift data
            emp_name_from_data = emp_name
            emp_dept = "Unknown"
        
        if self.debug_mode:
            print(f"\nDEBUG: Calculating attendance for Employee {emp_name_from_data} on {shift_date}")
        
        # Required times for the shift (including next day's 08:00)
        required_times = self._get_required_times_for_date(shift_date)
        
        if self.debug_mode:
            print(f"DEBUG: Required times for {shift_date}: {[rt['time_str'] for rt in required_times]}")
        
        # Find matching fingerprints for each required time
        attendance_status = self._match_fingerprints_to_required_times(fingerprints, required_times)
        
        # Determine the day status
        # Check if there are any punches for this employee on this specific shift day only
        # Exclude punches from the next day (for the 08:00 next day check-out)
        day_punches = fingerprints[fingerprints['Date'] == shift_date.strftime('%Y-%m-%d')]
        has_punches_on_day = len(day_punches) > 0
        
        if self.debug_mode:
            print(f"DEBUG: Has punches on day: {has_punches_on_day}, Day punches count: {len(day_punches)}")
        
        day_status = self._determine_day_status(attendance_status, has_punches_on_day, self.late_threshold if hasattr(self, 'late_threshold') else 3)
        
        # Calculate attendance metrics
        attendance_metrics = self._calculate_attendance_metrics(attendance_status)
        
        if self.debug_mode:
            print(f"DEBUG: Final day status for {emp_name_from_data} on {shift_date}: {day_status}")
        
        return {
            'Name': emp_name_from_data,
            'Department': emp_dept,
            'Date': shift_date,
            'Required Times': required_times,
            'Attendance Status': attendance_status,
            'Day Status': day_status,
            'Present Days': 1 if day_status in ['مستوفي', 'نقص بصمة'] else 0,  # Days with punches
            'Complete Days': 1 if day_status == 'مستوفي' else 0,
            'Incomplete Days': 1 if day_status == 'نقص بصمة' else 0,
            'Absent Days': 1 if day_status == 'غائب' else 0,
            **attendance_metrics
        }
        
    def _get_required_times_for_date(self, shift_date):
        """
        Get required check-in times for a specific date
        Uses times and descriptions from settings as the authoritative source
        """
        required_times = []
        
        for i, time_str in enumerate(self.required_times):
            # Create datetime for the required time
            # Ensure time_str is a string
            if isinstance(time_str, tuple):
                time_str = str(time_str[0]) if len(time_str) > 0 else "08:00"
            elif not isinstance(time_str, str):
                time_str = str(time_str)
            
            # Parse time
            try:
                hour, minute = map(int, time_str.split(':'))
            except (ValueError, AttributeError):
                # Fallback to default time if parsing fails
                hour, minute = 8, 0
                time_str = "08:00"
            
            # Determine if this is the last time (next day's check)
            # The last time in the list is always for the next day
            is_last_time = (i == len(self.required_times) - 1)
            
            if is_last_time:
                time_dt = datetime.combine(shift_date + timedelta(days=1), 
                                         datetime.min.time().replace(hour=hour, minute=minute))
            else:
                time_dt = datetime.combine(shift_date, 
                                         datetime.min.time().replace(hour=hour, minute=minute))
            
            # Get description from settings if available, otherwise use default
            if time_str in self.time_descriptions:
                description = self.time_descriptions[time_str]
            else:
                description = self._get_time_description(i)
            
            required_times.append({
                'time': time_dt,
                'time_str': time_str,
                'description': description  # Use description from settings
            })
        
        return required_times
        
    def _get_time_description(self, index):
        """
        Get description for each required time
        """
        descriptions = [
            "بداية الدوام",
            "نهاية الصباح", 
            "بداية الوردية",
            "نهاية الوردية",
            "نهاية الليل",
            "نهاية المناوبة (اليوم التالي)"
        ]
        
        return descriptions[index] if index < len(descriptions) else f"الوقت {index+1}"
        
    def _match_fingerprints_to_required_times(self, fingerprints, required_times):
        """
        Match fingerprint records to required check-in times within tolerance
        Following the authoritative specification:
        - Only match within tolerance window
        - Pick closest to target datetime
        - No AM/PM guessing or manual conversion
        - CRITICAL: Each fingerprint can only be matched to ONE required time
        """
        attendance_status = []
        
        # Track used fingerprints to prevent double-counting
        # Use a set to track which fingerprint DateTime values have been used
        used_fingerprints = set()
        
        for req_time_info in required_times:
            req_time = req_time_info['time']
            req_time_str = req_time_info['time_str']
            req_description = req_time_info['description']
            
            # Find fingerprints within tolerance window that haven't been used yet
            tolerance_start = req_time - timedelta(minutes=self.tolerance_minutes)
            tolerance_end = req_time + timedelta(minutes=self.tolerance_minutes)
            
            # Find matching fingerprints within tolerance, excluding already used ones
            matching_fingerprints = fingerprints[
                (fingerprints['DateTime'] >= tolerance_start) & 
                (fingerprints['DateTime'] <= tolerance_end) &
                (~fingerprints['DateTime'].isin(used_fingerprints))  # Exclude already used fingerprints
            ]
            
            # Handle duplicate fingerprints in the same minute (only count one)
            unique_minutes = set()
            valid_fingerprints = []
            
            for _, fp in matching_fingerprints.iterrows():
                minute_key = fp['DateTime'].strftime('%Y-%m-%d %H:%M')
                if minute_key not in unique_minutes:
                    unique_minutes.add(minute_key)
                    valid_fingerprints.append(fp)
            
            matched = len(valid_fingerprints) > 0
            actual_time = None
            
            # If matches found, pick the closest to target datetime
            if valid_fingerprints:
                closest_fingerprint = min(valid_fingerprints, key=lambda x: abs((x['DateTime'] - req_time).total_seconds()))
                actual_time = closest_fingerprint['DateTime']
                
                # Mark this fingerprint as used to prevent reuse
                used_fingerprints.add(actual_time)
                
                # Calculate delay if matched
                delay = (actual_time - req_time).total_seconds() / 60  # Convert to minutes
                if self.debug_mode:
                    status = "Matched" if matched else "Not Matched"
                    delay_str = f"(Late by {delay:.1f} min)" if delay > 0 else f"(Early by {abs(delay):.1f} min)" if delay < 0 else "(On time)"
                    print(f"DEBUG: Required {req_time_str} ({req_description}) - {status} - Actual: {actual_time.strftime('%H:%M')} {delay_str}")
            else:
                if self.debug_mode:
                    print(f"DEBUG: Required {req_time_str} ({req_description}) - Not Matched - No fingerprint within tolerance [{tolerance_start.strftime('%H:%M')}-{tolerance_end.strftime('%H:%M')}]")
            
            attendance_status.append({
                'required_time': req_time,
                'required_time_str': req_time_str,
                'description': req_description,
                'matched': matched,
                'actual_time': actual_time,
                'tolerance_start': tolerance_start,
                'tolerance_end': tolerance_end
            })
        
        return attendance_status
        
    def _determine_day_status(self, attendance_status, has_punches_on_day, late_threshold=3):
        """
        Determine the day status based on attendance
        According to the correct specification:
        - Absent (غائب) if there are 0 punches on a shift day ONLY
        - Otherwise, if there are punches but not all required times, it's 'نقص بصمة (X)'
        - If all required times are met, it's 'مستوفي'
        - Late arrivals are tracked separately as indicators, not as absence triggers
        """
        matched_count = sum(1 for status in attendance_status if status['matched'])
        total_required = len(attendance_status)
        missing_count = total_required - matched_count
        
        # According to the correct specification: employee is absent ONLY if:
        # There are 0 punches on a shift day
        # Late arrivals do NOT trigger absence - they are tracked as separate metrics
        if not has_punches_on_day:
            if self.debug_mode:
                print(f"DEBUG: Employee has no punches on shift day - marking as غائب (Absent)")
            return 'غائب'  # Absent - no punches at all on shift day
        
        # If there are punches, check if they are complete or incomplete
        if matched_count == total_required:
            if self.debug_mode:
                print(f"DEBUG: Employee has all required punches ({matched_count}/{total_required}) - marking as مستوفي (Complete)")
            return 'مستوفي'  # Complete - all required punches present
        else:
            if self.debug_mode:
                print(f"DEBUG: Employee has some punches but not all ({matched_count}/{total_required}) - marking as نقص بصمة (Incomplete)")
            # Include missing count in status: "نقص بصمة (2)"
            return f'نقص بصمة ({missing_count})'  # Incomplete - some punches present but not all required
            
        # NOTE: Late arrivals exceeding threshold does NOT make someone absent
        # Late metrics are tracked separately for performance analysis

            
    def _calculate_attendance_metrics(self, attendance_status):
        """
        Calculate attendance metrics following the authoritative specification
        IMPORTANT: Late is only counted if delay exceeds tolerance_minutes
        """
        matched_count = sum(1 for status in attendance_status if status['matched'])
        total_required = len(attendance_status)
            
        # Calculate late metrics
        # Late is only counted if the actual time exceeds the required time by MORE than tolerance_minutes
        late_count = 0
        late_minutes = 0
            
        for status in attendance_status:
            if status['matched']:
                # Calculate delay if the actual time is after the required time
                if status['actual_time'] and status['required_time']:
                    delay = (status['actual_time'] - status['required_time']).total_seconds() / 60  # Convert to minutes
                    # Only count as late if delay exceeds tolerance_minutes
                    # If delay is within tolerance (0 to tolerance_minutes), it's not considered late
                    if delay > self.tolerance_minutes:
                        late_count += 1
                        # Only count the minutes beyond tolerance as late minutes
                        late_minutes += (delay - self.tolerance_minutes)
            
        metrics = {
            'Required Checks': total_required,
            'Actual Checks': matched_count,
            'Missing Checks': total_required - matched_count,
            'LateCount': late_count,
            'LateMinutes': late_minutes
        }
            
        if total_required > 0:
            metrics['Compliance Rate'] = round((matched_count / total_required) * 100, 2)
        else:
            metrics['Compliance Rate'] = 0.0
            
        return metrics
        
    def _aggregate_employee_results(self, daily_results):
        """
        Aggregate daily results by employee
        """
        if daily_results.empty:
            return pd.DataFrame()
        
        # Group by employee and aggregate
        agg_functions = {
            'Present Days': 'sum',
            'Complete Days': 'sum', 
            'Incomplete Days': 'sum',
            'Absent Days': 'sum',
            'LateCount': 'sum',
            'LateMinutes': 'sum',
            'Required Checks': 'sum',
            'Actual Checks': 'sum',
            'Missing Checks': 'sum'
        }
        
        # Calculate average compliance rate
        employee_summary = daily_results.groupby(['Name', 'Department']).agg(agg_functions).reset_index()
        
        # Calculate compliance rate as Total Actual Checks / Total Required Checks
        # This is more accurate than averaging daily compliance rates
        total_checks = daily_results.groupby(['Name', 'Department'])[['Actual Checks', 'Required Checks']].sum().reset_index()
        total_checks['Compliance Rate'] = 0.0
        # Calculate compliance rate for each employee
        mask = total_checks['Required Checks'] > 0
        total_checks.loc[mask, 'Compliance Rate'] = round((total_checks.loc[mask, 'Actual Checks'] / total_checks.loc[mask, 'Required Checks']) * 100, 2)
        # Update the compliance rate in employee_summary
        for idx, row in employee_summary.iterrows():
            emp_mask = (total_checks['Name'] == row['Name']) & \
                         (total_checks['Department'] == row['Department'])
            if emp_mask.any():
                employee_summary.loc[idx, 'Compliance Rate'] = total_checks[emp_mask]['Compliance Rate'].iloc[0]
        
        # Calculate total working days based on shift schedule, not analysis results
        # Total working days should be the count of shift days for each employee
        # Use the stored shift counts from the original shift data
        if hasattr(self, 'employee_shift_counts'):
            employee_summary['Total Working Days'] = 0
            for idx, row in employee_summary.iterrows():
                emp_name = row['Name']
                if emp_name in self.employee_shift_counts:
                    employee_summary.loc[idx, 'Total Working Days'] = self.employee_shift_counts[emp_name]
        else:
            # Fallback: calculate from daily results
            shift_day_counts = daily_results.groupby(['Name', 'Department']).size().reset_index(name='Schedule_Count')
            employee_summary = employee_summary.merge(shift_day_counts, on=['Name', 'Department'], how='left')
            employee_summary['Total Working Days'] = employee_summary['Schedule_Count']
            # Drop the temporary column
            employee_summary = employee_summary.drop('Schedule_Count', axis=1)
        
        # Reorder columns to match the expected output
        column_order = [
            'Name', 'Department', 'Total Working Days',
            'Complete Days', 'Incomplete Days', 'Absent Days', 
            'LateCount', 'LateMinutes',
            'Required Checks', 'Actual Checks', 'Missing Checks',
            'Compliance Rate', 'Absence Reason', 'FinalStatus'
        ]
        
        # Ensure all columns exist
        for col in column_order:
            if col not in employee_summary.columns:
                employee_summary[col] = 0
        
        # Calculate Absent Days based on Missing Checks
        # Every absence_threshold missing punches = 1 absence day
        # Example: 12 missing punches with threshold=3 = 4 absence days
        employee_summary['Absent Days'] = 0
        employee_summary['Absence Reason'] = ''
        
        for idx, row in employee_summary.iterrows():
            missing_checks = row.get('Missing Checks', 0)
            # Calculate absence days: missing_checks / absence_threshold (integer division)
            if missing_checks > 0:
                absence_days = missing_checks // self.absence_threshold
                employee_summary.loc[idx, 'Absent Days'] = absence_days
                
                # Create absence reason
                if absence_days > 0:
                    reason = f"عدد البصمات المتروكة: {missing_checks}، كل {self.absence_threshold} بصمة = يوم غياب"
                    employee_summary.loc[idx, 'Absence Reason'] = reason
                else:
                    employee_summary.loc[idx, 'Absence Reason'] = f"عدد البصمات المتروكة: {missing_checks} (أقل من عتبة الغياب)"
            else:
                employee_summary.loc[idx, 'Absence Reason'] = "لا توجد بصمات متروكة"
        
        # Calculate FinalStatus according to the specification
        # IF IncompleteDays == 0 AND AbsentDays == 0: FinalStatus = "ملتزم" ELSE: "غير ملتزم"
        # Add this as a separate column in the final output
        employee_summary['FinalStatus'] = employee_summary.apply(
            lambda row: 'ملتزم' if row['Incomplete Days'] == 0 and row['Absent Days'] == 0 else 'غير ملتزم',
            axis=1
        )
        
        # Add FinalStatus to the column order
        column_order.append('FinalStatus')
        
        return employee_summary[column_order]
    
    def get_daily_results(self):
        """
        Return the daily attendance results for detailed view
        """
        if hasattr(self, 'daily_results'):
            return self.daily_results
        else:
            return pd.DataFrame()


def main():
    """
    Example usage of the AttendanceCalculator
    """
    # Example data for testing
    fingerprint_data = pd.DataFrame({
        'Name': ['Ahmad', 'Ahmad', 'Ahmad', 'Ahmad', 'Ahmad', 'Fatima', 'Fatima', 'Fatima'],
        'Department': ['Admin', 'Admin', 'Admin', 'Admin', 'Admin', 'Accounting', 'Accounting', 'Accounting'],
        'Date': ['2023-01-15', '2023-01-15', '2023-01-15', '2023-01-15', '2023-01-15', '2023-01-15', '2023-01-15', '2023-01-15'],
        'Time': ['08:05', '12:10', '14:55', '20:05', '22:58', '08:30', '12:00', '15:10']
    })
    
    shift_data = pd.DataFrame({
        'Name': ['Ahmad', 'Fatima'],
        'Shift Date': ['2023-01-15', '2023-01-15']
    })
    
    # Create calculator instance
    calculator = AttendanceCalculator()
    
    # Calculate attendance
    results = calculator.calculate_attendance(fingerprint_data, shift_data)
    
    print("Attendance Calculation Results:")
    print(results)
    

if __name__ == '__main__':
    main()