import csv
import pandas as pd
from typing import List, Dict, Any
from .models import CSVRow, LocationSchedule, MultiLocationSchedule, ScheduleEntry, DaySchedule, WeeklySchedule, DayOfWeek, SpecialityLevelSchedule, MultiLevelSchedule


class CSVConverter:
    """Converts CSV/Excel data to structured JSON format"""
    
    def __init__(self):
        self.day_mapping = {
            "الاحد": "sunday",
            "الاثنين": "monday", 
            "الثلاثاء": "tuesday",
            "الاربعاء": "wednesday",
            "الخميس": "thursday"
        }
        
        self.slot_mapping = {
            1: "first",
            2: "second", 
            3: "third",
            4: "fourth"
        }
    
    def read_csv(self, file_path: str) -> List[CSVRow]:
        """Read CSV file and convert to list of CSVRow objects"""
        rows = []
        ignored_rows = []
        
        with open(file_path, 'r', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            for row in reader:
                try:
                    # Filter out any None keys from the row
                    cleaned_row = {k: v for k, v in row.items() if k is not None}
                    
                    # Check if code is None/empty - if so, ignore this row
                    if not cleaned_row.get('code'):
                        ignored_rows.append(row)
                        print(f"⚠️  Ignoring row with empty code: {row}")
                        continue
                    
                    csv_row = CSVRow(**cleaned_row)
                    rows.append(csv_row)
                except Exception as e:
                    print(f"Error parsing row: {row}, Error: {e}")
                    ignored_rows.append(row)
        
        if ignored_rows:
            print(f"📊 Total rows ignored: {len(ignored_rows)}")
        
        return rows
    
    def read_excel(self, file_path: str) -> List[CSVRow]:
        """Read Excel file from 'table_full' sheet and convert to list of CSVRow objects"""
        try:
            # Read the 'table_full' sheet
            df = pd.read_excel(file_path, sheet_name='table_full')
            
            # Convert DataFrame to list of dictionaries
            rows_data = df.to_dict('records')
            
            rows = []
            ignored_rows = []
            
            for row_data in rows_data:
                try:
                    # Filter out any None keys and convert NaN to None
                    cleaned_row = {}
                    for k, v in row_data.items():
                        if k is not None:
                            # Convert pandas NaN to None
                            if pd.isna(v):
                                cleaned_row[k] = None
                            else:
                                # Convert to string, but handle None values properly
                                if v is not None:
                                    cleaned_row[k] = str(v)
                                else:
                                    cleaned_row[k] = None
                    
                    # Check if code is None/NaN - if so, ignore this row
                    if cleaned_row.get('code') is None:
                        ignored_rows.append(row_data)
                        print(f"⚠️  Ignoring row with NaN code: {row_data}")
                        continue
                    
                    csv_row = CSVRow(**cleaned_row)
                    rows.append(csv_row)
                except Exception as e:
                    print(f"Error parsing row: {row_data}, Error: {e}")
                    ignored_rows.append(row_data)
                    # Continue processing other rows instead of failing completely
                    continue
            
            if ignored_rows:
                print(f"📊 Total rows ignored: {len(ignored_rows)}")
            
            return rows
            
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            return []
    
    def read_file(self, file_path: str) -> List[CSVRow]:
        """Read file (CSV or Excel) and convert to list of CSVRow objects"""
        if file_path.lower().endswith(('.xlsx', '.xls')):
            return self.read_excel(file_path)
        else:
            return self.read_csv(file_path)
    
    def create_schedule_entry(self, csv_row: CSVRow) -> ScheduleEntry:
        """Create ScheduleEntry from CSVRow"""
        return ScheduleEntry(
            course_name=f"{csv_row.code} - {csv_row.course_name} - {csv_row.activity_type}" or "",  # Use empty string if None
            location=csv_row.location,
            instructor=csv_row.main_tutor_write or "",  # Use empty string if None
            assistant=csv_row.helping_stuff_write or "",  # Use empty string if None
            is_half_slot=csv_row.is_half_slot or False  # Handle half slot field
        )
    
    def create_empty_day_schedule(self) -> DaySchedule:
        """Create empty day schedule with all slots as None"""
        return DaySchedule(
            first=None,
            second=None,
            third=None,
            fourth=None
        )
    
    def group_rows_by_speciality_level(self, csv_rows: List[CSVRow]) -> Dict[tuple[str, str], List[CSVRow]]:
        """Group CSV rows by specialty and level combination"""
        grouped_rows = {}
        
        for csv_row in csv_rows:
            # Use speciality and level, with fallbacks for missing values
            speciality = csv_row.speciality or csv_row.specialy_level or "عام"
            level = csv_row.get_level()
            
            key = (speciality, level)
            if key not in grouped_rows:
                grouped_rows[key] = []
            grouped_rows[key].append(csv_row)
        
        return grouped_rows
    
    def convert_to_weekly_schedule(self, csv_rows: List[CSVRow]) -> WeeklySchedule:
        """Convert CSV rows to structured weekly schedule"""
        # Initialize empty weekly schedule
        weekly_schedule = WeeklySchedule(
            sunday=self.create_empty_day_schedule(),
            monday=self.create_empty_day_schedule(),
            tuesday=self.create_empty_day_schedule(),
            wednesday=self.create_empty_day_schedule(),
            thursday=self.create_empty_day_schedule()
        )
        
        # Process each CSV row
        for csv_row in csv_rows:
            day_key = self.day_mapping.get(csv_row.day)
            slot_key = self.slot_mapping.get(csv_row.slot)
            
            if day_key and slot_key:
                schedule_entry = self.create_schedule_entry(csv_row)
                weekly_schedule[day_key][slot_key] = schedule_entry
        
        return weekly_schedule
    
    def convert_to_multi_level_schedule(self, csv_rows: List[CSVRow]) -> MultiLevelSchedule:
        """Convert CSV rows to multi-level schedule with separate schedules for each specialty-level combination"""
        # Group rows by specialty and level
        grouped_rows = self.group_rows_by_speciality_level(csv_rows)
        
        schedules = []
        
        for (speciality, level), rows in grouped_rows.items():
            print(f"📋 Processing {speciality} - {level}: {len(rows)} entries")
            
            # Convert rows for this specialty-level combination to weekly schedule
            weekly_schedule = self.convert_to_weekly_schedule(rows)
            
            # Create SpecialityLevelSchedule
            speciality_level_schedule = SpecialityLevelSchedule(
                speciality=speciality,
                level=level,
                weekly_schedule=weekly_schedule
            )
            
            schedules.append(speciality_level_schedule)
        
        # Sort schedules first by specialty, then by level
        schedules.sort(key=lambda x: (x.speciality, x.level))
        
        return MultiLevelSchedule(schedules=schedules)
    
    def convert_file_to_json(self, file_path: str) -> WeeklySchedule:
        """Main method to convert CSV/Excel file to JSON structure (backward compatibility)"""
        csv_rows = self.read_file(file_path)
        return self.convert_to_weekly_schedule(csv_rows)
    
    def convert_file_to_multi_level_json(self, file_path: str) -> MultiLevelSchedule:
        """Main method to convert CSV/Excel file to multi-level JSON structure"""
        csv_rows = self.read_file(file_path)
        return self.convert_to_multi_level_schedule(csv_rows)

    def group_rows_by_location(self, csv_rows: List[CSVRow]) -> Dict[str, List[CSVRow]]:
        """Group CSV rows by location"""
        grouped_rows = {}
        for csv_row in csv_rows:
            location = csv_row.location
            if location not in grouped_rows:
                grouped_rows[location] = []
            grouped_rows[location].append(csv_row)
        return grouped_rows

    def convert_to_multi_location_schedule(self, csv_rows: List[CSVRow]) -> MultiLocationSchedule:
        """Convert CSV rows to multi-location schedule"""
        grouped_rows = self.group_rows_by_location(csv_rows)
        schedules = []
        for location, rows in grouped_rows.items():
            print(f"📋 Processing {location}: {len(rows)} entries")
            weekly_schedule = self.convert_to_weekly_schedule(rows)
            location_schedule = LocationSchedule(
                location=location,
                weekly_schedule=weekly_schedule
            )
            schedules.append(location_schedule)
        schedules.sort(key=lambda x: x.location)
        return MultiLocationSchedule(schedules=schedules)

    def convert_file_to_multi_location_json(self, file_path: str) -> MultiLocationSchedule:
        """Main method to convert CSV/Excel file to multi-location JSON structure"""
        csv_rows = self.read_file(file_path)
        return self.convert_to_multi_location_schedule(csv_rows)
