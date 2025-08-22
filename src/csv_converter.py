import csv
from typing import List, Dict, Any
from .models import CSVRow, ScheduleEntry, DaySchedule, WeeklySchedule, DayOfWeek


class CSVConverter:
    """Converts CSV data to structured JSON format"""
    
    def __init__(self):
        self.day_mapping = {
            "الأحد": "sunday",
            "الاثنين": "monday", 
            "الثلاثاء": "tuesday",
            "الأربعاء": "wednesday",
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
        with open(file_path, 'r', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            for row in reader:
                try:
                    csv_row = CSVRow(**row)
                    rows.append(csv_row)
                except Exception as e:
                    print(f"Error parsing row: {row}, Error: {e}")
        return rows
    
    def create_schedule_entry(self, csv_row: CSVRow) -> ScheduleEntry:
        """Create ScheduleEntry from CSVRow"""
        return ScheduleEntry(
            course_name=csv_row.course_name,
            location=csv_row.location,
            instructor=csv_row.main_tutor,
            assistant=csv_row.helping_stuff
        )
    
    def create_empty_day_schedule(self) -> DaySchedule:
        """Create empty day schedule with all slots as None"""
        return DaySchedule(
            first=None,
            second=None,
            third=None,
            fourth=None
        )
    
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
    
    def convert_csv_to_json(self, file_path: str) -> WeeklySchedule:
        """Main method to convert CSV file to JSON structure"""
        csv_rows = self.read_csv(file_path)
        return self.convert_to_weekly_schedule(csv_rows)
