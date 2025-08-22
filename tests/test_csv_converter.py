import pytest
import tempfile
import os
import pandas as pd
from src.csv_converter import CSVConverter
from src.models import CSVRow, ScheduleEntry, WeeklySchedule


class TestCSVConverter:
    """Test cases for CSVConverter class"""
    
    def setup_method(self):
        """Setup method for each test"""
        self.converter = CSVConverter()
    
    def test_day_mapping(self):
        """Test day mapping dictionary"""
        expected_mapping = {
            "الأحد": "sunday",
            "الاثنين": "monday",
            "الثلاثاء": "tuesday",
            "الأربعاء": "wednesday",
            "الخميس": "thursday"
        }
        
        assert self.converter.day_mapping == expected_mapping
    
    def test_slot_mapping(self):
        """Test slot mapping dictionary"""
        expected_mapping = {
            1: "first",
            2: "second",
            3: "third",
            4: "fourth"
        }
        
        assert self.converter.slot_mapping == expected_mapping
    
    def test_create_schedule_entry(self):
        """Test creating ScheduleEntry from CSVRow"""
        csv_row = CSVRow(
            day="الاثنين",
            slot=1,
            code="EMP-104",
            activity_type="تمارين",
            location="مدرج 1",
            course_name="Test Course",
            day_slot="الاثنين 1",
            time="المحاضرة الاولي 8:50 - 10:20",
            day_order=2,
            main_tutor="د.اميرة الدسوقي",
            helping_stuff="م.اندرو امجد"
        )
        
        entry = self.converter.create_schedule_entry(csv_row)
        
        assert isinstance(entry, ScheduleEntry)
        assert entry.course_name == "Test Course"
        assert entry.location == "مدرج 1"
        assert entry.instructor == "د.اميرة الدسوقي"
        assert entry.assistant == "م.اندرو امجد"
    
    def test_create_schedule_entry_with_none_helping_stuff(self):
        """Test creating ScheduleEntry from CSVRow with None helping_stuff"""
        csv_row = CSVRow(
            day="الاثنين",
            slot=1,
            code="EMP-104",
            activity_type="تمارين",
            location="مدرج 1",
            course_name="Test Course",
            day_slot="الاثنين 1",
            time="المحاضرة الاولي 8:50 - 10:20",
            day_order=2,
            main_tutor="د.اميرة الدسوقي",
            helping_stuff=None
        )
        
        entry = self.converter.create_schedule_entry(csv_row)
        
        assert isinstance(entry, ScheduleEntry)
        assert entry.course_name == "Test Course"
        assert entry.location == "مدرج 1"
        assert entry.instructor == "د.اميرة الدسوقي"
        assert entry.assistant == ""  # Should be empty string when None
    
    def test_create_schedule_entry_with_none_course_name_and_tutor(self):
        """Test creating ScheduleEntry from CSVRow with None course_name and main_tutor"""
        csv_row = CSVRow(
            day="الاثنين",
            slot=1,
            code="EMP-104",
            activity_type="تمارين",
            location="مدرج 1",
            course_name=None,
            day_slot="الاثنين 1",
            time="المحاضرة الاولي 8:50 - 10:20",
            day_order=2,
            main_tutor=None,
            helping_stuff="م.اندرو امجد"
        )
        
        entry = self.converter.create_schedule_entry(csv_row)
        
        assert isinstance(entry, ScheduleEntry)
        assert entry.course_name == ""  # Should be empty string when None
        assert entry.location == "مدرج 1"
        assert entry.instructor == ""  # Should be empty string when None
        assert entry.assistant == "م.اندرو امجد"
    
    def test_create_empty_day_schedule(self):
        """Test creating empty day schedule"""
        day_schedule = self.converter.create_empty_day_schedule()
        
        assert day_schedule["first"] is None
        assert day_schedule["second"] is None
        assert day_schedule["third"] is None
        assert day_schedule["fourth"] is None
    
    def test_convert_to_weekly_schedule(self):
        """Test converting CSV rows to weekly schedule"""
        csv_rows = [
            CSVRow(
                day="الاثنين",
                slot=1,
                code="EMP-104",
                activity_type="تمارين",
                location="مدرج 1",
                course_name="Test Course 1",
                day_slot="الاثنين 1",
                time="المحاضرة الاولي 8:50 - 10:20",
                day_order=2,
                main_tutor="د.اميرة الدسوقي",
                helping_stuff="م.اندرو امجد"
            ),
            CSVRow(
                day="الثلاثاء",
                slot=2,
                code="EMP-105",
                activity_type="محاضرة",
                location="مدرج 2",
                course_name="Test Course 2",
                day_slot="الثلاثاء 2",
                time="المحاضرة الثانية 10:40 - 12:10",
                day_order=3,
                main_tutor="د.احمد محمد",
                helping_stuff="م.سارة علي"
            )
        ]
        
        weekly_schedule = self.converter.convert_to_weekly_schedule(csv_rows)
        
        # Check structure
        assert "sunday" in weekly_schedule
        assert "monday" in weekly_schedule
        assert "tuesday" in weekly_schedule
        assert "wednesday" in weekly_schedule
        assert "thursday" in weekly_schedule
        
        # Check Monday first slot
        assert weekly_schedule["monday"]["first"] is not None
        assert weekly_schedule["monday"]["first"].course_name == "Test Course 1"
        assert weekly_schedule["monday"]["first"].location == "مدرج 1"
        
        # Check Tuesday second slot
        assert weekly_schedule["tuesday"]["second"] is not None
        assert weekly_schedule["tuesday"]["second"].course_name == "Test Course 2"
        assert weekly_schedule["tuesday"]["second"].location == "مدرج 2"
        
        # Check empty slots
        assert weekly_schedule["sunday"]["first"] is None
        assert weekly_schedule["monday"]["second"] is None
    
    def test_read_csv_file(self):
        """Test reading CSV file"""
        # Create temporary CSV file
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as f:
            f.write("day,slot,code,activityType,location,course_name,day_slot,time,day_order,main_tutor,helping_stuff\n")
            f.write("الاثنين,1,EMP-104,تمارين,مدرج 1,Test Course,الاثنين 1,المحاضرة الاولي 8:50 - 10:20,2,د.اميرة الدسوقي,م.اندرو امجد\n")
            temp_file = f.name
        
        try:
            csv_rows = self.converter.read_csv(temp_file)
            
            assert len(csv_rows) == 1
            assert csv_rows[0].day == "الاثنين"
            assert csv_rows[0].slot == 1
            assert csv_rows[0].course_name == "Test Course"
            
        finally:
            os.unlink(temp_file)
    
    def test_convert_csv_to_json_integration(self):
        """Test full integration of CSV to JSON conversion"""
        # Create temporary CSV file
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as f:
            f.write("day,slot,code,activityType,location,course_name,day_slot,time,day_order,main_tutor,helping_stuff\n")
            f.write("الاثنين,1,EMP-104,تمارين,مدرج 1,Test Course 1,الاثنين 1,المحاضرة الاولي 8:50 - 10:20,2,د.اميرة الدسوقي,م.اندرو امجد\n")
            f.write("الثلاثاء,2,EMP-105,محاضرة,مدرج 2,Test Course 2,الثلاثاء 2,المحاضرة الثانية 10:40 - 12:10,3,د.احمد محمد,م.سارة علي\n")
            temp_file = f.name
        
        try:
            weekly_schedule = self.converter.convert_file_to_json(temp_file)
            
            # Verify structure
            assert isinstance(weekly_schedule, dict)
            assert "monday" in weekly_schedule
            assert "tuesday" in weekly_schedule
            
            # Verify Monday data
            assert weekly_schedule["monday"]["first"] is not None
            assert weekly_schedule["monday"]["first"].course_name == "Test Course 1"
            
            # Verify Tuesday data
            assert weekly_schedule["tuesday"]["second"] is not None
            assert weekly_schedule["tuesday"]["second"].course_name == "Test Course 2"
            
            # Verify empty slots
            assert weekly_schedule["sunday"]["first"] is None
            assert weekly_schedule["monday"]["second"] is None
            
        finally:
            os.unlink(temp_file)
    
    def test_read_excel_file(self):
        """Test reading Excel file from 'table_full' sheet"""
        # Create temporary Excel file with 'table_full' sheet
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
            temp_file = f.name
        
        try:
            # Create DataFrame with the expected structure
            data = {
                'is_valid': ['1'],
                'day': ['الاثنين'],
                'slot': [1],
                'code': ['EMP-104'],
                'speciality': ['Computer Science'],
                'activityType': ['تمارين'],
                'location': ['مدرج 1'],
                'active_tutor': ['د.اميرة الدسوقي'],
                'level': ['Level 1'],
                'course_name': ['Test Course'],
                'day_slot': ['الاثنين 1'],
                'specialy_level': ['CS-1'],
                'time': ['المحاضرة الاولي 8:50 - 10:20'],
                'day_order': [2],
                'confirmed by tutor': ['Yes'],
                'teaching_hours': ['2'],
                'teachin_hourse_printalble': ['2 ساعة'],
                'sp_code': ['CS001'],
                'main_tutor': ['د.اميرة الدسوقي'],
                'helping_stuff': ['م.اندرو امجد']
            }
            
            df = pd.DataFrame(data)
            
            # Write to Excel file with 'table_full' sheet
            with pd.ExcelWriter(temp_file, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='table_full', index=False)
            
            csv_rows = self.converter.read_excel(temp_file)
            
            assert len(csv_rows) == 1
            assert csv_rows[0].day == "الاثنين"
            assert csv_rows[0].slot == 1
            assert csv_rows[0].course_name == "Test Course"
            assert csv_rows[0].main_tutor == "د.اميرة الدسوقي"
            assert csv_rows[0].helping_stuff == "م.اندرو امجد"
            
        finally:
            os.unlink(temp_file)
    
    def test_read_file_csv(self):
        """Test read_file method with CSV file"""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as f:
            f.write("day,slot,code,activityType,location,course_name,day_slot,time,day_order,main_tutor,helping_stuff\n")
            f.write("الاثنين,1,EMP-104,تمارين,مدرج 1,Test Course,الاثنين 1,المحاضرة الاولي 8:50 - 10:20,2,د.اميرة الدسوقي,م.اندرو امجد\n")
            temp_file = f.name
        
        try:
            csv_rows = self.converter.read_file(temp_file)
            
            assert len(csv_rows) == 1
            assert csv_rows[0].day == "الاثنين"
            assert csv_rows[0].slot == 1
            assert csv_rows[0].course_name == "Test Course"
            
        finally:
            os.unlink(temp_file)
    
    def test_read_file_excel(self):
        """Test read_file method with Excel file"""
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
            temp_file = f.name
        
        try:
            # Create DataFrame with the expected structure
            data = {
                'day': ['الاثنين'],
                'slot': [1],
                'code': ['EMP-104'],
                'activityType': ['تمارين'],
                'location': ['مدرج 1'],
                'course_name': ['Test Course'],
                'day_slot': ['الاثنين 1'],
                'time': ['المحاضرة الاولي 8:50 - 10:20'],
                'day_order': [2],
                'main_tutor': ['د.اميرة الدسوقي'],
                'helping_stuff': ['م.اندرو امجد']
            }
            
            df = pd.DataFrame(data)
            
            # Write to Excel file with 'table_full' sheet
            with pd.ExcelWriter(temp_file, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='table_full', index=False)
            
            csv_rows = self.converter.read_file(temp_file)
            
            assert len(csv_rows) == 1
            assert csv_rows[0].day == "الاثنين"
            assert csv_rows[0].slot == 1
            assert csv_rows[0].course_name == "Test Course"
            
        finally:
            os.unlink(temp_file)
    
    def test_read_excel_with_nan_values(self):
        """Test reading Excel file with NaN values in helping_stuff"""
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
            temp_file = f.name
        
        try:
            # Create DataFrame with NaN values
            data = {
                'day': ['الاثنين'],
                'slot': [1],
                'code': ['EMP-104'],
                'activityType': ['تمارين'],
                'location': ['مدرج 1'],
                'course_name': ['Test Course'],
                'day_slot': ['الاثنين 1'],
                'time': ['المحاضرة الاولي 8:50 - 10:20'],
                'day_order': [2],
                'main_tutor': ['د.اميرة الدسوقي'],
                'helping_stuff': [None]  # This will become NaN in Excel
            }
            
            df = pd.DataFrame(data)
            
            # Write to Excel file with 'table_full' sheet
            with pd.ExcelWriter(temp_file, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='table_full', index=False)
            
            csv_rows = self.converter.read_excel(temp_file)
            
            assert len(csv_rows) == 1
            assert csv_rows[0].day == "الاثنين"
            assert csv_rows[0].slot == 1
            assert csv_rows[0].course_name == "Test Course"
            assert csv_rows[0].helping_stuff is None  # Should be None, not NaN
            
        finally:
            os.unlink(temp_file)
    
    def test_read_excel_ignore_nan_code_rows(self):
        """Test reading Excel file and ignoring rows with NaN codes"""
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
            temp_file = f.name
        
        try:
            # Create DataFrame with some rows having NaN codes
            data = {
                'day': ['الاثنين', 'الثلاثاء', 'الأربعاء'],
                'slot': [1, 2, 3],
                'code': ['EMP-104', None, 'EMP-106'],  # Second row has NaN code
                'activityType': ['تمارين', 'محاضرة', 'تمارين'],
                'location': ['مدرج 1', 'مدرج 2', 'مدرج 3'],
                'course_name': ['Test Course 1', 'Test Course 2', 'Test Course 3'],
                'day_slot': ['الاثنين 1', 'الثلاثاء 2', 'الأربعاء 3'],
                'time': ['المحاضرة الاولي 8:50 - 10:20', 'المحاضرة الثانية 10:40 - 12:10', 'المحاضرة الثالثة 12:20 - 1:50'],
                'day_order': [2, 3, 4],
                'main_tutor': ['د.اميرة الدسوقي', 'د.احمد محمد', 'د.سارة علي'],
                'helping_stuff': ['م.اندرو امجد', 'م.سارة علي', 'م.محمد احمد']
            }
            
            df = pd.DataFrame(data)
            
            # Write to Excel file with 'table_full' sheet
            with pd.ExcelWriter(temp_file, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='table_full', index=False)
            
            csv_rows = self.converter.read_excel(temp_file)
            
            # Should only have 2 rows (first and third), second row should be ignored
            assert len(csv_rows) == 2
            assert csv_rows[0].day == "الاثنين"
            assert csv_rows[0].code == "EMP-104"
            assert csv_rows[1].day == "الأربعاء"
            assert csv_rows[1].code == "EMP-106"
            
        finally:
            os.unlink(temp_file)
    
    def test_convert_file_to_json_excel_integration(self):
        """Test full integration of Excel to JSON conversion"""
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
            temp_file = f.name
        
        try:
            # Create DataFrame with the expected structure
            data = {
                'day': ['الاثنين', 'الثلاثاء'],
                'slot': [1, 2],
                'code': ['EMP-104', 'EMP-105'],
                'activityType': ['تمارين', 'محاضرة'],
                'location': ['مدرج 1', 'مدرج 2'],
                'course_name': ['Test Course 1', 'Test Course 2'],
                'day_slot': ['الاثنين 1', 'الثلاثاء 2'],
                'time': ['المحاضرة الاولي 8:50 - 10:20', 'المحاضرة الثانية 10:40 - 12:10'],
                'day_order': [2, 3],
                'main_tutor': ['د.اميرة الدسوقي', 'د.احمد محمد'],
                'helping_stuff': ['م.اندرو امجد', 'م.سارة علي']
            }
            
            df = pd.DataFrame(data)
            
            # Write to Excel file with 'table_full' sheet
            with pd.ExcelWriter(temp_file, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='table_full', index=False)
            
            weekly_schedule = self.converter.convert_file_to_json(temp_file)
            
            # Verify structure
            assert isinstance(weekly_schedule, dict)
            assert "monday" in weekly_schedule
            assert "tuesday" in weekly_schedule
            
            # Verify Monday data
            assert weekly_schedule["monday"]["first"] is not None
            assert weekly_schedule["monday"]["first"].course_name == "Test Course 1"
            
            # Verify Tuesday data
            assert weekly_schedule["tuesday"]["second"] is not None
            assert weekly_schedule["tuesday"]["second"].course_name == "Test Course 2"
            
            # Verify empty slots
            assert weekly_schedule["sunday"]["first"] is None
            assert weekly_schedule["monday"]["second"] is None
            
        finally:
            os.unlink(temp_file)
