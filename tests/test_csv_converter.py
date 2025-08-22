import pytest
import tempfile
import os
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
            weekly_schedule = self.converter.convert_csv_to_json(temp_file)
            
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
