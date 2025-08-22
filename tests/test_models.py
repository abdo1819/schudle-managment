import pytest
from src.models import (
    CSVRow, ScheduleEntry, DaySchedule, WeeklySchedule,
    DayOfWeek, TimeSlot, DetailCategory, TableCell
)


class TestCSVRow:
    """Test cases for CSVRow model"""
    
    def test_csv_row_creation(self):
        """Test creating CSVRow with valid data"""
        data = {
            "day": "الاثنين",
            "slot": 1,
            "code": "EMP-104",
            "activityType": "تمارين",
            "location": "مدرج 1",
            "course_name": "Differential Equation and Numerical Analysis",
            "day_slot": "الاثنين 1",
            "time": "المحاضرة الاولي 8:50 - 10:20",
            "day_order": 2,
            "main_tutor": "د.اميرة الدسوقي",
            "helping_stuff": "م.اندرو امجد"
        }
        
        csv_row = CSVRow(**data)
        
        assert csv_row.day == "الاثنين"
        assert csv_row.slot == 1
        assert csv_row.code == "EMP-104"
        assert csv_row.activity_type == "تمارين"
        assert csv_row.location == "مدرج 1"
        assert csv_row.course_name == "Differential Equation and Numerical Analysis"
        assert csv_row.main_tutor == "د.اميرة الدسوقي"
        assert csv_row.helping_stuff == "م.اندرو امجد"
    
    def test_csv_row_with_optional_fields(self):
        """Test creating CSVRow with optional fields"""
        data = {
            "day": "الاثنين",
            "slot": 1,
            "code": "EMP-104",
            "activityType": "تمارين",
            "location": "مدرج 1",
            "course_name": "Test Course",
            "day_slot": "الاثنين 1",
            "time": "المحاضرة الاولي 8:50 - 10:20",
            "day_order": 2,
            "confirmed by tutor": "Yes",
            "teaching_hours": "3",
            "teachin_hourse_printalble": "(3)",
            "main_tutor": "د.اميرة الدسوقي",
            "helping_stuff": "م.اندرو امجد"
        }
        
        csv_row = CSVRow(**data)
        
        assert csv_row.confirmed_by_tutor == "Yes"
        assert csv_row.teaching_hours == "3"
        assert csv_row.teaching_hours_printable == "(3)"


class TestScheduleEntry:
    """Test cases for ScheduleEntry model"""
    
    def test_schedule_entry_creation(self):
        """Test creating ScheduleEntry"""
        entry = ScheduleEntry(
            course_name="Test Course",
            location="مدرج 1",
            instructor="د.اميرة الدسوقي",
            assistant="م.اندرو امجد"
        )
        
        assert entry.course_name == "Test Course"
        assert entry.location == "مدرج 1"
        assert entry.instructor == "د.اميرة الدسوقي"
        assert entry.assistant == "م.اندرو امجد"


class TestEnums:
    """Test cases for enum classes"""
    
    def test_day_of_week_enum(self):
        """Test DayOfWeek enum values"""
        assert DayOfWeek.SUNDAY == "الأحد"
        assert DayOfWeek.MONDAY == "الاثنين"
        assert DayOfWeek.TUESDAY == "الثلاثاء"
        assert DayOfWeek.WEDNESDAY == "الأربعاء"
        assert DayOfWeek.THURSDAY == "الخميس"
    
    def test_time_slot_enum(self):
        """Test TimeSlot enum values"""
        assert TimeSlot.FIRST == "المحاضرة الأولى 8.50-10.30"
        assert TimeSlot.SECOND == "المحاضرة الثانية 10.40 - 12.10"
        assert TimeSlot.THIRD == "المحاضرة الثالثة 12.20 - 1.50"
        assert TimeSlot.FOURTH == "المحاضرة الرابعة 2.00 - 3.30"
    
    def test_detail_category_enum(self):
        """Test DetailCategory enum values"""
        assert DetailCategory.COURSE_NAME == "اسم المادة"
        assert DetailCategory.LOCATION == "المكان"
        assert DetailCategory.INSTRUCTOR == "استاذ المادة"
        assert DetailCategory.ASSISTANT == "الهيئة المعاونة"


class TestTableCell:
    """Test cases for TableCell model"""
    
    def test_table_cell_creation(self):
        """Test creating TableCell"""
        cell = TableCell(
            content="Test Content",
            is_merged=True,
            merge_span=2,
            alignment="center"
        )
        
        assert cell.content == "Test Content"
        assert cell.is_merged is True
        assert cell.merge_span == 2
        assert cell.alignment == "center"
    
    def test_table_cell_defaults(self):
        """Test TableCell with default values"""
        cell = TableCell(content="Test")
        
        assert cell.content == "Test"
        assert cell.is_merged is False
        assert cell.merge_span == 1
        assert cell.alignment == "left"
