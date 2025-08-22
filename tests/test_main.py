import pytest
import tempfile
import os
from src.main import ScheduleConverter
from src.models import WeeklySchedule, ScheduleEntry, DaySchedule


class TestScheduleConverter:
    """Test cases for ScheduleConverter class"""
    
    def setup_method(self):
        """Setup method for each test"""
        self.converter = ScheduleConverter()
    
    def test_initialization(self):
        """Test ScheduleConverter initialization"""
        assert self.converter.csv_converter is not None
        assert self.converter.word_generator is not None
    
    def test_get_weekly_schedule(self):
        """Test getting weekly schedule from CSV"""
        # Create temporary CSV file
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as f:
            f.write("day,slot,code,activityType,location,course_name,day_slot,time,day_order,main_tutor,helping_stuff\n")
            f.write("الاثنين,1,EMP-104,تمارين,مدرج 1,Test Course 1,الاثنين 1,المحاضرة الاولي 8:50 - 10:20,2,د.اميرة الدسوقي,م.اندرو امجد\n")
            f.write("الثلاثاء,2,EMP-105,محاضرة,مدرج 2,Test Course 2,الثلاثاء 2,المحاضرة الثانية 10:40 - 12:10,3,د.احمد محمد,م.سارة علي\n")
            temp_file = f.name
        
        try:
            weekly_schedule = self.converter.get_weekly_schedule(temp_file)
            
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
            
        finally:
            os.unlink(temp_file)
    
    def test_convert_csv_to_word_integration(self):
        """Test full integration of CSV to Word conversion"""
        # Create temporary CSV file
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as f:
            f.write("day,slot,code,activityType,location,course_name,day_slot,time,day_order,main_tutor,helping_stuff\n")
            f.write("الاثنين,1,EMP-104,تمارين,مدرج 1,Test Course 1,الاثنين 1,المحاضرة الاولي 8:50 - 10:20,2,د.اميرة الدسوقي,م.اندرو امجد\n")
            f.write("الثلاثاء,2,EMP-105,محاضرة,مدرج 2,Test Course 2,الثلاثاء 2,المحاضرة الثانية 10:40 - 12:10,3,د.احمد محمد,م.سارة علي\n")
            csv_file = f.name
        
        # Create temporary output file
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as f:
            docx_file = f.name
        
        try:
            # Perform conversion
            self.converter.convert_csv_to_word(csv_file, docx_file)
            
            # Verify output file was created
            assert os.path.exists(docx_file)
            assert os.path.getsize(docx_file) > 0
            
        finally:
            # Clean up
            if os.path.exists(csv_file):
                os.unlink(csv_file)
            if os.path.exists(docx_file):
                os.unlink(docx_file)
    
    def test_convert_csv_to_word_with_empty_csv(self):
        """Test conversion with empty CSV file"""
        # Create temporary empty CSV file
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as f:
            f.write("day,slot,code,activityType,location,course_name,day_slot,time,day_order,main_tutor,helping_stuff\n")
            csv_file = f.name
        
        # Create temporary output file
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as f:
            docx_file = f.name
        
        try:
            # Perform conversion
            self.converter.convert_csv_to_word(csv_file, docx_file)
            
            # Verify output file was created (should create empty table)
            assert os.path.exists(docx_file)
            assert os.path.getsize(docx_file) > 0
            
        finally:
            # Clean up
            if os.path.exists(csv_file):
                os.unlink(csv_file)
            if os.path.exists(docx_file):
                os.unlink(docx_file)
    
    def test_convert_csv_to_word_with_full_schedule(self):
        """Test conversion with comprehensive schedule data"""
        # Create temporary CSV file with full schedule
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as f:
            f.write("day,slot,code,activityType,location,course_name,day_slot,time,day_order,main_tutor,helping_stuff\n")
            f.write("الأحد,1,EMP-101,تمارين,فصل 3,نظرية الاتصالات كيت 314 (2+2) تمارين / عملي,الأحد 1,المحاضرة الاولي 8:50 - 10:20,1,أ.م.د. أحمد على نشأت إسماعيل,م. عبد الرحمن اشرف سعد محمد\n")
            f.write("الاثنين,1,EMP-102,محاضرة,فصل 1,مقرر اختياري (1) تعلم الإله كه 321 (1+2),الاثنين 1,المحاضرة الاولي 8:50 - 10:20,2,د. سيد طه محمد إبراهيم,م. محمد ناصر أحمد عبد الباقي م . آلاء محمد أحمد فكيرين هلال\n")
            f.write("الاثنين,2,EMP-103,تمارين,فصل 3,هندسة التحكم - كهت 305 (3+2) تمارين / عملي,الاثنين 2,المحاضرة الثانية 10:40 - 12:10,2,د. احمد فرحان محمد فرحان,م. محمود محمد عادل رمضان و م. احمد عويس شعبان محمد\n")
            f.write("الاثنين,3,EMP-104,محاضرة,مدرج 4,المعالجات الدقيقة وتطبيقاتها كهج 301 (3+2),الاثنين 3,المحاضرة الثالثة 12:20 - 1:50,2,د. احمد مصطفى محمود صالح,م . ندى أحمد عبد الرحمن الجمال\n")
            f.write("الاثنين,4,EMP-105,محاضرة,مدرج 3,هندسة التحكم - كهت 305 (3+2),الاثنين 4,المحاضرة الرابعة 2:00 - 3:30,2,د. احمد فرحان محمد فرحان,م. محمود محمد عادل رمضان و م. احمد عويس شعبان محمد\n")
            csv_file = f.name
        
        # Create temporary output file
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as f:
            docx_file = f.name
        
        try:
            # Perform conversion
            self.converter.convert_csv_to_word(csv_file, docx_file)
            
            # Verify output file was created
            assert os.path.exists(docx_file)
            assert os.path.getsize(docx_file) > 0
            
        finally:
            # Clean up
            if os.path.exists(csv_file):
                os.unlink(csv_file)
            if os.path.exists(docx_file):
                os.unlink(docx_file)
