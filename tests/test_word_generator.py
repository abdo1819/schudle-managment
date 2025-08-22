import pytest
import tempfile
import os
from src.word_generator import WordGenerator
from src.models import WeeklySchedule, ScheduleEntry, DaySchedule


class TestWordGenerator:
    """Test cases for WordGenerator class"""
    
    def setup_method(self):
        """Setup method for each test"""
        self.generator = WordGenerator()
    
    def test_initialization(self):
        """Test WordGenerator initialization"""
        assert len(self.generator.days_arabic) == 5
        assert len(self.generator.detail_categories) == 4
        assert len(self.generator.time_slots) == 4
        
        # Check Arabic days
        expected_days = ["الأحد", "الاثنين", "الثلاثاء", "الأربعاء", "الخميس"]
        assert self.generator.days_arabic == expected_days
        
        # Check detail categories
        expected_categories = [
            "اسم المادة",
            "المكان",
            "استاذ المادة",
            "الهيئة المعاونة"
        ]
        assert self.generator.detail_categories == expected_categories
        
        # Check time slots
        expected_slots = [
            "المحاضرة الأولى 8.50-10.30",
            "المحاضرة الثانية 10.40 - 12.10",
            "المحاضرة الثالثة 12.20 - 1.50",
            "المحاضرة الرابعة 2.00 - 3.30"
        ]
        assert self.generator.time_slots == expected_slots
    
    def test_create_document(self):
        """Test document creation"""
        doc = self.generator.create_document()
        
        assert doc is not None
        # Check that document has sections
        assert len(doc.sections) > 0
    
    def test_create_empty_weekly_schedule(self):
        """Test creating empty weekly schedule for testing"""
        def create_empty_schedule():
            return WeeklySchedule(
                sunday=DaySchedule(first=None, second=None, third=None, fourth=None),
                monday=DaySchedule(first=None, second=None, third=None, fourth=None),
                tuesday=DaySchedule(first=None, second=None, third=None, fourth=None),
                wednesday=DaySchedule(first=None, second=None, third=None, fourth=None),
                thursday=DaySchedule(first=None, second=None, third=None, fourth=None)
            )
        
        schedule = create_empty_schedule()
        
        assert "sunday" in schedule
        assert "monday" in schedule
        assert "tuesday" in schedule
        assert "wednesday" in schedule
        assert "thursday" in schedule
        
        for day in schedule.values():
            assert day["first"] is None
            assert day["second"] is None
            assert day["third"] is None
            assert day["fourth"] is None
    
    def test_generate_word_document_empty_schedule(self):
        """Test generating Word document with empty schedule"""
        # Create empty schedule
        schedule = WeeklySchedule(
            sunday=DaySchedule(first=None, second=None, third=None, fourth=None),
            monday=DaySchedule(first=None, second=None, third=None, fourth=None),
            tuesday=DaySchedule(first=None, second=None, third=None, fourth=None),
            wednesday=DaySchedule(first=None, second=None, third=None, fourth=None),
            thursday=DaySchedule(first=None, second=None, third=None, fourth=None)
        )
        
        # Create temporary file
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as f:
            temp_file = f.name
        
        try:
            self.generator.generate_word_document(schedule, temp_file)
            
            # Check that file was created
            assert os.path.exists(temp_file)
            assert os.path.getsize(temp_file) > 0
            
        finally:
            if os.path.exists(temp_file):
                os.unlink(temp_file)

    def test_generate_word_document_empty_schedule_manual_review(self):
        """Test generating Word document with empty schedule"""
        # Create empty schedule
        schedule = WeeklySchedule(
            sunday=DaySchedule(first=None, second=None, third=None, fourth=None),
            monday=DaySchedule(first=None, second=None, third=None, fourth=None),
            tuesday=DaySchedule(first=None, second=None, third=None, fourth=None),
            wednesday=DaySchedule(first=None, second=None, third=None, fourth=None),
            thursday=DaySchedule(first=None, second=None, third=None, fourth=None)
        )
        
        # Create temporary file
        with open("empty_schedule.docx", "wb") as f:
            self.generator.generate_word_document(schedule, f)
    
    def test_generate_word_document_with_data(self):
        """Test generating Word document with actual schedule data"""
        # Create schedule with some data
        schedule = WeeklySchedule(
            sunday=DaySchedule(first=None, second=None, third=None, fourth=None),
            monday=DaySchedule(
                first=ScheduleEntry(
                    course_name="Test Course 1",
                    location="مدرج 1",
                    instructor="د.اميرة الدسوقي",
                    assistant="م.اندرو امجد"
                ),
                second=None,
                third=None,
                fourth=None
            ),
            tuesday=DaySchedule(
                first=None,
                second=ScheduleEntry(
                    course_name="Test Course 2",
                    location="مدرج 2",
                    instructor="د.احمد محمد",
                    assistant="م.سارة علي"
                ),
                third=None,
                fourth=None
            ),
            wednesday=DaySchedule(first=None, second=None, third=None, fourth=None),
            thursday=DaySchedule(first=None, second=None, third=None, fourth=None)
        )
        
        # Create temporary file
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as f:
            temp_file = f.name
        
        try:
            self.generator.generate_word_document(schedule, temp_file)
            
            # Check that file was created
            assert os.path.exists(temp_file)
            assert os.path.getsize(temp_file) > 0
            
        finally:
            if os.path.exists(temp_file):
                os.unlink(temp_file)
    
    def test_generate_word_document_full_schedule(self):
        """Test generating Word document with full schedule data"""
        # Create comprehensive schedule
        schedule = WeeklySchedule(
            sunday=DaySchedule(
                first=ScheduleEntry(
                    course_name="نظرية الاتصالات كيت 314 (2+2) تمارين / عملي",
                    location="فصل 3",
                    instructor="أ.م.د. أحمد على نشأت إسماعيل",
                    assistant="م. عبد الرحمن اشرف سعد محمد"
                ),
                second=None,
                third=None,
                fourth=None
            ),
            monday=DaySchedule(
                first=ScheduleEntry(
                    course_name="مقرر اختياري (1) تعلم الإله كه 321 (1+2)",
                    location="فصل 1",
                    instructor="د. سيد طه محمد إبراهيم",
                    assistant="م. محمد ناصر أحمد عبد الباقي م . آلاء محمد أحمد فكيرين هلال"
                ),
                second=ScheduleEntry(
                    course_name="هندسة التحكم - كهت 305 (3+2) تمارين / عملي",
                    location="فصل 3",
                    instructor="د. احمد فرحان محمد فرحان",
                    assistant="م. محمود محمد عادل رمضان و م. احمد عويس شعبان محمد"
                ),
                third=ScheduleEntry(
                    course_name="المعالجات الدقيقة وتطبيقاتها كهج 301 (3+2)",
                    location="مدرج 4",
                    instructor="د. احمد مصطفى محمود صالح",
                    assistant="م . ندى أحمد عبد الرحمن الجمال"
                ),
                fourth=ScheduleEntry(
                    course_name="هندسة التحكم - كهت 305 (3+2)",
                    location="مدرج 3",
                    instructor="د. احمد فرحان محمد فرحان",
                    assistant="م. محمود محمد عادل رمضان و م. احمد عويس شعبان محمد"
                )
            ),
            tuesday=DaySchedule(
                first=ScheduleEntry(
                    course_name="الرسم بالحاسب كه 302 (2+2) تمارين / عملي",
                    location="فصل 2",
                    instructor="أ. د. عمر و محمد رفعت جودي",
                    assistant="م. أمنية حسني محمد السيد"
                ),
                second=ScheduleEntry(
                    course_name="الرسم بالحاسب كه 302 (2+2)",
                    location="فصل 4",
                    instructor="أ. د. عمرو محمد رفعت جودي",
                    assistant="م. محمد ناصر أحمد عبد الباقي م. أمنية حسني محمد السيد"
                ),
                third=None,
                fourth=None
            ),
            wednesday=DaySchedule(
                first=ScheduleEntry(
                    course_name="أساسيات شبكات الحاسب كيج 303 (2+2) تمارين / عملي",
                    location="فصل 2",
                    instructor="أ.د. رانيا أحمد عبد العظيم أبو السعود",
                    assistant="م. أمنية حسني محمد السيد"
                ),
                second=ScheduleEntry(
                    course_name="نظرية الاتصالات كيت 314 (2+2)",
                    location="مدرج 3",
                    instructor="أ.م.د. أحمد على نشأت إسماعيل",
                    assistant="م عبد الرحمن اشرف سعد محمد"
                ),
                third=None,
                fourth=ScheduleEntry(
                    course_name="أساسيات شبكات الحاسب كهج 303 (2+2)",
                    location="مدرج 4",
                    instructor="أ.د. رانيا أحمد عبد العظيم أبو السعود",
                    assistant="م. أمنية حسني محمد السيد"
                )
            ),
            thursday=DaySchedule(
                first=ScheduleEntry(
                    course_name="مقرر اختياري (1) تعلم الإله كه 321 (1+2) تمارين / عملي",
                    location="فصل 2",
                    instructor="د. سيد طه محمد إبراهيم",
                    assistant="م. محمد ناصر أحمد عبد الباقي م. آلاء محمد أحمد فكيرين هلال"
                ),
                second=None,
                third=ScheduleEntry(
                    course_name="المعالجات الدقيقة وتطبيقاتها كهح 301 (3+2) تمارين / عملي",
                    location="فصل 3",
                    instructor="د. احمد مصطفى محمود صالح",
                    assistant="م . ندى أحمد عبد الرحمن الجمال"
                ),
                fourth=None
            )
        )
        
        # Create temporary file
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as f:
            temp_file = f.name
        
        try:
            self.generator.generate_word_document(schedule, temp_file)
            
            # Check that file was created
            assert os.path.exists(temp_file)
            assert os.path.getsize(temp_file) > 0
            
        finally:
            if os.path.exists(temp_file):
                os.unlink(temp_file)
