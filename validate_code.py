#!/usr/bin/env python3
"""
Validation script to check that all new multi-level functionality works correctly
"""

def validate_imports():
    """Validate that all imports work correctly"""
    try:
        print("🔍 Testing imports...")
        
        # Test basic imports
        from src.models import (
            WeeklySchedule, ScheduleEntry, DaySchedule, 
            SpecialityLevelSchedule, MultiLevelSchedule
        )
        print("✅ Models imported successfully")
        
        # Test converter imports
        from src.csv_converter import CSVConverter
        print("✅ CSVConverter imported successfully")
        
        # Test word generator imports
        from src.word_generator import WordGenerator
        print("✅ WordGenerator imported successfully")
        
        # Test main converter imports
        from src.main import ScheduleConverter
        print("✅ ScheduleConverter imported successfully")
        
        return True
        
    except Exception as e:
        print(f"❌ Import error: {e}")
        return False


def validate_models():
    """Validate that new models work correctly"""
    try:
        print("\n🔍 Testing models...")
        
        from src.models import SpecialityLevelSchedule, MultiLevelSchedule, WeeklySchedule
        
        # Test SpecialityLevelSchedule creation
        empty_schedule = WeeklySchedule(
            sunday={"first": None, "second": None, "third": None, "fourth": None},
            monday={"first": None, "second": None, "third": None, "fourth": None},
            tuesday={"first": None, "second": None, "third": None, "fourth": None},
            wednesday={"first": None, "second": None, "third": None, "fourth": None},
            thursday={"first": None, "second": None, "third": None, "fourth": None}
        )
        
        speciality_schedule = SpecialityLevelSchedule(
            speciality="هندسة الحاسبات",
            level="المستوى الأول",
            weekly_schedule=empty_schedule
        )
        print("✅ SpecialityLevelSchedule created successfully")
        
        # Test MultiLevelSchedule creation
        multi_level = MultiLevelSchedule(schedules=[speciality_schedule])
        print("✅ MultiLevelSchedule created successfully")
        
        # Test methods
        levels = multi_level.get_speciality_levels()
        print(f"✅ get_speciality_levels() returned: {levels}")
        
        schedule = multi_level.get_schedule_by_speciality_level("هندسة الحاسبات", "المستوى الأول")
        print("✅ get_schedule_by_speciality_level() returned schedule")
        
        return True
        
    except Exception as e:
        print(f"❌ Model error: {e}")
        return False


def validate_converter():
    """Validate that converter works correctly"""
    try:
        print("\n🔍 Testing converter...")
        
        from src.csv_converter import CSVConverter
        
        converter = CSVConverter()
        print("✅ CSVConverter created successfully")
        
        # Test grouping method with empty data
        from src.models import CSVRow
        
        # Create a sample CSV row
        sample_row = CSVRow(
            day="الأحد",
            slot=1,
            code="TEST-001",
            speciality="هندسة الحاسبات",
            level="المستوى الأول",
            activity_type="محاضرة",
            location="مدرج 1",
            course_name="Test Course",
            day_slot="الأحد 1",
            time="المحاضرة الأولى 8.50-10.30",
            day_order=1,
            main_tutor="د. أحمد",
            helping_stuff="م. محمد"
        )
        
        grouped = converter.group_rows_by_speciality_level([sample_row])
        print(f"✅ group_rows_by_speciality_level() returned {len(grouped)} groups")
        
        return True
        
    except Exception as e:
        print(f"❌ Converter error: {e}")
        return False


def main():
    """Main validation function"""
    print("🧪 Validating multi-level functionality...")
    print("=" * 50)
    
    success = True
    
    # Test imports
    if not validate_imports():
        success = False
    
    # Test models
    if not validate_models():
        success = False
    
    # Test converter
    if not validate_converter():
        success = False
    
    print("\n" + "=" * 50)
    if success:
        print("✅ All validations passed! Multi-level functionality is ready.")
    else:
        print("❌ Some validations failed. Please check the errors above.")
    
    return success


if __name__ == "__main__":
    main()
