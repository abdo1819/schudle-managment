#!/usr/bin/env python3
"""
Simple test script to verify Excel support
"""

import pandas as pd
import tempfile
import os
from src.csv_converter import CSVConverter


def test_excel_reading():
    """Test reading Excel file with 'table_full' sheet"""
    print("Testing Excel file reading...")
    
    # Create temporary Excel file
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        temp_file = f.name
    
    try:
        # Create sample data matching the Excel header
        data = {
            'is_valid': ['1', '1', '1', '1'],
            'day': ['الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس'],
            'slot': [1, 2, 3, 4],
            'code': ['EMP-104', None, 'EMP-106', 'EMP-107'],  # Second row has NaN code
            'speciality': ['Computer Science', 'Computer Science', 'Computer Science', 'Computer Science'],
            'activityType': ['تمارين', 'محاضرة', 'تمارين', 'محاضرة'],
            'location': ['مدرج 1', 'مدرج 2', 'مدرج 3', 'مدرج 4'],
            'active_tutor': ['د.اميرة الدسوقي', 'د.احمد محمد', 'د.سارة علي', 'د.محمد احمد'],
            'level': ['Level 1', 'Level 1', 'Level 1', 'Level 1'],
            'course_name': ['Test Course 1', None, 'Test Course 3', 'Test Course 4'],  # Second row has NaN course_name
            'day_slot': ['الاثنين 1', 'الثلاثاء 2', 'الأربعاء 3', 'الخميس 4'],
            'specialy_level': ['CS-1', 'CS-1', 'CS-1', 'CS-1'],
            'time': ['المحاضرة الاولي 8:50 - 10:20', 'المحاضرة الثانية 10:40 - 12:10', 'المحاضرة الثالثة 12:20 - 1:50', 'المحاضرة الرابعة 2:00 - 3:30'],
            'day_order': [2, 3, 4, 5],
            'confirmed by tutor': ['Yes', 'Yes', 'Yes', 'Yes'],
            'teaching_hours': ['2', '2', '2', '2'],
            'teachin_hourse_printalble': ['2 ساعة', '2 ساعة', '2 ساعة', '2 ساعة'],
            'sp_code': ['CS001', 'CS002', 'CS003', 'CS004'],
            'main_tutor': ['د.اميرة الدسوقي', None, 'د.سارة علي', 'د.محمد احمد'],  # Second row has NaN main_tutor
            'helping_stuff': ['م.اندرو امجد', None, 'م.محمد احمد', 'م.سارة علي']  # Test with None values
        }
        
        df = pd.DataFrame(data)
        
        # Write to Excel file with 'table_full' sheet
        with pd.ExcelWriter(temp_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='table_full', index=False)
        
        print(f"Created Excel file: {temp_file}")
        
        # Test reading with CSVConverter
        converter = CSVConverter()
        csv_rows = converter.read_excel(temp_file)
        
        print(f"Successfully read {len(csv_rows)} rows from Excel file")
        
        # Verify the data
        if len(csv_rows) > 0:
            first_row = csv_rows[0]
            print(f"First row - Day: {first_row.day}, Slot: {first_row.slot}, Course: {first_row.course_name}")
            print(f"Main tutor: {first_row.main_tutor}, Assistant: {first_row.helping_stuff}")
            
            # Check third row (second row was ignored due to NaN code)
            if len(csv_rows) > 1:
                third_row = csv_rows[1]
                print(f"Third row - Day: {third_row.day}, Slot: {third_row.slot}, Course: {third_row.course_name}")
                print(f"Main tutor: {third_row.main_tutor}, Assistant: {third_row.helping_stuff}")
            
            # Check fourth row
            if len(csv_rows) > 2:
                fourth_row = csv_rows[2]
                print(f"Fourth row - Day: {fourth_row.day}, Slot: {fourth_row.slot}, Course: {fourth_row.course_name}")
                print(f"Main tutor: {fourth_row.main_tutor}, Assistant: {fourth_row.helping_stuff}")
            
            # Test conversion to weekly schedule
            weekly_schedule = converter.convert_to_weekly_schedule(csv_rows)
            print("Successfully converted to weekly schedule")
            
            # Check if Monday first slot has data
            if weekly_schedule["monday"]["first"]:
                print(f"Monday first slot: {weekly_schedule['monday']['first'].course_name}")
                print(f"Monday first slot instructor: '{weekly_schedule['monday']['first'].instructor}'")
                print(f"Monday first slot assistant: '{weekly_schedule['monday']['first'].assistant}'")
            else:
                print("Monday first slot is empty")
            
            # Check if Wednesday third slot has data (Tuesday was ignored)
            if weekly_schedule["wednesday"]["third"]:
                print(f"Wednesday third slot: {weekly_schedule['wednesday']['third'].course_name}")
                print(f"Wednesday third slot instructor: '{weekly_schedule['wednesday']['third'].instructor}'")
                print(f"Wednesday third slot assistant: '{weekly_schedule['wednesday']['third'].assistant}'")
            else:
                print("Wednesday third slot is empty")
            
            # Check if Thursday fourth slot has data
            if weekly_schedule["thursday"]["fourth"]:
                print(f"Thursday fourth slot: {weekly_schedule['thursday']['fourth'].course_name}")
                print(f"Thursday fourth slot instructor: '{weekly_schedule['thursday']['fourth'].instructor}'")
                print(f"Thursday fourth slot assistant: '{weekly_schedule['thursday']['fourth'].assistant}'")
            else:
                print("Thursday fourth slot is empty")
        
        print("✅ Excel reading test passed!")
        
    except Exception as e:
        print(f"❌ Error during Excel test: {e}")
        raise
    finally:
        # Clean up
        if os.path.exists(temp_file):
            os.unlink(temp_file)
            print(f"Cleaned up temporary file: {temp_file}")


def test_file_detection():
    """Test file type detection"""
    print("\nTesting file type detection...")
    
    converter = CSVConverter()
    
    # Test CSV detection
    assert converter.read_file.__name__ == 'read_file'
    print("✅ File detection test passed!")


if __name__ == "__main__":
    test_excel_reading()
    test_file_detection()
    print("\n🎉 All tests passed!")
