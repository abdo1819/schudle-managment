#!/usr/bin/env python3
"""
Test script for half slot functionality
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from src.models import CSVRow, ScheduleEntry, WeeklySchedule
from src.csv_converter import CSVConverter
from src.word_generator import WordGenerator

def test_half_slot_functionality():
    """Test the half slot functionality"""
    print("🧪 Testing half slot functionality...")
    
    # Test 1: Create CSVRow with half slot
    csv_row_half = CSVRow(
        day="الاثنين",
        slot=1,
        is_half_slot=True,
        code="EMP-104",
        activity_type="تمارين",
        location="مدرج 1",
        course_name="Half Slot Course",
        day_slot="الاثنين 1",
        time="المحاضرة الاولي 8:50 - 10:20",
        day_order=2,
        main_tutor="د.اميرة الدسوقي",
        helping_stuff="م.اندرو امجد"
    )
    
    # Test 2: Create CSVRow without half slot
    csv_row_full = CSVRow(
        day="الاثنين",
        slot=2,
        is_half_slot=False,
        code="EMP-105",
        activity_type="محاضرة",
        location="مدرج 2",
        course_name="Full Slot Course",
        day_slot="الاثنين 2",
        time="المحاضرة الثانية 10:40 - 12:10",
        day_order=2,
        main_tutor="د.احمد محمد",
        helping_stuff="م.سارة علي"
    )
    
    print(f"✅ CSVRow half slot: {csv_row_half.is_half_slot}")
    print(f"✅ CSVRow full slot: {csv_row_full.is_half_slot}")
    
    # Test 3: Convert to ScheduleEntry
    converter = CSVConverter()
    entry_half = converter.create_schedule_entry(csv_row_half)
    entry_full = converter.create_schedule_entry(csv_row_full)
    
    print(f"✅ ScheduleEntry half slot: {entry_half.is_half_slot}")
    print(f"✅ ScheduleEntry full slot: {entry_full.is_half_slot}")
    
    # Test 4: Create weekly schedule with both types
    weekly_schedule = WeeklySchedule(
        sunday=converter.create_empty_day_schedule(),
        monday=converter.create_empty_day_schedule(),
        tuesday=converter.create_empty_day_schedule(),
        wednesday=converter.create_empty_day_schedule(),
        thursday=converter.create_empty_day_schedule()
    )
    
    # Add entries to Monday
    weekly_schedule["monday"]["first"] = entry_half
    weekly_schedule["monday"]["second"] = entry_full
    
    print("✅ Weekly schedule created with half and full slots")
    
    # Test 5: Generate Word document
    generator = WordGenerator()
    output_path = "test_half_slot_output.docx"
    
    try:
        generator.generate_word_document(weekly_schedule, output_path)
        print(f"✅ Word document generated: {output_path}")
        
        # Check if file exists
        if os.path.exists(output_path):
            print(f"✅ File size: {os.path.getsize(output_path)} bytes")
        else:
            print("❌ File was not created")
            
    except Exception as e:
        print(f"❌ Error generating Word document: {e}")
    
    print("🎉 Half slot functionality test completed!")

if __name__ == "__main__":
    test_half_slot_functionality()
