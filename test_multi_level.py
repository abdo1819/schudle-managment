#!/usr/bin/env python3
"""
Test script to demonstrate multi-level schedule conversion
"""

import sys
import os
from src.main import ScheduleConverter


def test_multi_level_conversion():
    """Test the multi-level conversion functionality"""
    
    # Check if test data file exists
    test_files = [
        "time table term_1 2026.xlsx",
    ]
    
    input_file = None
    for file in test_files:
        if os.path.exists(file):
            input_file = file
            break
    
    if not input_file:
        print("❌ No test data file found. Please ensure one of these files exists:")
        for file in test_files:
            print(f"   - {file}")
        return
    
    output_file = "output_multi_level.docx"
    
    print(f"🧪 Testing multi-level conversion with: {input_file}")
    print("=" * 50)
    
    try:
        # Create converter
        converter = ScheduleConverter()
        
        # Get multi-level schedule structure
        print("📊 Analyzing data structure...")
        multi_level_schedule = converter.get_multi_level_schedule(input_file)
        
        print(f"📋 Found {len(multi_level_schedule.schedules)} specialty-level combinations:")
        for i, schedule in enumerate(multi_level_schedule.schedules, 1):
            print(f"   {i}. {schedule.speciality} - {schedule.level}")
        
        print("\n🔄 Generating Word document...")
        converter.convert_file_to_multi_level_word(input_file, output_file)
        
        print(f"\n✅ Test completed successfully!")
        print(f"📄 Output file: {output_file}")
        print(f"📊 Generated {len(multi_level_schedule.schedules)} tables")
        
    except Exception as e:
        print(f"❌ Test failed: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    test_multi_level_conversion()
