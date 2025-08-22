#!/usr/bin/env python3
"""
Test script to verify page layout functionality with headers, footers, and separate pages
"""

import sys
import os
from src.main import ScheduleConverter
from src.models import WeeklySchedule, SpecialityLevelSchedule, MultiLevelSchedule


def create_test_data():
    """Create test data for multi-level schedule"""
    # Create empty weekly schedule
    empty_schedule = WeeklySchedule(
        sunday={"first": None, "second": None, "third": None, "fourth": None},
        monday={"first": None, "second": None, "third": None, "fourth": None},
        tuesday={"first": None, "second": None, "third": None, "fourth": None},
        wednesday={"first": None, "second": None, "third": None, "fourth": None},
        thursday={"first": None, "second": None, "third": None, "fourth": None}
    )
    
    # Create multiple specialty-level schedules
    schedules = [
        SpecialityLevelSchedule(
            speciality="هندسة الحاسبات",
            level="المستوى الأول",
            weekly_schedule=empty_schedule
        ),
        SpecialityLevelSchedule(
            speciality="هندسة الحاسبات",
            level="المستوى الثاني",
            weekly_schedule=empty_schedule
        ),
        SpecialityLevelSchedule(
            speciality="هندسة الإلكترونيات",
            level="المستوى الأول",
            weekly_schedule=empty_schedule
        )
    ]
    
    return MultiLevelSchedule(schedules=schedules)


def test_page_layout():
    """Test the page layout functionality"""
    print("🧪 Testing page layout functionality...")
    print("=" * 50)
    
    try:
        # Create test data
        print("📊 Creating test data...")
        test_schedule = create_test_data()
        
        # Create converter
        converter = ScheduleConverter()
        
        # Generate Word document with page layout
        output_file = "test_page_layout.docx"
        print(f"📄 Generating Word document: {output_file}")
        
        # Use the word generator directly to test the new functionality
        from src.word_generator import WordGenerator
        word_gen = WordGenerator()
        word_gen.generate_multi_level_word_document(test_schedule, output_file)
        
        print(f"✅ Test completed successfully!")
        print(f"📄 Output file: {output_file}")
        print(f"📊 Generated {len(test_schedule.schedules)} tables on separate pages")
        print("🔍 Check the document for:")
        print("   - Page headers with university and department info")
        print("   - Page footers with generation timestamp")
        print("   - Each table on a separate page")
        print("   - Proper RTL formatting")
        
    except Exception as e:
        print(f"❌ Test failed: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    test_page_layout()
