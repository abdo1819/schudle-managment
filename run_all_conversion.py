#!/usr/bin/env python3
"""
Test script to demonstrate multi-level schedule conversion
"""

import sys
import os
from datetime import datetime
from src.main import ScheduleConverter
from src.document_converter import convert_to_pdf_and_open


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
    
    # Create timestamped output folder and filename
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_folder = f"output_{timestamp}"
    
    # Create the output folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"📁 Created output folder: {output_folder}")
    
    # --- File Paths ---
    output_level_file = os.path.join(output_folder, f"output_multi_level_{timestamp}.docx")
    output_location_file = os.path.join(output_folder, f"output_multi_location_{timestamp}.docx")
    output_main_tutor_file = os.path.join(output_folder, f"output_main_tutor_load_{timestamp}.docx")
    output_helping_stuff_file = os.path.join(output_folder, f"output_helping_staff_load_{timestamp}.docx")
    
    print(f"🧪 Testing multi-level conversion with: {input_file}")
    print("=" * 50)
    
    try:
        # Create converter
        converter = ScheduleConverter()
        
        # --- Level View ---
        print("\n🔄 Generating Word document for level view...")
        converter.convert_file_to_multi_level_word(input_file, output_level_file)
        
        # --- Location View ---
        print("\n🔄 Generating Word document for location view...")
        converter.convert_file_to_multi_location_word(input_file, output_location_file)

        # --- Staff Load View ---
        print("\n🔄 Generating Word document for main tutor load view...")
        converter.convert_file_to_multi_staff_word(input_file, output_main_tutor_file, 'main_tutor_write')

        print("\n🔄 Generating Word document for helping stuff load view...")
        converter.convert_file_to_multi_staff_word(input_file, output_helping_stuff_file, 'helping_stuff_write')
        
        print(f"\n✅ Test completed successfully!")
        print(f"📁 Output folder: {output_folder}")
        print(f"📄 Output file (level view): {output_level_file}")
        print(f"📄 Output file (location view): {output_location_file}")
        print(f"📄 Output file (main tutor load): {output_main_tutor_file}")
        print(f"📄 Output file (helping stuff load): {output_helping_stuff_file}")

        # Convert to PDF and open
        convert_to_pdf_and_open(output_level_file)
        convert_to_pdf_and_open(output_location_file)
        convert_to_pdf_and_open(output_main_tutor_file)
        convert_to_pdf_and_open(output_helping_stuff_file)
        
    except Exception as e:
        print(f"❌ Test failed: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    test_multi_level_conversion()