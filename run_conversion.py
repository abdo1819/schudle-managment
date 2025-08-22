#!/usr/bin/env python3
"""
Simple script to run CSV to Word conversion with multi-level support
"""

import sys
import os
from src.main import ScheduleConverter


def main():
    """Main function to run the conversion"""
    if len(sys.argv) != 3:
        print("Usage: python run_conversion.py <input_file_path> <output_docx_path>")
        print("Supported formats: CSV (.csv), Excel (.xlsx, .xls)")
        print("This will generate separate tables for each specialty-level combination.")
        print("Example: python run_conversion.py sample.xlsx output.docx")
        print("Example: python run_conversion.py sample.csv output.docx")
        sys.exit(1)
    
    file_path = sys.argv[1]
    output_path = sys.argv[2]
    
    # Check if input file exists
    if not os.path.exists(file_path):
        print(f"Error: Input file '{file_path}' does not exist.")
        sys.exit(1)
    
    # Check if file format is supported
    supported_extensions = ['.csv', '.xlsx', '.xls']
    file_extension = os.path.splitext(file_path)[1].lower()
    if file_extension not in supported_extensions:
        print(f"Error: Unsupported file format '{file_extension}'. Supported formats: {', '.join(supported_extensions)}")
        sys.exit(1)
    
    try:
        # Create converter and run multi-level conversion
        converter = ScheduleConverter()
        converter.convert_file_to_multi_level_word(file_path, output_path)
        
        print(f"✅ Multi-level conversion completed successfully!")
        print(f"📄 Input file: {file_path}")
        print(f"📄 Output Word: {output_path}")
        print("📋 Generated separate tables for each specialty-level combination")
        
    except Exception as e:
        print(f"❌ Error during conversion: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
