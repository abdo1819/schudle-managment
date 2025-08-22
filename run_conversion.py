#!/usr/bin/env python3
"""
Simple script to run CSV to Word conversion
"""

import sys
import os
from src.main import ScheduleConverter


def main():
    """Main function to run the conversion"""
    if len(sys.argv) != 3:
        print("Usage: python run_conversion.py <input_csv_path> <output_docx_path>")
        print("Example: python run_conversion.py sample.csv output.docx")
        sys.exit(1)
    
    csv_path = sys.argv[1]
    output_path = sys.argv[2]
    
    # Check if input file exists
    if not os.path.exists(csv_path):
        print(f"Error: Input file '{csv_path}' does not exist.")
        sys.exit(1)
    
    try:
        # Create converter and run conversion
        converter = ScheduleConverter()
        converter.convert_csv_to_word(csv_path, output_path)
        
        print(f"✅ Conversion completed successfully!")
        print(f"📄 Input CSV: {csv_path}")
        print(f"📄 Output Word: {output_path}")
        
    except Exception as e:
        print(f"❌ Error during conversion: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
