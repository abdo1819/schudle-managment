#!/usr/bin/env python3
"""
Simple script to run CSV to Word conversion with multi-level support
"""

import sys
import os
import argparse
from src.main import ScheduleConverter


def main():
    """Main function to run the conversion"""
    parser = argparse.ArgumentParser(description="Convert CSV/Excel to Word schedule document.")
    parser.add_argument("input_file", help="Path to the input CSV or Excel file.")
    parser.add_argument("output_file", help="Path to the output DOCX file.")
    parser.add_argument("--view", choices=["level", "location"], default="level",
                        help="Type of view to generate: 'level' (default) or 'location'.")

    args = parser.parse_args()

    # Check if input file exists
    if not os.path.exists(args.input_file):
        print(f"Error: Input file '{args.input_file}' does not exist.")
        sys.exit(1)

    # Check if file format is supported
    supported_extensions = ['.csv', '.xlsx', '.xls']
    file_extension = os.path.splitext(args.input_file)[1].lower()
    if file_extension not in supported_extensions:
        print(f"Error: Unsupported file format '{file_extension}'. Supported formats: {', '.join(supported_extensions)}")
        sys.exit(1)

    try:
        converter = ScheduleConverter()

        if args.view == "level":
            converter.convert_file_to_multi_level_word(args.input_file, args.output_file)
            print(f"✅ Multi-level conversion completed successfully!")
            print(f"📄 Input file: {args.input_file}")
            print(f"📄 Output Word: {args.output_file}")
            print("📋 Generated separate tables for each specialty-level combination")
        elif args.view == "location":
            converter.convert_file_to_multi_location_word(args.input_file, args.output_file)
            print(f"✅ Multi-location conversion completed successfully!")
            print(f"📄 Input file: {args.input_file}")
            print(f"📄 Output Word: {args.output_file}")
            print("📋 Generated separate tables for each location")

    except Exception as e:
        print(f"❌ Error during conversion: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
