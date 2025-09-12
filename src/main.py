from typing import Optional
from .csv_converter import CSVConverter
from .word_generator import WordGenerator, LocationWordGenerator
from .models import WeeklySchedule, MultiLevelSchedule, MultiLocationSchedule


class ScheduleConverter:
    """Main class to orchestrate CSV/Excel to Word conversion"""
    
    def __init__(self):
        self.csv_converter = CSVConverter()
        self.word_generator = WordGenerator()
        self.location_word_generator = LocationWordGenerator()
    
    def convert_file_to_word(self, file_path: str, output_path: str) -> None:
        """Convert CSV/Excel file to Word document (single table - backward compatibility)"""
        try:
            # Step 1: Convert file to JSON structure
            print(f"Converting file: {file_path}")
            weekly_schedule = self.csv_converter.convert_file_to_json(file_path)
            
            # Step 2: Generate Word document
            print(f"Generating Word document: {output_path}")
            self.word_generator.generate_word_document(weekly_schedule, output_path)
            
            print("Conversion completed successfully!")
            
        except Exception as e:
            print(f"Error during conversion: {e}")
            raise
    
    def convert_file_to_multi_level_word(self, file_path: str, output_path: str) -> None:
        """Convert CSV/Excel file to Word document with multiple tables for each specialty-level combination"""
        try:
            # Step 1: Convert file to multi-level JSON structure
            print(f"Converting file: {file_path}")
            multi_level_schedule = self.csv_converter.convert_file_to_multi_level_json(file_path)
            
            # Step 2: Generate Word document with multiple tables
            print(f"Generating multi-level Word document: {output_path}")
            self.word_generator.generate_multi_level_word_document(multi_level_schedule, output_path)
            
            print("Multi-level conversion completed successfully!")
            print(f"Generated {len(multi_level_schedule.schedules)} tables for different specialty-level combinations")
            
        except Exception as e:
            print(f"Error during multi-level conversion: {e}")
            raise

    def convert_file_to_multi_location_word(self, file_path: str, output_path: str) -> None:
        """Convert CSV/Excel file to Word document with multiple tables for each location"""
        try:
            # Step 1: Convert file to multi-location JSON structure
            print(f"Converting file: {file_path}")
            multi_location_schedule = self.csv_converter.convert_file_to_multi_location_json(file_path)
            
            # Step 2: Generate Word document with multiple tables
            print(f"Generating multi-location Word document: {output_path}")
            self.location_word_generator.generate_multi_location_word_document(multi_location_schedule, output_path)
            
            print("Multi-location conversion completed successfully!")
            print(f"Generated {len(multi_location_schedule.schedules)} tables for different locations")
            
        except Exception as e:
            print(f"Error during multi-location conversion: {e}")
            raise
    
    def get_weekly_schedule(self, file_path: str) -> WeeklySchedule:
        """Get weekly schedule from file without generating Word document (backward compatibility)"""
        return self.csv_converter.convert_file_to_json(file_path)
    
    def get_multi_level_schedule(self, file_path: str) -> MultiLevelSchedule:
        """Get multi-level schedule from file without generating Word document"""
        return self.csv_converter.convert_file_to_multi_level_json(file_path)

    def get_multi_location_schedule(self, file_path: str) -> MultiLocationSchedule:
        """Get multi-location schedule from file without generating Word document"""
        return self.csv_converter.convert_file_to_multi_location_json(file_path)
    
    # Backward compatibility method
    def convert_csv_to_word(self, csv_path: str, output_path: str) -> None:
        """Convert CSV file to Word document (backward compatibility)"""
        return self.convert_file_to_word(csv_path, output_path)


def main():
    """Main function for command line usage"""
    import sys
    
    if len(sys.argv) != 3:
        print("Usage: python -m src.main <input_file_path> <output_docx_path>")
        print("Supported formats: CSV (.csv), Excel (.xlsx, .xls)")
        print("Note: This will generate a single table. Use run_conversion.py for multi-level tables.")
        sys.exit(1)
    
    file_path = sys.argv[1]
    output_path = sys.argv[2]
    
    converter = ScheduleConverter()
    converter.convert_file_to_word(file_path, output_path)


if __name__ == "__main__":
    main()
