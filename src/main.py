from typing import Optional
from .csv_converter import CSVConverter
from .word_generator import WordGenerator
from .models import WeeklySchedule


class ScheduleConverter:
    """Main class to orchestrate CSV to Word conversion"""
    
    def __init__(self):
        self.csv_converter = CSVConverter()
        self.word_generator = WordGenerator()
    
    def convert_csv_to_word(self, csv_path: str, output_path: str) -> None:
        """Convert CSV file to Word document"""
        try:
            # Step 1: Convert CSV to JSON structure
            print(f"Converting CSV file: {csv_path}")
            weekly_schedule = self.csv_converter.convert_csv_to_json(csv_path)
            
            # Step 2: Generate Word document
            print(f"Generating Word document: {output_path}")
            self.word_generator.generate_word_document(weekly_schedule, output_path)
            
            print("Conversion completed successfully!")
            
        except Exception as e:
            print(f"Error during conversion: {e}")
            raise
    
    def get_weekly_schedule(self, csv_path: str) -> WeeklySchedule:
        """Get weekly schedule from CSV without generating Word document"""
        return self.csv_converter.convert_csv_to_json(csv_path)


def main():
    """Main function for command line usage"""
    import sys
    
    if len(sys.argv) != 3:
        print("Usage: python -m src.main <input_csv_path> <output_docx_path>")
        sys.exit(1)
    
    csv_path = sys.argv[1]
    output_path = sys.argv[2]
    
    converter = ScheduleConverter()
    converter.convert_csv_to_word(csv_path, output_path)


if __name__ == "__main__":
    main()
