# CSV to Word Schedule Converter

A Python application that converts CSV schedule data to a formatted Word document table, specifically designed for academic timetables with Arabic text support.

## Features

- **CSV to JSON Conversion**: Converts CSV data to structured JSON format
- **Word Document Generation**: Creates formatted Word documents with tables
- **Arabic Text Support**: Full support for Arabic text and right-to-left alignment
- **Type Safety**: Comprehensive type hints and Pydantic models
- **Test-Driven Development**: Complete unit test coverage
- **Modular Design**: Well-organized, maintainable code structure

## Project Structure

```
word_auto/
├── src/
│   ├── __init__.py
│   ├── models.py          # Pydantic models and data structures
│   ├── csv_converter.py   # CSV to JSON conversion logic
│   ├── word_generator.py  # Word document generation
│   └── main.py           # Main orchestration class
├── tests/
│   ├── __init__.py
│   ├── test_models.py
│   ├── test_csv_converter.py
│   ├── test_word_generator.py
│   └── test_main.py
├── requirements.txt
├── pytest.ini
├── run_conversion.py
├── README.md
└── sample.csv
```

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd word_auto
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Command Line Usage

```bash
# Convert CSV to Word document
python run_conversion.py sample.csv output.docx

# Or use the module directly
python -m src.main sample.csv output.docx
```

### Programmatic Usage

```python
from src.main import ScheduleConverter

# Create converter instance
converter = ScheduleConverter()

# Convert CSV to Word
converter.convert_csv_to_word("input.csv", "output.docx")

# Get JSON structure without generating Word document
weekly_schedule = converter.get_weekly_schedule("input.csv")
```

## CSV Format

The CSV file should have the following columns:

| Column | Description | Example |
|--------|-------------|---------|
| day | Day of the week in Arabic | الأحد, الاثنين, etc. |
| slot | Time slot number (1-4) | 1, 2, 3, 4 |
| code | Course code | EMP-104 |
| activityType | Type of activity | تمارين, محاضرة |
| location | Location/room | مدرج 1, فصل 3 |
| course_name | Course name | Differential Equation and Numerical Analysis |
| day_slot | Day and slot combination | الاثنين 1 |
| time | Time description | المحاضرة الاولي 8:50 - 10:20 |
| day_order | Day order number | 2 |
| main_tutor | Main instructor | د.اميرة الدسوقي |
| helping_stuff | Assistant staff | م.اندرو امجد |

### Optional Columns:
- `confirmed by tutor`: Confirmation status
- `teaching_hours`: Teaching hours
- `teachin_hourse_printalble`: Printable teaching hours

## Output Format

The generated Word document contains a table with the following structure:

- **Columns**: Days of the week, detail categories, and 4 time slots
- **Rows**: 21 rows (5 days × 4 categories + 1 header)
- **Time Slots**:
  - المحاضرة الأولى 8.50-10.30
  - المحاضرة الثانية 10.40 - 12.10
  - المحاضرة الثالثة 12.20 - 1.50
  - المحاضرة الرابعة 2.00 - 3.30

- **Detail Categories**:
  - اسم المادة (Course Name)
  - المكان (Location)
  - استاذ المادة (Instructor)
  - الهيئة المعاونة (Assistant Staff)

## Testing

Run the test suite:

```bash
# Run all tests
pytest

# Run with coverage
pytest --cov=src

# Run specific test file
pytest tests/test_models.py

# Run specific test class
pytest tests/test_csv_converter.py::TestCSVConverter
```

## Data Models

### CSVRow
Represents a single row from the CSV file with all required and optional fields.

### ScheduleEntry
Represents a single schedule entry with course information.

### WeeklySchedule
Complete weekly schedule structure with all days and time slots.

### DaySchedule
Schedule for a single day with all time slots (first, second, third, fourth).

## Key Components

### CSVConverter
- Reads CSV files and converts to structured data
- Maps Arabic day names to English keys
- Maps slot numbers to time slot names
- Handles data validation and error processing

### WordGenerator
- Creates Word documents with proper formatting
- Generates tables with merged cells
- Applies Arabic text formatting and alignment
- Sets up borders and styling

### ScheduleConverter
- Main orchestration class
- Coordinates CSV conversion and Word generation
- Provides error handling and logging

## Error Handling

The application includes comprehensive error handling:

- CSV parsing errors
- File I/O errors
- Data validation errors
- Word document generation errors

## Dependencies

- `python-docx`: Word document generation
- `pydantic`: Data validation and serialization
- `pytest`: Testing framework
- `pytest-cov`: Test coverage

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Run the test suite
6. Submit a pull request

## License

This project is licensed under the MIT License.
