# CSV/Excel to Word Schedule Converter

A Python application that converts CSV or Excel schedule data to a formatted Word document table, specifically designed for academic timetables with Arabic text support. **Now supports multi-level schedules with separate tables for each specialty-level combination, each on its own page with professional headers and footers.**

## Features

- **CSV/Excel to JSON Conversion**: Converts CSV or Excel data to structured JSON format
- **Multi-Level Support**: Generates separate tables for each specialty-level combination
- **Professional Page Layout**: Each table on a separate page with headers and footers
- **University Identity**: Page headers with university and department information
- **Auto-Generation Footer**: Page footers with timestamp and system information
- **Excel Support**: Reads from Excel files with 'table_full' sheet
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
│   ├── csv_converter.py   # CSV/Excel to JSON conversion logic
│   ├── word_generator.py  # Word document generation with page layout
│   └── main.py           # Main orchestration class
├── tests/
│   ├── __init__.py
│   ├── test_models.py
│   ├── test_csv_converter.py
│   ├── test_word_generator.py
│   └── test_main.py
├── requirements.txt
├── pytest.ini
├── run_conversion.py      # Multi-level conversion script
├── test_multi_level.py    # Test script for multi-level functionality
├── test_page_layout.py    # Test script for page layout functionality
├── validate_code.py       # Code validation script
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

### Multi-Level Conversion with Professional Layout (Recommended)

The application now supports multi-level schedules where each specialty-level combination gets its own page with professional formatting:

```bash
# Convert with multi-level support (generates separate pages for each specialty-level)
python run_conversion.py sample.xlsx output.docx

# Test the multi-level functionality
python test_multi_level.py

# Test the page layout functionality
python test_page_layout.py
```

### Single Table Conversion (Backward Compatibility)

```bash
# Convert to single table (legacy mode)
python -m src.main sample.xlsx output.docx
```

### Programmatic Usage

```python
from src.main import ScheduleConverter

# Create converter instance
converter = ScheduleConverter()

# Multi-level conversion with professional layout (recommended)
converter.convert_file_to_multi_level_word("input.xlsx", "output.docx")

# Get multi-level schedule structure
multi_level_schedule = converter.get_multi_level_schedule("input.xlsx")
print(f"Found {len(multi_level_schedule.schedules)} specialty-level combinations")

# Single table conversion (backward compatibility)
converter.convert_file_to_word("input.xlsx", "output.docx")
weekly_schedule = converter.get_weekly_schedule("input.xlsx")
```

## Multi-Level Schedule Structure

The application now groups data by specialty and level combinations:

### New Data Models

- **SpecialityLevelSchedule**: Represents a schedule for a specific specialty-level combination
- **MultiLevelSchedule**: Contains all specialty-level schedules in a single structure

### Professional Page Layout

The generated Word document now features:

1. **Document Title Page**: 
   - "الجداول الدراسية" (Study Schedules)
   - "كلية الهندسة - جامعة القاهرة" (Faculty of Engineering - Cairo University)
   - "العام الدراسي 2025-2026" (Academic Year 2025-2026)

2. **Separate Pages**: Each specialty-level combination gets its own page

3. **Page Headers**: Professional headers with:
   - University name: "جامعة القاهرة"
   - Faculty name: "كلية الهندسة"
   - Department: "قسم [Specialty]"
   - Academic year: "العام الدراسي 2025-2026"

4. **Page Footers**: Auto-generation information:
   - Timestamp: "تم إنشاء هذا الجدول تلقائياً في [Date] [Time]"
   - System info: "نظام إدارة الجداول الدراسية"

5. **Table Titles**: Each page has a title: "جدول [Specialty] - [Level]"

### Example Output Structure

```
[Page 1 - Title Page]
الجداول الدراسية
كلية الهندسة - جامعة القاهرة
العام الدراسي 2025-2026

[Page 2 - Table 1]
[Header: جامعة القاهرة | كلية الهندسة]
[Header: قسم هندسة الحاسبات | العام الدراسي 2025-2026]
جدول هندسة الحاسبات - المستوى الأول
[Complete schedule table]
[Footer: تم إنشاء هذا الجدول تلقائياً في 2025-01-15 14:30:25 - نظام إدارة الجداول الدراسية]

[Page 3 - Table 2]
[Header: جامعة القاهرة | كلية الهندسة]
[Header: قسم هندسة الحاسبات | العام الدراسي 2025-2026]
جدول هندسة الحاسبات - المستوى الثاني
[Complete schedule table]
[Footer: تم إنشاء هذا الجدول تلقائياً في 2025-01-15 14:30:25 - نظام إدارة الجداول الدراسية]

[Page 4 - Table 3]
[Header: جامعة القاهرة | كلية الهندسة]
[Header: قسم هندسة الإلكترونيات | العام الدراسي 2025-2026]
جدول هندسة الإلكترونيات - المستوى الأول
[Complete schedule table]
[Footer: تم إنشاء هذا الجدول تلقائياً في 2025-01-15 14:30:25 - نظام إدارة الجداول الدراسية]
```

## Input Format

The application supports both CSV and Excel files. For Excel files, the data should be in a sheet named 'table_full'.

### Required Columns

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

### Multi-Level Columns

| Column | Description | Usage |
|--------|-------------|-------|
| speciality | Specialization field | Used for grouping tables and page headers |
| level | Academic level | Used for grouping tables and page headers |
| specialy_level | Specialization level | Fallback for speciality |

### Optional Columns:
- `is_valid`: Validation status
- `active_tutor`: Active tutor information
- `confirmed by tutor`: Confirmation status
- `teaching_hours`: Teaching hours
- `teachin_hourse_printalble`: Printable teaching hours
- `sp_code`: Specialization code

## Output Format

Each table in the generated Word document contains:

- **Page Layout**: Each table on a separate page with professional headers and footers
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

# Test multi-level functionality
python test_multi_level.py

# Test page layout functionality
python test_page_layout.py

# Validate the code
python validate_code.py

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

### SpecialityLevelSchedule
Schedule for a specific specialty and level combination.

### MultiLevelSchedule
Complete multi-level schedule with all specialty-level combinations.

## Key Components

### CSVConverter
- Reads CSV and Excel files and converts to structured data
- Groups data by specialty and level combinations
- Reads from 'table_full' sheet in Excel files
- Maps Arabic day names to English keys
- Maps slot numbers to time slot names
- Handles data validation and error processing

### WordGenerator
- Creates Word documents with professional page layout
- Generates multiple tables for different specialty-level combinations
- Adds page headers with university and department identity
- Adds page footers with auto-generation information
- Places each table on a separate page
- Generates tables with merged cells
- Applies Arabic text formatting and alignment
- Sets up borders and styling

### ScheduleConverter
- Main orchestration class
- Coordinates CSV/Excel conversion and Word generation
- Supports both single-table and multi-level conversion
- Provides error handling and logging

## Error Handling

The application includes comprehensive error handling:

- CSV parsing errors
- File I/O errors
- Data validation errors
- Word document generation errors
- Multi-level grouping errors
- Page layout generation errors

## Dependencies

- `python-docx`: Word document generation
- `pydantic`: Data validation and serialization
- `pandas`: Excel file reading and data manipulation
- `openpyxl`: Excel file format support
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
