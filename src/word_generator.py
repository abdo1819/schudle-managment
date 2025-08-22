from typing import List, Dict, Any
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from .models import WeeklySchedule, ScheduleEntry, DayOfWeek, TimeSlot, DetailCategory, TableCell
from enum import Enum


class ColumnType(Enum):
    """Enum for different column types"""
    DAYS = 0
    CATEGORIES = 1
    TIME_SLOT_1 = 2
    SEPARATOR_1 = 3
    TIME_SLOT_2 = 4
    SEPARATOR_2 = 5
    TIME_SLOT_3 = 6
    SEPARATOR_3 = 7
    TIME_SLOT_4 = 8


class ColorScheme:
    """Color constants for the document"""
    DAYS_COLUMN = "8DB3E2"
    CATEGORIES_COLUMN = "B7DDE8"
    HEADER_BACKGROUND = "8DB3E2"
    SEPARATOR_BACKGROUND = "B7DDE8"


class TableDimensions:
    """Table dimension constants"""
    # Page dimensions
    A4_WIDTH_INCHES = 8.27
    MARGIN_INCHES = 0.5
    AVAILABLE_WIDTH_INCHES = A4_WIDTH_INCHES - (2 * MARGIN_INCHES)
    
    # Table structure
    TOTAL_ROWS = 21  # 5 days * 4 categories + 1 header
    TOTAL_COLUMNS = 9  # 6 data columns + 3 separators
    
    # Column widths (in inches)
    DAYS_COLUMN_WIDTH = 0.8
    CATEGORIES_COLUMN_WIDTH = 1.0
    TIME_SLOT_WIDTH = 1.2
    SEPARATOR_WIDTH = 0.2
    
    # Row structure
    HEADER_ROW_INDEX = 0
    CONTENT_START_ROW_INDEX = 1
    ROWS_PER_DAY = 4  # 4 categories per day


class WordGenerator:
    """Generates Word document with schedule table"""
    
    def __init__(self):
        self.days_arabic = ["الأحد", "الاثنين", "الثلاثاء", "الأربعاء", "الخميس"]
        self.detail_categories = [
            "اسم المادة",
            "المكان", 
            "استاذ المادة",
            "الهيئة المعاونة"
        ]
        self.time_slots = [
            "المحاضرة الأولى 8.50-10.30",
            "المحاضرة الثانية 10.40 - 12.10",
            "المحاضرة الثالثة 12.20 - 1.50",
            "المحاضرة الرابعة 2.00 - 3.30"
        ]
        
        # Define column positions for better readability
        self.time_slot_positions = [
            ColumnType.TIME_SLOT_1.value,
            ColumnType.TIME_SLOT_2.value,
            ColumnType.TIME_SLOT_3.value,
            ColumnType.TIME_SLOT_4.value
        ]
        self.separator_positions = [
            ColumnType.SEPARATOR_1.value,
            ColumnType.SEPARATOR_2.value,
            ColumnType.SEPARATOR_3.value
        ]
        self.slot_keys = ["first", "second", "third", "fourth"]
        
        # Column width mapping
        self.column_widths = {
            ColumnType.DAYS.value: TableDimensions.DAYS_COLUMN_WIDTH,
            ColumnType.CATEGORIES.value: TableDimensions.CATEGORIES_COLUMN_WIDTH,
            ColumnType.TIME_SLOT_1.value: TableDimensions.TIME_SLOT_WIDTH,
            ColumnType.SEPARATOR_1.value: TableDimensions.SEPARATOR_WIDTH,
            ColumnType.TIME_SLOT_2.value: TableDimensions.TIME_SLOT_WIDTH,
            ColumnType.SEPARATOR_2.value: TableDimensions.SEPARATOR_WIDTH,
            ColumnType.TIME_SLOT_3.value: TableDimensions.TIME_SLOT_WIDTH,
            ColumnType.SEPARATOR_3.value: TableDimensions.SEPARATOR_WIDTH,
            ColumnType.TIME_SLOT_4.value: TableDimensions.TIME_SLOT_WIDTH
        }
    
    def create_document(self) -> Document:
        """Create a new Word document"""
        doc = Document()
        
        # Set document margins
        sections = doc.sections
        for section in sections:
            section.top_margin = Inches(TableDimensions.MARGIN_INCHES)
            section.bottom_margin = Inches(TableDimensions.MARGIN_INCHES)
            section.left_margin = Inches(TableDimensions.MARGIN_INCHES)
            section.right_margin = Inches(TableDimensions.MARGIN_INCHES)
        
        return doc
    
    def create_table_structure(self, doc: Document, weekly_schedule: WeeklySchedule) -> None:
        """Create the main table structure"""
        # Create table with defined dimensions
        table = doc.add_table(rows=TableDimensions.TOTAL_ROWS, cols=TableDimensions.TOTAL_COLUMNS)
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Set table width and disable autofit
        table.autofit = False
        table.allow_autofit = False
        
        # Calculate total width from column definitions
        total_width_inches = sum(self.column_widths.values())
        table.width = Inches(total_width_inches)
        
        # Set individual column widths using XML properties
        for i, column in enumerate(table.columns):
            # Get width for this column
            width = Inches(self.column_widths[i])
            
            # Apply width to all cells in the column
            for cell in column.cells:
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                
                # Remove any existing width setting
                for child in tcPr:
                    if child.tag.endswith('tcW'):
                        tcPr.remove(child)
                
                # Add width property
                tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{int(width.inches * 1440)}" w:type="dxa"/>')
                tcPr.append(tcW)
        
        self._fill_header_row(table)
        self._fill_content_rows(table, weekly_schedule)
        self._apply_formatting(table)
    
    def _fill_header_row(self, table) -> None:
        """Fill the header row with time slots"""
        header_row = table.rows[TableDimensions.HEADER_ROW_INDEX]
        
        # First two cells are empty in header
        header_row.cells[ColumnType.DAYS.value].text = ""
        header_row.cells[ColumnType.CATEGORIES.value].text = ""
        
        # Fill each time slot in separate columns with separators
        for i, time_slot in enumerate(self.time_slots):
            header_row.cells[self.time_slot_positions[i]].text = time_slot
            header_row.cells[self.time_slot_positions[i]].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Set separator columns to empty
        for pos in self.separator_positions:
            header_row.cells[pos].text = ""
    
    def _fill_content_rows(self, table, weekly_schedule: WeeklySchedule) -> None:
        """Fill the content rows with schedule data"""
        row_index = TableDimensions.CONTENT_START_ROW_INDEX
        
        for day_index, day_arabic in enumerate(self.days_arabic):
            day_key = list(weekly_schedule.keys())[day_index]
            day_schedule = weekly_schedule[day_key]
            
            # For each day, create 4 rows (one for each detail category)
            for category_index, category in enumerate(self.detail_categories):
                row = table.rows[row_index]
                
                # Day column (merged vertically for each day)
                if category_index == 0:
                    row.cells[ColumnType.DAYS.value].text = day_arabic
                    row.cells[ColumnType.DAYS.value].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    # Merge vertically with next 3 rows
                    if row_index + 3 < len(table.rows):
                        row.cells[ColumnType.DAYS.value].merge(table.rows[row_index + 3].cells[ColumnType.DAYS.value])
                
                # Category column
                row.cells[ColumnType.CATEGORIES.value].text = category
                row.cells[ColumnType.CATEGORIES.value].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
                
                # Time slot columns with separators
                for slot_index, slot_key in enumerate(self.slot_keys):
                    cell = row.cells[self.time_slot_positions[slot_index]]
                    schedule_entry = day_schedule[slot_key]
                    
                    if schedule_entry:
                        if category_index == 0:  # Course name
                            cell.text = schedule_entry.course_name
                        elif category_index == 1:  # Location
                            cell.text = schedule_entry.location
                        elif category_index == 2:  # Instructor
                            cell.text = schedule_entry.instructor
                        elif category_index == 3:  # Assistant
                            cell.text = schedule_entry.assistant
                    else:
                        cell.text = ""
                
                # Set separator columns to empty
                for pos in self.separator_positions:
                    row.cells[pos].text = ""
                
                row_index += 1
    
    def _apply_formatting(self, table) -> None:
        """Apply formatting to the table"""
        # Apply borders and formatting to all cells
        for row_index, row in enumerate(table.rows):
            for col_index, cell in enumerate(row.cells):
                # Set font
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = 'Arial'
                        run.font.size = Pt(10)
                
                # Set cell borders
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}>'
                                    f'<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                                    f'<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                                    f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                                    f'<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                                    f'</w:tcBorders>')
                tcPr.append(tcBorders)
                
                # Apply background colors based on column type and row position
                self._apply_cell_background_color(tcPr, col_index, row_index)
    
    def _apply_cell_background_color(self, tcPr, col_index: int, row_index: int) -> None:
        """Apply background color to a table cell based on its position"""
        # Day names column (except header row)
        if col_index == ColumnType.DAYS.value and row_index > TableDimensions.HEADER_ROW_INDEX:
            self._apply_background_color(tcPr, ColorScheme.DAYS_COLUMN)
        
        # Detail categories column (except header row)
        elif col_index == ColumnType.CATEGORIES.value and row_index > TableDimensions.HEADER_ROW_INDEX:
            self._apply_background_color(tcPr, ColorScheme.CATEGORIES_COLUMN)
        
        # Time slots in header row
        elif row_index == TableDimensions.HEADER_ROW_INDEX and col_index in self.time_slot_positions:
            self._apply_background_color(tcPr, ColorScheme.HEADER_BACKGROUND)
        
        # Separator columns in header row
        elif row_index == TableDimensions.HEADER_ROW_INDEX and col_index in self.separator_positions:
            self._apply_background_color(tcPr, ColorScheme.HEADER_BACKGROUND)
        
        # Separator columns in content rows
        elif row_index > TableDimensions.HEADER_ROW_INDEX and col_index in self.separator_positions:
            self._apply_background_color(tcPr, ColorScheme.SEPARATOR_BACKGROUND)
    
    def _apply_background_color(self, tcPr, color_hex: str) -> None:
        """Apply background color to a table cell"""
        # Remove any existing shading
        for child in tcPr:
            if child.tag.endswith('shd'):
                tcPr.remove(child)
        
        # Add shading with the specified color
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
        tcPr.append(shd)
    
    def generate_word_document(self, weekly_schedule: WeeklySchedule, output_path: str) -> None:
        """Generate Word document from weekly schedule"""
        doc = self.create_document()
        self.create_table_structure(doc, weekly_schedule)
        doc.save(output_path)
