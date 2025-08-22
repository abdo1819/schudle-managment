from docx import Document
from typing import List, Dict, Any
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from .models import WeeklySchedule, ScheduleEntry, DayOfWeek, TimeSlot, DetailCategory, TableCell, MultiLevelSchedule, SpecialityLevelSchedule
from enum import Enum
from datetime import datetime
import os


# Header constants
HEADER_UNIVERSITY_NAME = "جامعة الفيوم"
HEADER_FACULTY_NAME = "كلية الهندسة"
HEADER_DEPARTMENT_PREFIX = "قسم الهندسة الكهربية"
HEADER_ACADEMIC_YEAR = "العام الجامعي 2025 - 2026"
HEADER_SEMESTER = "الفصل الدراسي الأول"
HEADER_DIVISION_PREFIX = "شعبة"
HEADER_LEVEL_PREFIX = "الفرقة"
HEADER_SCHEDULE_TEMPLATE = "شئون التعليم والطلاب نموذج الجداول الدراسية"

# Footer constants
FOOTER_DEAN_TITLE = "عميد الكلية"
FOOTER_DEAN_NAME = "ا.د. رانيا احمد عبدالعظيم"
FOOTER_VICE_DEAN_TITLE = "وكيل الكلية لشئون التعليم والطلاب"
FOOTER_VICE_DEAN_NAME = "ا.د. احمد سرج فريد"
FOOTER_HEAD_DEPARTMENT_TITLE = "رئيس قسم الهندسة الكهربية"
FOOTER_HEAD_DEPARTMENT_NAME = "ا.د. عمرو رفعت"
FOOTER_GENERATION_INFO = "{date}"


class ColumnType(Enum):
    """Enum for different column types"""
    DAYS = 0
    CATEGORIES = 1
    TIME_SLOT_1 = 2
    TIME_SLOT_1_HALF = 3  # Half slot column for slot 1
    SEPARATOR_1 = 4
    TIME_SLOT_2 = 5
    TIME_SLOT_2_HALF = 6  # Half slot column for slot 2
    SEPARATOR_2 = 7
    TIME_SLOT_3 = 8
    TIME_SLOT_3_HALF = 9  # Half slot column for slot 3
    SEPARATOR_3 = 10
    TIME_SLOT_4 = 11
    TIME_SLOT_4_HALF = 12  # Half slot column for slot 4


class RowType(Enum):
    """Enum for different row types"""
    HEADER = 0
    DAY_START = 1  # First row of each day (course_name)
    DAY_MIDDLE = 2  # Middle rows of each day (location, teacher)
    DAY_END = 3  # Last row of each day (assistant)


class ColorScheme:
    """Color constants for the document"""
    DAYS_COLUMN = "8DB3E2"
    CATEGORIES_COLUMN = "B7DDE8"
    HEADER_BACKGROUND = "8DB3E2"
    SEPARATOR_BACKGROUND = "B7DDE8"


class BorderWidth:
    """Border width constants in Word units (1/8 point)"""
    THIN = 4      # 0.5pt
    THICK = 18    # 2.25pt


class FontConfig:
    """Font configuration constants"""
    FONT_NAME = 'Arial'
    TITLE_SIZE = Pt(18)
    HEADER_SIZE = Pt(12)
    FOOTER_SIZE = Pt(10)
    TABLE_CELL_SIZE = Pt(8)


class TableDimensions:
    """Table dimension constants"""
    # Page dimensions
    A4_WIDTH_INCHES = 8.27
    MARGIN_INCHES = 0.5
    AVAILABLE_WIDTH_INCHES = A4_WIDTH_INCHES - (2 * MARGIN_INCHES)
    
    # Table structure
    TOTAL_ROWS = 21  # 5 days * 4 categories + 1 header
    TOTAL_COLUMNS = 13  # 6 data columns + 3 separators + 4 half slots
    
    # Column widths (in inches)
    DAYS_COLUMN_WIDTH = 0.8
    CATEGORIES_COLUMN_WIDTH = 1.0
    TIME_SLOT_WIDTH = 1.0
    TIME_SLOT_HALF_WIDTH = 0.2  # Half width for half slots
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
        self.time_slot_half_positions = [
            ColumnType.TIME_SLOT_1_HALF.value,
            ColumnType.TIME_SLOT_2_HALF.value,
            ColumnType.TIME_SLOT_3_HALF.value,
            ColumnType.TIME_SLOT_4_HALF.value
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
            ColumnType.TIME_SLOT_1_HALF.value: TableDimensions.TIME_SLOT_HALF_WIDTH,
            ColumnType.SEPARATOR_1.value: TableDimensions.SEPARATOR_WIDTH,
            ColumnType.TIME_SLOT_2.value: TableDimensions.TIME_SLOT_WIDTH,
            ColumnType.TIME_SLOT_2_HALF.value: TableDimensions.TIME_SLOT_HALF_WIDTH,
            ColumnType.SEPARATOR_2.value: TableDimensions.SEPARATOR_WIDTH,
            ColumnType.TIME_SLOT_3.value: TableDimensions.TIME_SLOT_WIDTH,
            ColumnType.TIME_SLOT_3_HALF.value: TableDimensions.TIME_SLOT_HALF_WIDTH,
            ColumnType.SEPARATOR_3.value: TableDimensions.SEPARATOR_WIDTH,
            ColumnType.TIME_SLOT_4.value: TableDimensions.TIME_SLOT_WIDTH,
            ColumnType.TIME_SLOT_4_HALF.value: TableDimensions.TIME_SLOT_HALF_WIDTH
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
    
    def add_page_header(self, doc: Document, speciality: str, level: str) -> None:
        """Add page header with identity information"""
        section = doc.sections[0]
        self._add_header_to_section(section, speciality, level)
    
    def add_page_footer(self, doc: Document) -> None:
        """Add page footer with generation information"""
        section = doc.sections[0]
        self._add_footer_to_section(section)
    
    def add_speciality_level_title(self, doc: Document, speciality: str, level: str) -> None:
        """Add title for specialty and level combination"""
        # Add spacing before title
        doc.add_paragraph()
        
        # Create title paragraph
        title_para = doc.add_paragraph()
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Add title text
        title_text = f"جدول {speciality} - {level}"
        title_run = title_para.add_run(title_text)
        title_run.font.name = FontConfig.FONT_NAME
        title_run.font.size = FontConfig.TITLE_SIZE
        title_run.font.bold = True
        
        # Set title to RTL
        self._set_paragraph_rtl(title_para)
        
        # Add spacing after title
        doc.add_paragraph()
    
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
        self._set_table_column_widths(table)
        
        # Set table to RTL
        self._set_table_rtl(table)
        
        self._fill_header_row(table)
        self._fill_content_rows(table, weekly_schedule)
        self._apply_formatting(table)
        self._apply_table_outline_borders(table)
    
    def _fill_header_row(self, table) -> None:
        """Fill the header row with time slots"""
        header_row = table.rows[TableDimensions.HEADER_ROW_INDEX]
        
        # First two cells are empty in header
        header_row.cells[ColumnType.DAYS.value].text = ""
        header_row.cells[ColumnType.CATEGORIES.value].text = ""
        
        # Fill each time slot in separate columns with separators
        for i, time_slot in enumerate(self.time_slots):
            # Main time slot column
            main_cell = header_row.cells[self.time_slot_positions[i]]
            half_cell = header_row.cells[self.time_slot_half_positions[i]]
            
            # Set content in main cell and merge with half cell
            main_cell.text = time_slot
            main_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            self._set_cell_rtl(main_cell)
            
            # Merge main and half slot columns for header
            main_cell.merge(half_cell)
            # Clean up any extra paragraphs that might cause newlines
            if len(main_cell.paragraphs) > 1:
                # Keep only the first paragraph
                for i in range(len(main_cell.paragraphs) - 1, 0, -1):
                    main_cell.paragraphs[i]._element.getparent().remove(main_cell.paragraphs[i]._element)
            # Re-apply formatting after cleanup
            main_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            self._set_cell_rtl(main_cell)
        
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
                    day_cell = row.cells[ColumnType.DAYS.value]
                    day_cell.text = day_arabic
                    day_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    self._set_cell_rtl(day_cell)
                    # Merge vertically with next 3 rows
                    if row_index + 3 < len(table.rows):
                        day_cell.merge(table.rows[row_index + 3].cells[ColumnType.DAYS.value])
                        # Reapply thick borders to merged day cell
                        self._apply_day_cell_borders(day_cell, row_index)
                
                # Category column
                category_cell = row.cells[ColumnType.CATEGORIES.value]
                category_cell.text = category
                category_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                self._set_cell_rtl(category_cell)
                
                # Time slot columns with separators
                for slot_index, slot_key in enumerate(self.slot_keys):
                    main_cell = row.cells[self.time_slot_positions[slot_index]]
                    half_cell = row.cells[self.time_slot_half_positions[slot_index]]
                    schedule_entry = day_schedule[slot_key]
                    
                    if schedule_entry:
                        # Determine content based on category
                        if category_index == 0:  # Course name
                            content = schedule_entry.course_name
                        elif category_index == 1:  # Location
                            content = schedule_entry.location
                        elif category_index == 2:  # Instructor
                            content = schedule_entry.instructor
                        elif category_index == 3:  # Assistant
                            content = schedule_entry.assistant
                        else:
                            content = ""
                        
                        # Handle half slot logic
                        if schedule_entry.is_half_slot:
                            # For half slots: put content in main column, leave half column empty
                            main_cell.text = content
                            half_cell.text = ""
                        else:
                            # For full slots: merge the cells and put content across both columns
                            main_cell.text = content
                            half_cell.text = ""
                            # Merge the cells horizontally (main + half slot columns)
                            main_cell.merge(half_cell)
                            # Clean up any extra paragraphs that might cause newlines
                            if len(main_cell.paragraphs) > 1:
                                # Keep only the first paragraph
                                for i in range(len(main_cell.paragraphs) - 1, 0, -1):
                                    main_cell.paragraphs[i]._element.getparent().remove(main_cell.paragraphs[i]._element)
                            # Re-apply formatting after cleanup
                            main_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                            self._set_cell_rtl(main_cell)
                    else:
                        # Empty slot: merge the cells and leave both empty
                        main_cell.text = ""
                        half_cell.text = ""
                        # Merge the cells horizontally (main + half slot columns)
                        main_cell.merge(half_cell)
                        # Clean up any extra paragraphs that might cause newlines
                        if len(main_cell.paragraphs) > 1:
                            # Keep only the first paragraph
                            for i in range(len(main_cell.paragraphs) - 1, 0, -1):
                                main_cell.paragraphs[i]._element.getparent().remove(main_cell.paragraphs[i]._element)
                        # Re-apply formatting after cleanup
                        main_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        self._set_cell_rtl(main_cell)
                
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
                        run.font.name = FontConfig.FONT_NAME
                        run.font.size = FontConfig.TABLE_CELL_SIZE
                
                # Set cell borders with different widths based on row and column types
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                
                # Determine row type
                if row_index == TableDimensions.HEADER_ROW_INDEX:
                    row_type = RowType.HEADER
                else:
                    # Calculate which row within the day this is (0-3)
                    day_row_index = (row_index - TableDimensions.CONTENT_START_ROW_INDEX) % TableDimensions.ROWS_PER_DAY
                    if day_row_index == 0:
                        row_type = RowType.DAY_START
                    elif day_row_index == TableDimensions.ROWS_PER_DAY - 1:
                        row_type = RowType.DAY_END
                    else:
                        row_type = RowType.DAY_MIDDLE
                
                # Ensure thick vertical borders for the day names and categories column (outer-left outline)
                vertical_border_width = str(BorderWidth.THICK)
                
                # Horizontal borders depend on row type
                if row_type == RowType.HEADER:
                    # Header row – make top border thick to complete outer outline
                    top_border_width = str(BorderWidth.THICK)
                    bottom_border_width = str(BorderWidth.THIN)
                elif row_type == RowType.DAY_START:
                    # First row of day - top border thick, bottom border thin
                    top_border_width = str(BorderWidth.THICK)
                    # If this is the merged day cell row, we also want a thick bottom
                    # border to ensure the outline of the merged cell across rows.
                    if col_index == ColumnType.DAYS.value:
                        bottom_border_width = str(BorderWidth.THICK)
                    else:
                        bottom_border_width = str(BorderWidth.THIN)
                elif row_type == RowType.DAY_END:
                    # Last row of day - top border thin, bottom border thick
                    top_border_width = str(BorderWidth.THIN)
                    bottom_border_width = str(BorderWidth.THICK)
                else:  # DAY_MIDDLE
                    # Middle rows of day - all borders thin
                    top_border_width = str(BorderWidth.THIN)
                    bottom_border_width = str(BorderWidth.THIN)
                
                tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}>'
                                    f'<w:top w:val="single" w:sz="{top_border_width}" w:space="0" w:color="000000"/>'
                                    f'<w:left w:val="single" w:sz="{vertical_border_width}" w:space="0" w:color="000000"/>'
                                    f'<w:bottom w:val="single" w:sz="{bottom_border_width}" w:space="0" w:color="000000"/>'
                                    f'<w:right w:val="single" w:sz="{vertical_border_width}" w:space="0" w:color="000000"/>'
                                    f'</w:tcBorders>')
                tcPr.append(tcBorders)
                
                # Apply background colors based on column type and row position
                self._apply_cell_background_color(tcPr, col_index, row_index)
                
                # Ensure all cells have RTL formatting and center alignment
                if cell.text.strip():  # Only apply to non-empty cells
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    self._set_cell_rtl(cell)
    
    def _apply_day_cell_borders(self, day_cell, row_index: int) -> None:
        """Apply thick borders to merged day cells after merging"""
        tc = day_cell._tc
        tcPr = tc.get_or_add_tcPr()
        
        # Remove any existing borders
        for child in tcPr:
            if child.tag.endswith('tcBorders'):
                tcPr.remove(child)
        
        # For merged cells, we need to ensure the borders apply to the entire merged region
        # Check if this is a merged cell by looking for vMerge element
        is_merged = False
        for child in tcPr:
            if child.tag.endswith('vMerge'):
                is_merged = True
                break
        
        # Apply thick borders all around the merged day cell
        # For merged cells, we need to be more explicit about the border application
        if is_merged:
            # Use more explicit border settings for merged cells
            tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}>'
                                f'<w:top w:val="single" w:sz="{BorderWidth.THICK}" w:space="0" w:color="000000"/>'
                                f'<w:left w:val="single" w:sz="{BorderWidth.THICK}" w:space="0" w:color="000000"/>'
                                f'<w:bottom w:val="single" w:sz="{BorderWidth.THICK}" w:space="0" w:color="000000"/>'
                                f'<w:right w:val="single" w:sz="{BorderWidth.THICK}" w:space="0" w:color="000000"/>'
                                f'<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                                f'<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                                f'</w:tcBorders>')
        else:
            # Standard border application for non-merged cells
            tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}>'
                                f'<w:top w:val="single" w:sz="{BorderWidth.THICK}" w:space="0" w:color="000000"/>'
                                f'<w:left w:val="single" w:sz="{BorderWidth.THICK}" w:space="0" w:color="000000"/>'
                                f'<w:bottom w:val="single" w:sz="{BorderWidth.THICK}" w:space="0" w:color="000000"/>'
                                f'<w:right w:val="single" w:sz="{BorderWidth.THICK}" w:space="0" w:color="000000"/>'
                                f'</w:tcBorders>')
        
        tcPr.append(tcBorders)
        
        # Reapply background color for day column
        self._apply_background_color(tcPr, ColorScheme.DAYS_COLUMN)
        
        # Ensure font formatting is maintained
        for paragraph in day_cell.paragraphs:
            for run in paragraph.runs:
                run.font.name = FontConfig.FONT_NAME
                run.font.size = FontConfig.TABLE_CELL_SIZE
        
        # For merged cells, also ensure the table-level borders are properly set
        if is_merged:
            self._ensure_merged_cell_table_borders(day_cell)
    
    def _ensure_merged_cell_table_borders(self, merged_cell) -> None:
        """Ensure table-level borders are properly set for merged cells"""
        # Get the table that contains this cell
        table = merged_cell._tc.getparent().getparent()
        
        # Get or create table properties
        tblPr = table.tblPr
        if tblPr is None:
            tblPr = parse_xml(f'<w:tblPr {nsdecls("w")}/>')
            table.insert(0, tblPr)
        
        # Remove any existing table borders
        for child in tblPr:
            if child.tag.endswith('tblBorders'):
                tblPr.remove(child)
        
        # Add table-level borders that ensure merged cells respect the border rules
        tblBorders = parse_xml(f'<w:tblBorders {nsdecls("w")}>'
                             f'<w:top w:val="single" w:sz="{BorderWidth.THICK}" w:space="0" w:color="000000"/>'
                             f'<w:left w:val="single" w:sz="{BorderWidth.THICK}" w:space="0" w:color="000000"/>'
                             f'<w:bottom w:val="single" w:sz="{BorderWidth.THICK}" w:space="0" w:color="000000"/>'
                             f'<w:right w:val="single" w:sz="{BorderWidth.THICK}" w:space="0" w:color="000000"/>'
                             f'<w:insideH w:val="single" w:sz="{BorderWidth.THIN}" w:space="0" w:color="000000"/>'
                             f'<w:insideV w:val="single" w:sz="{BorderWidth.THIN}" w:space="0" w:color="000000"/>'
                             f'</w:tblBorders>')
        tblPr.append(tblBorders)
    
    def _apply_cell_background_color(self, tcPr, col_index: int, row_index: int) -> None:
        """Apply background color to a table cell based on its position"""
        # Day names column (except header row)
        if col_index == ColumnType.DAYS.value and row_index > TableDimensions.HEADER_ROW_INDEX:
            self._apply_background_color(tcPr, ColorScheme.DAYS_COLUMN)
        
        # Detail categories column (except header row)
        elif col_index == ColumnType.CATEGORIES.value and row_index > TableDimensions.HEADER_ROW_INDEX:
            self._apply_background_color(tcPr, ColorScheme.CATEGORIES_COLUMN)
        
        # Time slots in header row (both main and half slot columns)
        elif row_index == TableDimensions.HEADER_ROW_INDEX and (col_index in self.time_slot_positions or col_index in self.time_slot_half_positions):
            self._apply_background_color(tcPr, ColorScheme.HEADER_BACKGROUND)
        
        # Separator columns in header row
        elif row_index == TableDimensions.HEADER_ROW_INDEX and col_index in self.separator_positions:
            self._apply_background_color(tcPr, ColorScheme.HEADER_BACKGROUND)
        
        # Separator columns in content rows
        elif row_index > TableDimensions.HEADER_ROW_INDEX and col_index in self.separator_positions:
            self._apply_background_color(tcPr, ColorScheme.SEPARATOR_BACKGROUND)
    
    def _apply_table_outline_borders(self, table) -> None:
        """Apply thick borders to the entire table outline"""
        rows = table.rows
        total_rows = len(rows)
        total_cols = len(rows[0].cells) if rows else 0
        
        # Get the Word namespace URI
        w_ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        
        for row_index, row in enumerate(rows):
            for col_index, cell in enumerate(row.cells):
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                
                # Check if this cell is on the table outline
                is_top_edge = row_index == 0
                is_bottom_edge = row_index == total_rows - 1
                is_left_edge = col_index == 0
                is_right_edge = col_index == total_cols - 1
                
                # If cell is on any edge, ensure that edge has thick border
                if is_top_edge or is_bottom_edge or is_left_edge or is_right_edge:
                    # Get existing borders or create new ones
                    existing_borders = None
                    for child in tcPr:
                        if child.tag.endswith('tcBorders'):
                            existing_borders = child
                            break
                    
                    if existing_borders is not None:
                        # Update existing borders for outline edges
                        if is_top_edge:
                            for border in existing_borders:
                                if border.tag.endswith('top'):
                                    border.set(f'{w_ns}sz', str(BorderWidth.THICK))
                        if is_bottom_edge:
                            for border in existing_borders:
                                if border.tag.endswith('bottom'):
                                    border.set(f'{w_ns}sz', str(BorderWidth.THICK))
                        if is_left_edge:
                            for border in existing_borders:
                                if border.tag.endswith('left'):
                                    border.set(f'{w_ns}sz', str(BorderWidth.THICK))
                        if is_right_edge:
                            for border in existing_borders:
                                if border.tag.endswith('right'):
                                    border.set(f'{w_ns}sz', str(BorderWidth.THICK))
    
    def _apply_background_color(self, tcPr, color_hex: str) -> None:
        """Apply background color to a table cell"""
        # Remove any existing shading
        for child in tcPr:
            if child.tag.endswith('shd'):
                tcPr.remove(child)
        
        # Add shading with the specified color
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
        tcPr.append(shd)
    
    def _set_table_column_widths(self, table) -> None:
        """Set column widths for a table"""
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
    
    def _set_table_rtl(self, table) -> None:
        """Set the entire table to RTL direction"""
        tbl = table._tbl
        
        # Get or create tblPr element
        tblPr = tbl.tblPr
        if tblPr is None:
            tblPr = parse_xml(f'<w:tblPr {nsdecls("w")}/>')
            tbl.insert(0, tblPr)
        
        # Remove any existing bidiVisual setting
        for child in tblPr:
            if child.tag.endswith('bidiVisual'):
                tblPr.remove(child)
        
        # Add RTL setting
        bidiVisual = parse_xml(f'<w:bidiVisual {nsdecls("w")}/>')
        tblPr.append(bidiVisual)
    
    def _set_cell_rtl(self, cell) -> None:
        """Set a cell to RTL direction"""
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        
        # Remove any existing textDirection setting
        for child in tcPr:
            if child.tag.endswith('textDirection'):
                tcPr.remove(child)
        
        # Add RTL text direction (horizontal RTL)
        textDirection = parse_xml(f'<w:textDirection {nsdecls("w")} w:val="rtl"/>')
        tcPr.append(textDirection)
        
        # Set paragraph direction to RTL
        for paragraph in cell.paragraphs:
            self._set_paragraph_rtl(paragraph)
    
    def _set_paragraph_rtl(self, paragraph) -> None:
        """Set a paragraph to RTL direction"""
        p = paragraph._p
        
        # Get or create pPr element
        pPr = p.pPr
        if pPr is None:
            pPr = parse_xml(f'<w:pPr {nsdecls("w")}/>')
            p.insert(0, pPr)
        
        # Remove any existing bidi setting
        for child in pPr:
            if child.tag.endswith('bidi'):
                pPr.remove(child)
        
        # Add RTL paragraph direction
        bidi = parse_xml(f'<w:bidi {nsdecls("w")}/>')
        pPr.append(bidi)
    
    def _clear_section_content(self, section_element) -> None:
        """Safely clear content from a section element (header/footer)"""
        try:
            for element in list(section_element._element):
                section_element._element.remove(element)
        except:
            # If clearing fails, continue silently
            pass
    
    def _create_header_table(self, header) -> Any:
        """Create and configure header table with proper column widths"""
        # Add header table with three columns: right, center, left
        header_table = header.add_table(rows=3, cols=3, width=Inches(TableDimensions.AVAILABLE_WIDTH_INCHES))
        header_table.style = 'Table Grid'
        header_table.autofit = False
        header_table.allow_autofit = False
        
        # Set column widths (right, center, left)
        column_widths = [2.5, 3.0, 2.5]  # in inches
        for i, column in enumerate(header_table.columns):
            for cell in column.cells:
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{int(column_widths[i] * 1440)}" w:type="dxa"/>')
                tcPr.append(tcW)
        
        return header_table
    
    def _fill_header_content(self, header_table, speciality: str, level: str) -> None:
        """Fill header table with content"""
        # Right column (University/Faculty/Department info)
        header_table.rows[0].cells[0].text = HEADER_UNIVERSITY_NAME
        header_table.rows[1].cells[0].text = HEADER_FACULTY_NAME
        header_table.rows[2].cells[0].text = f"{HEADER_DEPARTMENT_PREFIX}"
        
        # Center column (Logo and title)
        # Add logo to the first row
        logo_cell = header_table.rows[0].cells[1]
        self._add_logo_to_cell(logo_cell)
        
        # Add title to the third row (different from original logic)
        title_cell = header_table.rows[2].cells[1]
        if title_cell.paragraphs[0].runs:
            # If logo was added, add title to a new paragraph
            title_para = title_cell.add_paragraph()
            title_para.text = HEADER_SCHEDULE_TEMPLATE
            title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            self._set_paragraph_rtl(title_para)
        else:
            # If no logo, add title to first paragraph
            title_cell.text = HEADER_SCHEDULE_TEMPLATE
            title_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            self._set_cell_rtl(title_cell)
        
        # Empty second row in center column
        header_table.rows[1].cells[1].text = ""
        
        # Left column (Academic details)
        header_table.rows[0].cells[2].text = HEADER_ACADEMIC_YEAR
        header_table.rows[1].cells[2].text = HEADER_SEMESTER
        header_table.rows[2].cells[2].text = f"{HEADER_LEVEL_PREFIX} {level} {HEADER_DIVISION_PREFIX} {speciality}"
    
    def _apply_header_formatting(self, header_table) -> None:
        """Apply formatting to header table"""
        for row in header_table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in paragraph.runs:
                        run.font.name = FontConfig.FONT_NAME
                        run.font.size = FontConfig.HEADER_SIZE
                        run.font.bold = True
                        self._set_paragraph_rtl(paragraph)
        
        # Set header table to RTL
        self._set_table_rtl(header_table)
        
        # Remove borders from header table to match image
        self._remove_table_borders(header_table)
    
    def _add_header_to_section(self, section, speciality: str, level: str) -> None:
        """Add header to a specific section"""
        # Create header
        header = section.header
        
        # Clear existing content safely
        self._clear_section_content(header)
        
        # Create and configure header table
        header_table = self._create_header_table(header)
        
        # Fill header content
        self._fill_header_content(header_table, speciality, level)
        
        # Apply formatting
        self._apply_header_formatting(header_table)
    
    def _add_footer_to_section(self, section) -> None:
        """Add footer to a specific section"""
        # Create footer
        footer = section.footer
        
        # Clear existing content safely
        self._clear_section_content(footer)
        
        # Create footer table with three columns
        footer_table = self._create_footer_table(footer)
        
        # Fill footer content
        self._fill_footer_content(footer_table)
        
        # Apply formatting
        self._apply_footer_formatting(footer_table)
        
        # Add generation info below the table
        self._add_generation_info(footer)
    
    def _add_logo_to_cell(self, cell) -> None:
        """Add logo image to a table cell"""
        try:
            # Get the path to the logo file
            current_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            logo_path = os.path.join(current_dir, "assets", "logo.png")
            
            # Check if logo file exists
            if os.path.exists(logo_path):
                # Clear any existing content in the cell
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.text = ""
                
                # Add the logo image to the first paragraph
                paragraph = cell.paragraphs[0]
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Add the image with appropriate size (smaller to match image)
                run = paragraph.add_run()
                run.add_picture(logo_path, width=Inches(0.98), height=Inches(0.55))
                
                # Set RTL for the paragraph
                self._set_paragraph_rtl(paragraph)
            else:
                # If logo doesn't exist, add placeholder text
                cell.text = "[LOGO]"
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                self._set_cell_rtl(cell)
                
        except Exception as e:
            # If there's any error, add placeholder text
            cell.text = "[LOGO]"
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            self._set_cell_rtl(cell)
    
    def _create_footer_table(self, footer) -> Any:
        """Create and configure footer table with proper column widths"""
        # Add footer table with three columns: right, center, left
        footer_table = footer.add_table(rows=2, cols=3, width=Inches(TableDimensions.AVAILABLE_WIDTH_INCHES))
        footer_table.style = 'Table Grid'
        footer_table.autofit = False
        footer_table.allow_autofit = False
        
        # Set column widths (right, center, left)
        column_widths = [2.5, 3.0, 2.5]  # in inches
        for i, column in enumerate(footer_table.columns):
            for cell in column.cells:
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{int(column_widths[i] * 1440)}" w:type="dxa"/>')
                tcPr.append(tcW)
        
        return footer_table
    
    def _fill_footer_content(self, footer_table) -> None:
        """Fill footer table with content"""
        # Right column (Dean info)
        footer_table.rows[0].cells[0].text = FOOTER_DEAN_TITLE
        footer_table.rows[1].cells[0].text = FOOTER_DEAN_NAME
        
        # Center column (Vice Dean info)
        footer_table.rows[0].cells[1].text = FOOTER_VICE_DEAN_TITLE
        footer_table.rows[1].cells[1].text = FOOTER_VICE_DEAN_NAME
        
        # Left column (Head of Department info)
        footer_table.rows[0].cells[2].text = FOOTER_HEAD_DEPARTMENT_TITLE
        footer_table.rows[1].cells[2].text = FOOTER_HEAD_DEPARTMENT_NAME
    
    def _apply_footer_formatting(self, footer_table) -> None:
        """Apply formatting to footer table"""
        for row in footer_table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in paragraph.runs:
                        run.font.name = FontConfig.FONT_NAME
                        run.font.size = FontConfig.FOOTER_SIZE
                        run.font.bold = True
                        self._set_paragraph_rtl(paragraph)
        
        # Set footer table to RTL
        self._set_table_rtl(footer_table)
        
        # Remove borders from footer table to match image
        self._remove_table_borders(footer_table)
    
    def _add_generation_info(self, footer) -> None:
        """Add generation information below the footer table"""
        # Add spacing
        
        
        # Add generation info paragraph
        footer_para = footer.add_paragraph()
        footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Add footer text
        current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        footer_text = FOOTER_GENERATION_INFO.format(date=current_date)
        footer_run = footer_para.add_run(footer_text)
        footer_run.font.name = FontConfig.FONT_NAME
        footer_run.font.size = FontConfig.FOOTER_SIZE
        footer_run.font.italic = True
        
        # Set footer to RTL
        self._set_paragraph_rtl(footer_para)
    
    def _remove_table_borders(self, table) -> None:
        """Remove all borders from a table"""
        tbl = table._tbl
        
        # Get or create tblPr element
        tblPr = tbl.tblPr
        if tblPr is None:
            tblPr = parse_xml(f'<w:tblPr {nsdecls("w")}/>')
            tbl.insert(0, tblPr)
        
        # Remove any existing tblBorders setting
        for child in tblPr:
            if child.tag.endswith('tblBorders'):
                tblPr.remove(child)
        
        # Add no borders setting
        tblBorders = parse_xml(f'<w:tblBorders {nsdecls("w")}>'
                             f'<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                             f'<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                             f'<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                             f'<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                             f'<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                             f'<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                             f'</w:tblBorders>')
        tblPr.append(tblBorders)
    
    def generate_word_document(self, weekly_schedule: WeeklySchedule, output_path: str) -> None:
        """Generate Word document from weekly schedule (backward compatibility)"""
        doc = self.create_document()
        self.create_table_structure(doc, weekly_schedule)
        doc.save(output_path)
    
    def generate_multi_level_word_document(self, multi_level_schedule: MultiLevelSchedule, output_path: str) -> None:
        """Generate Word document with multiple tables for each specialty-level combination"""
        doc = self.create_document()
        
        # Generate table for each specialty-level combination with sections
        for i, schedule in enumerate(multi_level_schedule.schedules):
            print(f"📄 Generating table for {schedule.speciality} - {schedule.level}")
            
            # Create new section for each page (except the first one)
            if i > 0:
                # Add page break
                doc.add_page_break()
                # Create new section
                new_section = doc.add_section()
                # Break the link to previous section's header
                new_section.header.is_linked_to_previous = False
                new_section.footer.is_linked_to_previous = False
                # Add header and footer to the new section
                self._add_header_to_section(new_section, schedule.speciality, schedule.level)
                self._add_footer_to_section(new_section)
            else:
                # First page uses the default section
                section = doc.sections[0]
                # Break the link to previous section's header (for first section, this ensures it's independent)
                section.header.is_linked_to_previous = False
                section.footer.is_linked_to_previous = False
                self._add_header_to_section(section, schedule.speciality, schedule.level)
                self._add_footer_to_section(section)
            
            # Create table for this combination
            doc.add_paragraph()
            self.create_table_structure(doc, schedule.weekly_schedule)
        
        doc.save(output_path)
