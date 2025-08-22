from typing import List, Dict, Any
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from .models import WeeklySchedule, ScheduleEntry, DayOfWeek, TimeSlot, DetailCategory, TableCell, MultiLevelSchedule, SpecialityLevelSchedule
from enum import Enum
from datetime import datetime


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
    
    def add_page_header(self, doc: Document, speciality: str, level: str) -> None:
        """Add page header with identity information"""
        # Get the current section
        section = doc.sections[0]
        
        # Create header
        header = section.header
        
        # Clear existing content safely
        try:
            for element in list(header._element):
                header._element.remove(element)
        except:
            # If clearing fails, create a new header
            pass
        
        # Add header table with width parameter
        header_table = header.add_table(rows=2, cols=2, width=Inches(TableDimensions.AVAILABLE_WIDTH_INCHES))
        header_table.style = 'Table Grid'
        header_table.autofit = False
        header_table.allow_autofit = False
        
        # Set column widths
        for column in header_table.columns:
            for cell in column.cells:
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{int(TableDimensions.AVAILABLE_WIDTH_INCHES * 720)}" w:type="dxa"/>')
                tcPr.append(tcW)
        
        # First row: University name and logo placeholder
        header_table.rows[0].cells[0].text = "جامعة القاهرة"
        header_table.rows[0].cells[1].text = "كلية الهندسة"
        
        # Second row: Department and academic year
        header_table.rows[1].cells[0].text = f"قسم {speciality} - {level}"
        header_table.rows[1].cells[1].text = f"العام الدراسي 2025-2026"
        
        # Apply formatting to header
        for row in header_table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in paragraph.runs:
                        run.font.name = 'Arial'
                        run.font.size = Pt(12)
                        run.font.bold = True
                        self._set_paragraph_rtl(paragraph)
        
        # Set header table to RTL
        self._set_table_rtl(header_table)
    
    def add_page_footer(self, doc: Document) -> None:
        """Add page footer with generation information"""
        # Get the current section
        section = doc.sections[0]
        
        # Create footer
        footer = section.footer
        
        # Clear existing content
        for element in footer._element:
            footer._element.remove(element)
        
        # Add footer paragraph
        footer_para = footer.add_paragraph()
        footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Add footer text
        current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        footer_text = f"{current_date}"
        footer_run = footer_para.add_run(footer_text)
        footer_run.font.name = 'Arial'
        footer_run.font.size = Pt(10)
        footer_run.font.italic = True
        
        # Set footer to RTL
        self._set_paragraph_rtl(footer_para)
    
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
        title_run.font.name = 'Arial'
        title_run.font.size = Pt(18)
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
        
        # Set table to RTL
        self._set_table_rtl(table)
        
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
            cell = header_row.cells[self.time_slot_positions[i]]
            cell.text = time_slot
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            # Set RTL for header cells
            self._set_cell_rtl(cell)
        
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
                
                # Category column
                category_cell = row.cells[ColumnType.CATEGORIES.value]
                category_cell.text = category
                category_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                self._set_cell_rtl(category_cell)
                
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
                    
                    # Set RTL and center alignment for all time slot cells
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    self._set_cell_rtl(cell)
                
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
                
                # Determine border widths
                # Vertical borders: thick for time slot columns, thin for others
                if col_index in self.time_slot_positions:
                    vertical_border_width = str(BorderWidth.THICK)
                else:
                    vertical_border_width = str(BorderWidth.THIN)
                
                # Horizontal borders depend on row type
                if row_type == RowType.HEADER:
                    # Header row - all borders thin
                    top_border_width = str(BorderWidth.THIN)
                    bottom_border_width = str(BorderWidth.THIN)
                elif row_type == RowType.DAY_START:
                    # First row of day - top border thick, bottom border thin
                    top_border_width = str(BorderWidth.THICK)
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
    

    

    
    def _add_headers_and_footers_to_pages(self, doc: Document, multi_level_schedule: MultiLevelSchedule) -> None:
        """Add headers and footers to each page with unique specialty-level information"""
        # For now, we'll use a simpler approach - just add headers to the first section
        # This will show the same header on all pages, but with the correct specialty and level
        if multi_level_schedule.schedules:
            # Use the first schedule's information for the header
            first_schedule = multi_level_schedule.schedules[0]
            self.add_page_header(doc, first_schedule.speciality, first_schedule.level)
            self.add_page_footer(doc)
    
    def _add_header_to_section(self, section, speciality: str, level: str) -> None:
        """Add header to a specific section"""
        # Create header
        header = section.header
        
        # Clear existing content safely
        try:
            for element in list(header._element):
                header._element.remove(element)
        except:
            pass
        
        # Add header table with width parameter
        header_table = header.add_table(rows=2, cols=2, width=Inches(TableDimensions.AVAILABLE_WIDTH_INCHES))
        header_table.style = 'Table Grid'
        header_table.autofit = False
        header_table.allow_autofit = False
        
        # Set column widths
        for column in header_table.columns:
            for cell in column.cells:
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{int(TableDimensions.AVAILABLE_WIDTH_INCHES * 720)}" w:type="dxa"/>')
                tcPr.append(tcW)
        
        # First row: University name and logo placeholder
        header_table.rows[0].cells[0].text = "جامعة القاهرة"
        header_table.rows[0].cells[1].text = "كلية الهندسة"
        
        # Second row: Department and academic year
        header_table.rows[1].cells[0].text = f"قسم {speciality} - {level}"
        header_table.rows[1].cells[1].text = f"العام الدراسي 2025-2026"
        
        # Apply formatting to header
        for row in header_table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in paragraph.runs:
                        run.font.name = 'Arial'
                        run.font.size = Pt(12)
                        run.font.bold = True
                        self._set_paragraph_rtl(paragraph)
        
        # Set header table to RTL
        self._set_table_rtl(header_table)
    
    def _add_footer_to_section(self, section) -> None:
        """Add footer to a specific section"""
        # Create footer
        footer = section.footer
        
        # Clear existing content safely
        try:
            for element in list(footer._element):
                footer._element.remove(element)
        except:
            pass
        
        # Add footer paragraph
        footer_para = footer.add_paragraph()
        footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Add footer text
        current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        footer_text = f"تم إنشاء هذا الجدول تلقائياً في {current_date} - نظام إدارة الجداول الدراسية"
        footer_run = footer_para.add_run(footer_text)
        footer_run.font.name = 'Arial'
        footer_run.font.size = Pt(10)
        footer_run.font.italic = True
        
        # Set footer to RTL
        self._set_paragraph_rtl(footer_para)
