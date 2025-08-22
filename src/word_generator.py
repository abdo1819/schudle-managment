from typing import List, Dict, Any
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from .models import WeeklySchedule, ScheduleEntry, DayOfWeek, TimeSlot, DetailCategory, TableCell


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
    
    def create_document(self) -> Document:
        """Create a new Word document"""
        doc = Document()
        
        # Set document margins
        sections = doc.sections
        for section in sections:
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(0.5)
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)
        
        return doc
    
    def create_table_structure(self, doc: Document, weekly_schedule: WeeklySchedule) -> None:
        """Create the main table structure"""
        # Create table: 21 rows (5 days * 4 categories + 1 header) x 9 columns (6 data + 3 separators)
        table = doc.add_table(rows=21, cols=9)
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Set table width and disable autofit
        table.autofit = False
        table.allow_autofit = False
        
        # A4 page width is 8.27 inches, accounting for margins (0.5 inch each side) = 7.27 inches available
        # Calculate proportional widths to fit within A4
        available_width = 7.27  # inches
        
        # Define column widths that add up to available width
        column_widths = {
            0: 0.8,   # Days column
            1: 1.0,   # Categories column  
            2: 1.2,   # Time slot 1
            3: 0.2,  # Separator 1
            4: 1.2,   # Time slot 2
            5: 0.2,  # Separator 2
            6: 1.2,   # Time slot 3
            7: 0.2,  # Separator 3
            8: 1.2    # Time slot 4
        }
        
        # Verify total width
        total_width_inches = sum(column_widths.values())
        table.width = Inches(total_width_inches)
        
        # Set individual column widths using XML properties
        for i, column in enumerate(table.columns):
            # Get width for this column
            width = Inches(column_widths[i])
            
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
        header_row = table.rows[0]
        
        # First two cells are empty in header
        header_row.cells[0].text = ""
        header_row.cells[1].text = ""
        
        # Fill each time slot in separate columns with separators
        time_slot_positions = [2, 4, 6, 8]  # Positions for time slots
        separator_positions = [3, 5, 7]     # Positions for separators
        
        for i, time_slot in enumerate(self.time_slots):
            header_row.cells[time_slot_positions[i]].text = time_slot
            header_row.cells[time_slot_positions[i]].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Set separator columns to empty
        for pos in separator_positions:
            header_row.cells[pos].text = ""
    
    def _fill_content_rows(self, table, weekly_schedule: WeeklySchedule) -> None:
        """Fill the content rows with schedule data"""
        row_index = 1
        
        for day_index, day_arabic in enumerate(self.days_arabic):
            day_key = list(weekly_schedule.keys())[day_index]
            day_schedule = weekly_schedule[day_key]
            
            # For each day, create 4 rows (one for each detail category)
            for category_index, category in enumerate(self.detail_categories):
                row = table.rows[row_index]
                
                # Day column (merged vertically for each day)
                if category_index == 0:
                    row.cells[0].text = day_arabic
                    row.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    # Merge vertically with next 3 rows
                    if row_index + 3 < len(table.rows):
                        row.cells[0].merge(table.rows[row_index + 3].cells[0])
                
                # Category column
                row.cells[1].text = category
                row.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
                
                # Time slot columns with separators
                slot_keys = ["first", "second", "third", "fourth"]
                time_slot_positions = [2, 4, 6, 8]  # Positions for time slots
                separator_positions = [3, 5, 7]     # Positions for separators
                
                for slot_index, slot_key in enumerate(slot_keys):
                    cell = row.cells[time_slot_positions[slot_index]]
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
                for pos in separator_positions:
                    row.cells[pos].text = ""
                
                row_index += 1
    
    def _apply_formatting(self, table) -> None:
        """Apply formatting to the table"""
        # Apply borders and formatting to all cells
        for row in table.rows:
            for cell in row.cells:
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
    
    def generate_word_document(self, weekly_schedule: WeeklySchedule, output_path: str) -> None:
        """Generate Word document from weekly schedule"""
        doc = self.create_document()
        self.create_table_structure(doc, weekly_schedule)
        doc.save(output_path)
