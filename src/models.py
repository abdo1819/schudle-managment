from typing import Dict, List, Optional, TypedDict
from pydantic import BaseModel, Field
from enum import Enum


class DayOfWeek(str, Enum):
    """Days of the week in Arabic"""
    SUNDAY = "الأحد"
    MONDAY = "الاثنين"
    TUESDAY = "الثلاثاء"
    WEDNESDAY = "الأربعاء"
    THURSDAY = "الخميس"


class TimeSlot(str, Enum):
    """Time slots for lectures"""
    FIRST = "المحاضرة الأولى 8.50-10.30"
    SECOND = "المحاضرة الثانية 10.40 - 12.10"
    THIRD = "المحاضرة الثالثة 12.20 - 1.50"
    FOURTH = "المحاضرة الرابعة 2.00 - 3.30"


class DetailCategory(str, Enum):
    """Detail categories for each day"""
    COURSE_NAME = "اسم المادة"
    LOCATION = "المكان"
    INSTRUCTOR = "استاذ المادة"
    ASSISTANT = "الهيئة المعاونة"


class CSVRow(BaseModel):
    """Model for CSV row data"""
    day: str
    slot: int
    code: str
    activity_type: str = Field(alias="activityType")
    location: str
    course_name: str = Field(alias="course_name")
    day_slot: str = Field(alias="day_slot")
    time: str
    day_order: int = Field(alias="day_order")
    confirmed_by_tutor: Optional[str] = Field(alias="confirmed by tutor", default=None)
    teaching_hours: Optional[str] = Field(alias="teaching_hours", default=None)
    teaching_hours_printable: Optional[str] = Field(alias="teachin_hourse_printalble", default=None)
    main_tutor: str = Field(alias="main_tutor")
    helping_stuff: str = Field(alias="helping_stuff")


class ScheduleEntry(BaseModel):
    """Model for a single schedule entry"""
    course_name: str
    location: str
    instructor: str
    assistant: str


class DaySchedule(TypedDict):
    """Schedule for a single day with all time slots"""
    first: Optional[ScheduleEntry]
    second: Optional[ScheduleEntry]
    third: Optional[ScheduleEntry]
    fourth: Optional[ScheduleEntry]


class WeeklySchedule(TypedDict):
    """Complete weekly schedule"""
    sunday: DaySchedule
    monday: DaySchedule
    tuesday: DaySchedule
    wednesday: DaySchedule
    thursday: DaySchedule


class TableCell(BaseModel):
    """Model for a table cell"""
    content: str
    is_merged: bool = False
    merge_span: int = 1
    alignment: str = "left"  # left, center, right
