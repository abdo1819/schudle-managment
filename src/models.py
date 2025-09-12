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
    """Model for Excel/CSV row data"""
    is_valid: Optional[str] = None
    day: str
    slot: int
    is_half_slot: Optional[bool] = Field(default=False, description="Whether this is a half slot")
    code: str
    speciality: Optional[str] = None
    activity_type: str = Field(alias="activityType")
    location: str
    active_tutor: Optional[str] = None
    level: Optional[str] = None
    course_name: Optional[str] = Field(alias="course_name", default=None)
    day_slot: str = Field(alias="day_slot")
    specialy_level: Optional[str] = None
    time: str
    day_order: int = Field(alias="day_order")
    confirmed_by_tutor: Optional[str] = Field(alias="confirmed by tutor", default=None)
    teaching_hours: Optional[str] = Field(alias="teaching_hours", default=None)
    teaching_hours_printable: Optional[str] = Field(alias="teachin_hourse_printalble", default=None)
    sp_code: Optional[str] = None
    main_tutor: Optional[str] = Field(alias="main_tutor", default=None)
    helping_stuff: Optional[str] = Field(alias="helping_stuff", default=None)
    main_tutor_write: Optional[str] = Field(alias="main_tutor_write", default=None)
    helping_stuff_write: Optional[str] = Field(alias="helping_stuff_write", default=None)

    def get_level(self) -> str:
        """Get the level as a string, converting float values like 100.0 to '100'"""
        if self.level is None:
            return "عام"
        # Convert to float first to handle string numbers, then to int to remove decimals, then to string
        return str(int(float(self.level)))


class ScheduleEntry(BaseModel):
    """Model for a single schedule entry"""
    course_name: Optional[str] = None
    location: str
    instructor: Optional[str] = None
    assistant: Optional[str] = None
    is_half_slot: bool = Field(default=False, description="Whether this is a half slot")


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


class SpecialityLevelSchedule(BaseModel):
    """Schedule for a specific specialty and level combination"""
    speciality: str
    level: str
    weekly_schedule: WeeklySchedule


class MultiLevelSchedule(BaseModel):
    """Complete multi-level schedule with all specialty-level combinations"""
    schedules: List[SpecialityLevelSchedule]
    
    def get_speciality_levels(self) -> List[tuple[str, str]]:
        """Get list of all specialty-level combinations"""
        return [(schedule.speciality, schedule.level) for schedule in self.schedules]
    
    def get_schedule_by_speciality_level(self, speciality: str, level: str) -> Optional[WeeklySchedule]:
        """Get weekly schedule for specific specialty and level"""
        for schedule in self.schedules:
            if schedule.speciality == speciality and schedule.level == level:
                return schedule.weekly_schedule
        return None


class LocationSchedule(BaseModel):
    """Schedule for a specific location"""
    location: str
    weekly_schedule: WeeklySchedule


class MultiLocationSchedule(BaseModel):
    """Complete multi-location schedule with all locations"""
    schedules: List[LocationSchedule]

    def get_locations(self) -> List[str]:
        """Get list of all locations"""
        return [schedule.location for schedule in self.schedules]

    def get_schedule_by_location(self, location: str) -> Optional[WeeklySchedule]:
        """Get weekly schedule for a specific location"""
        for schedule in self.schedules:
            if schedule.location == location:
                return schedule.weekly_schedule
        return None


class TableCell(BaseModel):
    """Model for a table cell"""
    content: str
    is_merged: bool = False
    merge_span: int = 1
    alignment: str = "left"  # left, center, right
