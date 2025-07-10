from dataclasses import dataclass
from typing import List, Optional

@dataclass
class Student:
    first_name: str
    last_name: str
    stream: str = "Default"

    @property
    def full_name(self) -> str:
        return f"{self.last_name}, {self.first_name}"

@dataclass
class ClassGroup:
    name: str
    cohort: str
    students: List[Student]

    def add_student(self, student: Student):
        self.students.append(student)

    def pad_students(self, target_count: int = 15):
        """
        如果学生数不够 target_count, 就补空白
        """
        while len(self.students) < target_count:
            self.students.append(Student(first_name="", last_name="", stream=self.name))
