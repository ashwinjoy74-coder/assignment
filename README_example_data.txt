SCHOOL TIMETABLE GENERATOR - EXAMPLE PROJECT

Files:
- school_timetable_generator.py: simplified Python timetable generator
- requirements.csv: fictional sample input for all 48 divisions
- requirements.txt: Python packages needed by the program

WHAT THE PROGRAM DOES
- Reads subject, teacher and weekly period information from requirements.csv.
- Uses OR-Tools to assign every lesson to a day and period.
- Prevents a division from having two lessons at the same time.
- Prevents a teacher from being assigned to two divisions at the same time.
- Creates school_timetable.xlsx with one sheet for every division and teacher.

CSV COLUMNS
- class: grade number, 1 through 12
- division: A, B, C, or D
- subject: subject taught to that division
- teacher: teacher name (use the same spelling everywhere for the same teacher)
- periods_per_week: number of lessons per week for that subject/teacher assignment

IMPORTANT
The sample data uses fictional teacher names and illustrative subject allocations. Replace it with your school's actual details before using the timetable.

HOW TO RUN IT IN PYCHARM
1. Keep school_timetable_generator.py, requirements.csv, and requirements.txt in the same folder.
2. Open the PyCharm Terminal in that folder.
3. Run: python -m pip install -r requirements.txt
4. Run: python school_timetable_generator.py
5. The Excel output, school_timetable.xlsx, will be saved in the same folder as the Python file.

The school week is currently set to Monday-Friday with 7 periods each day. Change DAYS and PERIODS_PER_DAY near the top of the Python file to adjust it.
