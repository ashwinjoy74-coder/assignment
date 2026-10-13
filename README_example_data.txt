SCHOOL TIMETABLE GENERATOR — EXAMPLE INPUT DATA

Files:
- school_timetable_generator.py: timetable generator
- requirements.csv: populated EXAMPLE input data (ready for the script)
- requirements.txt: Python dependencies

IMPORTANT
This is fictional example data to demonstrate the expected CSV structure. Teacher names, subjects, and weekly lesson counts are illustrative. Replace them with your school's actual assignments before relying on the generated timetable.

CSV columns
- class: grade number, 1 through 12
- division: A, B, C, or D
- subject: subject taught to that division
- teacher: consistent teacher identifier; identical names are treated as the same teacher
- periods_per_week: weekly number of lessons for this teacher/subject assignment

A row represents one subject-teacher assignment for one division. Add multiple rows for each division (one per subject assignment). The example includes all 48 divisions and keeps each division within 35 weekly periods (5 days x 7 periods).

Run
1. Install Python 3.10+.
2. In this folder, run: python -m pip install -r requirements.txt
3. Run: python school_timetable_generator.py
4. The script writes school_timetable.xlsx with class and teacher timetables.
