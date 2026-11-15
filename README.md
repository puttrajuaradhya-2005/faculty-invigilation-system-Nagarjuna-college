# Faculty Free-Slot Based Invigilation Duty Allocation System

Django project for Nagarjuna Degree College faculty invigilation allocation.

## What this version does

The application has **exactly three input files**:

1. Faculty Timetable
2. Exam Timetable
3. Room Details

There is **no Class/XA timetable input** and no faculty/course/room data is hard-coded.

After upload, the system automatically:

1. Reads the faculty timetable.
2. Extracts faculty names and their recurring weekly occupied slots.
3. Reads the exam timetable, including scanned PDFs through OCR.
4. Converts exam Slot A-H into actual date/time records when the timetable uses the supported college CA format.
5. Reads room numbers from the room file.
6. Creates database records for the imported data.
7. During allocation, checks each exam date's weekday and rejects faculty whose timetable overlaps the exam time.
8. Assigns a different available faculty member to each room and balances the total duty count.
9. Shows alerts where there are not enough free faculty.
10. Exports the final duty list to CSV.

## Windows setup

Open PowerShell in this folder:

```powershell
cd "C:\Users\Admin\Downloads\FACULTY_INVIGILATION_SYSTEM_NAGARJUNA_DEGREE_COLLEGE"
py --version
py -m venv venv
.\venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
python manage.py migrate
python manage.py createsuperuser
python manage.py runserver
```

If PowerShell blocks activation:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\venv\Scripts\Activate.ps1
```

Open `http://127.0.0.1:8000/`.

## Important OCR requirement

For scanned exam PDFs, Python packages are included, but Tesseract OCR must also be installed on Windows and available on PATH. If `tesseract --version` is not recognized, install Tesseract OCR and restart VS Code/PowerShell.

## Allocation logic

A faculty member is considered unavailable when any imported recurring timetable slot has the same weekday and overlaps the exam interval. The allocator then sorts eligible faculty by the number of duties already assigned, with faculty name as a deterministic tie-breaker.

The room file is currently treated as the available room pool for each imported exam. If your room document contains a specific mapping such as `Date + Slot + Room`, the parser can be extended to make room assignment exam-specific.


FIX: CSV files are now read with UTF-8/UTF-8-SIG/CP1252/Latin-1 fallbacks. Legacy .doc files give a clear conversion message instead of a UnicodeDecodeError.
