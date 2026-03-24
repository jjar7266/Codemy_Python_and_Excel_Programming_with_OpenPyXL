"""
Codemy.com
Course: Python and Excel Programming With OpenPyXL
Instructor: John Elder
Student: Jose "Joe" Ruiz

Lesson: Add Data To New Spreadsheet Workbook
"""

# ---------------------------------------------------------
# Import Modules
# ---------------------------------------------------------
from openpyxl import Workbook  # Correct import path
from openpyxl import load_workbook  # Included for course continuity

# ---------------------------------------------------------
# Create a New Workbook
# ---------------------------------------------------------
wb = Workbook()

# ---------------------------------------------------------
# Select Active Worksheet
# ---------------------------------------------------------
ws = wb.active

# ---------------------------------------------------------
# Set Worksheet Title
# ---------------------------------------------------------
ws.title = "Names and Colors"  # type: ignore

# ---------------------------------------------------------
# Create Python Lists of Names and Colors
# ---------------------------------------------------------
names = ["Dan", "April", "Neal", "Sara"]
colors = ["Blue", "Purple", "Green", "White"]

# ---------------------------------------------------------
# Add Column Headers
# ---------------------------------------------------------
ws["A1"] = "Names"   # type: ignore
ws["B1"] = "Colors"  # type: ignore

# ---------------------------------------------------------
# Add Names to Worksheet
# ---------------------------------------------------------
starting_name_row = 2
for name in names:
    ws.cell(row=starting_name_row, column=1).value = name  # type: ignore
    starting_name_row += 1

# ---------------------------------------------------------
# Add Colors to Worksheet
# ---------------------------------------------------------
starting_color_row = 2
for color in colors:
    ws.cell(row=starting_color_row, column=2).value = color  # type: ignore
    starting_color_row += 1

# ---------------------------------------------------------
# Save Spreadsheet
# ---------------------------------------------------------
wb.save("colors.xlsx")

print("File Was Saved...")

# =========================================================
# Teaching Notes
# =========================================================
"""
1. Workbook() and Active Sheet
   - Workbook() creates a brand‑new Excel file in memory.
   - wb.active returns the first worksheet automatically.

2. Adding Data
   - You can assign values using ws["A1"] or ws.cell().
   - ws.cell() is ideal when looping because it accepts numeric row/column.

3. Parallel Lists
   - Using two lists (names and colors) keeps data organized.
   - Each loop writes one row at a time.

4. Row Tracking
   - starting_name_row and starting_color_row ensure data lines up correctly.
"""

# =========================================================
# Example Variations
# =========================================================
"""
Example 1 — Add Names and Colors in One Loop
--------------------------------------------
for row, (name, color) in enumerate(zip(names, colors), start=2):
    ws.cell(row=row, column=1).value = name
    ws.cell(row=row, column=2).value = color

Example 2 — Add More Columns
----------------------------
ages = [30, 28, 35, 22]
ws["C1"] = "Ages"
for row, age in enumerate(ages, start=2):
    ws.cell(row=row, column=3).value = age

Example 3 — Auto‑Number Rows
----------------------------
ws["D1"] = "ID"
for i, _ in enumerate(names, start=1):
    ws.cell(row=i+1, column=4).value = i
"""

# =========================================================
# Common Mistakes
# =========================================================
"""
- Forgetting to increment the row counter inside loops.
- Mixing up row and column numbers.
- Saving to the wrong filename and overwriting previous lessons.
- Using mismatched list lengths (names vs colors).
"""

# =========================================================
# Takeaways
# =========================================================
"""
- Workbook() creates a new Excel file from scratch.
- Lists + loops make it easy to populate spreadsheets.
- ws.cell() is ideal for dynamic row/column operations.
- Always save your workbook after writing data.
"""
