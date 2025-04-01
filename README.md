# Logbook Project

This Excel VBA Logbook is designed to help manage and record tasks for each shift. It provides:

1. **Automatic Tasks Loading**  
   When you choose a shift (morning/evening/night/friday/saturday), the system clears old tasks and loads default tasks from a separate "Tasks" sheet into specific cells in the main sheet.

2. **Save and Load Functions**  
   - **SaveShiftData**: Reads all relevant blocks (D10..D18, F10..F18, etc.) and saves them to a “DB” sheet, keyed by Date (G5) and Shift (E5).  
   - **LoadShiftData**: When you switch Date or Shift, the code looks up previous entries in “DB” and loads them back into the main sheet, clearing old data if needed.

3. **CheckBoxes**  
   - Two sets of checkboxes (B6..B14, C6..C14) let you mark tasks as Done (TRUE) or Not Done (FALSE).  
   - The main cells (E10..E18, G10..G18) display **“Done”** or **“Not Done”** depending on the checkbox values.

4. **Operator, Shift, and Date**  
   - **Operator** in D5 – The user’s name for the shift.  
   - **Shift** in E5 – e.g., "morning", "evening", "night", "friday", "saturday".  
   - **Date** in G5 – The current date of the shift.

5. **Blocks** (for Save/Load)  
   - **Block1**: D10..D18 → columns D..L  
   - **Block2**: F10..F18 → columns M..U  
   - **Block3**: D21..D33 → columns V..AH  
   - **Block4**: E21..E33 → columns AI..AU  
   - **Block5**: F21..F33 → columns AV..BH  
   - **Block6**: G21..G33 → columns BI..BU  
   - **Block7**: B6..B14  → columns BV..CD  (checkboxes for E10..E18)  
   - **Block8**: C6..C14  → columns CE..CM  (checkboxes for G10..G18)

6. **CreateYesterdayStaticCopy** (optional)  
   A utility macro that creates a static copy of the workbook (e.g., for archiving).

---

## Getting Started

1. **Enable Macros** in Excel so the VBA code can run.  
2. **Open** `logbook.xlsm`.  
3. In the "Main" sheet:  
   - Fill in **Date** (G5) and **Shift** (E5).  
   - If you pick a shift for the first time, tasks will load from the "Tasks" sheet.  
   - Fill out any additional details, and if needed, check/uncheck boxes in B6..B14 or C6..C14 to mark tasks as “Done” or “Not Done.”  
   - At the end of the shift, press **Save** (or run the `SaveShiftData` macro).  
4. **Reloading** old data: if you re-select an existing date & shift, the `LoadShiftData` macro auto-runs and loads previously saved info from the “DB” sheet.

---

## Project Structure

- **Main_VBA.bas**: Contains the `Worksheet_Change` event that reacts to changes in Shift (E5) and Date (G5), clearing old tasks and calling `LoadShiftData`.  
- **SaveShiftData_VBA.bas**: The code for saving all blocks into the “DB” sheet.  
- **LoadShiftData_VBA.bas**: The code for loading data from “DB” back into the main sheet.  
- **CreateCopy_VBA.bas** (optional): A macro that can create a static archive copy of the workbook.

---

## Demo & Screenshots

![image](https://github.com/user-attachments/assets/e9f69684-4b91-49ed-9a07-eae76a06b92b)

