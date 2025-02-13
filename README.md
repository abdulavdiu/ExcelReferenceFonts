# ExcelReferenceFonts
This VBA macro colors cells in an Excel worksheet based on whether they:
- Contain a formula referencing another sheet (green font)
- Contain a formula referencing the same sheet (black font)
- Contain no formula (blue font)

**How It Works**
- Select a range and run the macro to color only that range.
- If no range is selected (or if only a single cell is selected) when the macro is run, you will be prompted to color the entire worksheet.

**Important Note**
- When you apply this to the entire worksheet, any cells without a formula (including header cells) will be colored blue.

**How to Use**
- In Excel, press ALT + F11 to open the Visual Basic Editor.
- In the Project Explorer, right-click your workbook, select Insert, then choose Module.
- Copy and paste the macro code into the new module.
- Save the file as a macro-enabled workbook (.xlsm).
- Run the macro from the Developer tab (or via Tools > Macros in the VBA Editor).
