# macronator

This VBA script automates the process of opening multiple Excel workbooks, executing a specified macro in each workbook, saving the workbooks with new names in a user-selected location, and closing the workbooks.

## Prerequisites

- Microsoft Excel

## Usage

1. Open the Excel file containing the macro you want to execute.
2. Press `Alt + F11` to open the Visual Basic Editor.
3. Insert a new module or open an existing module.
4. Copy and paste the VBA code from [`macronator_script.cls`](macronator_script.cls) into the module.
5. Save the Excel file as a macro-enabled workbook with the extension `.xlsm`.
6. Close the Visual Basic Editor.

To run the script:

1. Open the Excel file that contains the macro and the VBA script.
2. Press `Alt + F8` to open the "Macro" dialog box.
3. Select the `Multi_Excel_Macro_Automation` macro from the list.
4. Click "Run" to execute the macro.

The script will prompt you for the following inputs:

- Macro Name: Enter the name of the macro you want to execute in each workbook.
- Number of Workbooks: Enter the number of workbooks you want to process.
- Save Location: Select the location where you want to save the new files.

The script will then open a file picker dialog for each workbook, allowing you to select the files one by one. After selecting a workbook, the specified macro will be executed, and the workbook will be saved with a new name in the chosen save location. Once all the workbooks have been processed, a success message will be displayed.

Please note that the script assumes the macro you want to execute exists in each selected workbook with the same name entered by you. Make sure to replace the file paths, file names, and macro names in the VBA code to match your specific requirements.

## License

This script is licensed under the [MIT License](Lisence.md).

