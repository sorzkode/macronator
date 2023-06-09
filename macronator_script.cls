Sub Multi_Excel_Macro_Automation()
    '************************************************************************************
    '+-+-+-+-+
    'macronator    
    '+-+-+-+-+
    'VBA script that automates running macros inside of other excel files.
    
    'Author:
    'Mister Riley
    'sorzkode@proton.me
    'https://github.com/sorzkode

    'MIT License
    'Copyright (c) 2023 Mister Riley
    'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
    'to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
    'and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
    'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
    'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
    'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
    'WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
    '************************************************************************************
    
    Dim wb As Workbook
    Dim filePath As Variant
    Dim macroName As String
    Dim newFilePath As Variant
    Dim newFileName As String
    Dim i As Integer
    Dim numWorkbooks As Integer
    
    On Error GoTo ErrorHandler
    
    ' User input for macro name
    macroName = InputBox("Please enter the name of the macro to execute:", "Macro Name")
    
    ' User input for number of workbooks to process
    numWorkbooks = InputBox("Please enter the number of workbooks to process:", "Number of Workbooks")
    
    ' User input for save location
    newFilePath = Application.FileDialog(msoFileDialogFolderPicker).Show
    If newFilePath = -1 Then
        newFilePath = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
    Else
        Exit Sub ' User cancelled save location selection
    End If
    
    ' Loop to process multiple workbooks
    For i = 1 To numWorkbooks
        ' Prompt user to select the workbook
        filePath = Application.GetOpenFilename("Excel Files (*.xlsm), *.xlsm", Title:="Select Workbook " & i)
        
        If filePath <> False Then ' User selected a file
            ' Open the workbook
            Set wb = Workbooks.Open(filePath)
            
            ' Execute the specified macro
            Application.Run (wb.Name & "!" & macroName)
            
            ' Generate new file name
            newFileName = Format(Date, "dd-mm-yy") & " - NewFile" & i & ".xlsm"
            
            ' Generate new file path
            newFilePath = newFilePath & "\" & newFileName
            
            ' Save the workbook with the new name and location
            wb.SaveAs Filename:=newFilePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
            
            ' Close the workbook
            wb.Close SaveChanges:=False
        Else
            Exit For ' User cancelled file selection
        End If
    Next i
    
    ' Popup message that initiates after the macro finishes
    MsgBox "Success!"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
    
End Sub


