Sub Multi-Excel_Macro_Automation()
'************************************************************************************
'Opens multiple excel workbooks, executes a macro in each workbook, saves the workbook with a new name in a new location, and closes the workbook. 
'************************************************************************************
‘
Dim wb As Workbook
‘
Dim Workbook_1 As String
Dim Workbook_2 As String
Dim Workbook_3 As String
Dim Workbook_4 As String
Dim Workbook_5 As String
‘
'User input for week ending date to add to beginning of each new file name
wedate = InputBox("Please enter the week ending date using the format below:", "Weekending Date", "00-00-00")
‘
‘Replace file paths, file names, and macro names to match your requirements
‘
'Script for Workbook_1 
Workbook_1 = "C:\Location\originalfile1.xlsm"
Set wb = Workbooks.Open(Workbook_1)
Application.Run ("'originalfile1.xlsm'!macroname")
newname = wedate & Space(1) & "- newfile1.xlsm"
wb.SaveAs Filename:="C:\New Location\ " & newname, FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
wb.Close
‘
'Script for Workbook_2 
Workbook_2 = "C:\Location\originalfile2.xlsm"
Set wb = Workbooks.Open(Workbook_2)
Application.Run ("'originalfile2.xlsm'!macroname")
newname = wedate & Space(1) & "- newfile2.xlsm"
wb.SaveAs Filename:="C:\New Location\ " & newname, FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
wb.Close
‘
'Script for Workbook_3 
Workbook_3 = "C:\Location\originalfile3.xlsm"
Set wb = Workbooks.Open(Workbook_3)
Application.Run ("'originalfile3.xlsm'!macroname")
newname = wedate & Space(1) & "- newfile3.xlsm"
wb.SaveAs Filename:="C:\New Location\ " & newname, FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
wb.Close
‘
'Script for Workbook_4 
Workbook_4 = "C:\Location\originalfile4.xlsm"
Set wb = Workbooks.Open(Workbook_4)
Application.Run ("'originalfile4.xlsm'!macroname")
newname = wedate & Space(1) & "- newfile4.xlsm"
wb.SaveAs Filename:="C:\New Location\ " & newname, FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
wb.Close
‘
'Script for Workbook_5 
Workbook_5 = "C:\Location\originalfile5.xlsm"
Set wb = Workbooks.Open(Workbook_5)
Application.Run ("'originalfile5.xlsm'!macroname")
newname = wedate & Space(1) & "- newfile5.xlsm"
wb.SaveAs Filename:="C:\New Location\ " & newname, FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
wb.Close
‘
‘Popup message that initiates after macro finishes 
MsgBox "Success!"
‘
End Sub
