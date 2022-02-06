Attribute VB_Name = "SplitWorkbook_row_wise_to_sheet"
'Callback for customButton3 onAction
Sub rwSp(control As IRibbonControl)
Call rowWiseSplit
End Sub


Sub rowWiseSplit()
question = "Warning : Once you done can't undo the task,Please backup your main file" & vbCrLf & vbCrLf & "Are you sure you want to run this Macro? "
If MsgBox(question, vbYesNo + vbQuestion) = vbYes Then
On Error GoTo ER


Dim lngLastRow As Long
Dim lngNumberOfRows As Long
Dim lngI As Long
Dim strMainSheetName As String
Dim currSheet As Worksheet
Dim prevSheet As Worksheet

'Number of rows to split among worksheets

lngNumberOfRows = InputBox("Plese enter row count :", "PyGems Split_Row_Wise", "")


'Current worksheet in workbook
Set prevSheet = ActiveWorkbook.ActiveSheet
'First worksheet name
strMainSheetName = prevSheet.Name
'Number of rows in worksheet
lngLastRow = prevSheet.Cells(Rows.Count, 1).End(xlUp).Row
'Worksheet counter for added worksheets
lngI = 1
While lngLastRow > lngNumberOfRows
Set currSheet = ActiveWorkbook.Worksheets.Add
With currSheet
.Move after:=Worksheets(Worksheets.Count)
.Name = strMainSheetName + "(" + CStr(lngI) + ")"
End With

With prevSheet.Rows(lngNumberOfRows + 1 & ":" & lngLastRow).EntireRow
.Cut currSheet.Range("A1")
End With

lngLastRow = currSheet.Cells(Rows.Count, 1).End(xlUp).Row
Set prevSheet = currSheet
lngI = lngI + 1
Wend

MsgBox "Completed successfully."

Else
End If
ER:

End Sub






