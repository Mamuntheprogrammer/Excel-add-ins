Attribute VB_Name = "Format"
Sub rww(control As IRibbonControl)
Call Frmt
End Sub


Sub Frmt()
question = "Warning : Once you done can't undo the task,Please backup your main file" & vbCrLf & vbCrLf & "Are you sure you want to run this Macro? "
If MsgBox(question, vbYesNo + vbQuestion) = vbYes Then
On Error GoTo ER
    Dim xSheet As Worksheet
    Dim xSheetOne As Worksheet
    Dim xSheets As Sheets
    Dim ptRange As Range
    Dim lFirstRow, lLastRow As Long


    Set xSheets = ActiveWorkbook.Worksheets
    Set ptRange = Application.InputBox("Please select the top rows to repeat :", "pygems", "", Type:=8)
    lFirstRow = ptRange(1).Row
    lLastRow = ptRange(ptRange.Rows.Count, 1).Row

    For Each xSheet In xSheets
        xSheet.PageSetup.PrintTitleRows = lFirstRow & ":" & lLastRow
    xSheet.PageSetup.CenterFooter = "Page &P of &N"
    Next xSheet

MsgBox "Completed successfully."

Else
End If
ER:
End Sub
