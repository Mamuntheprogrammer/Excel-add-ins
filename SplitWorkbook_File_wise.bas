Attribute VB_Name = "SplitWorkbook_File_wise"
'Callback for customButton6 onAction
Sub swSp(control As IRibbonControl)
Call fileWiseSheet
End Sub




Sub fileWiseSheet()
question = "Warning : Once you done can't undo the task,Please backup your main file" & vbCrLf & vbCrLf & "Are you sure you want to run this Macro? "
If MsgBox(question, vbYesNo + vbQuestion) = vbYes Then
On Error GoTo ER


Dim xPath As String
xPath = Application.ActiveWorkbook.Path
Application.ScreenUpdating = False
Application.DisplayAlerts = False
For Each xWs In ActiveWorkbook.Sheets
    xWs.Copy
    Application.ActiveWorkbook.SaveAs FileName:=xPath & "\" & xWs.Name & ".xlsx"
    Application.ActiveWorkbook.Close False
Next
Application.DisplayAlerts = True
Application.ScreenUpdating = True

MsgBox "Completed successfully."

Else
End If
ER:

End Sub
