Attribute VB_Name = "Split_Value_wise_to_Sheet"
'Callback for customButton4 onAction
Sub vwSp(control As IRibbonControl)
Call valueWiseSheet
End Sub


Sub valueWiseSheet()
question = "Warning : Once you done can't undo the task,Please backup your main file" & vbCrLf & vbCrLf & "Are you sure you want to run this Macro? "
If MsgBox(question, vbYesNo + vbQuestion) = vbYes Then
On Error GoTo ER


Dim lr As Long
Dim ws As Worksheet
Dim vcol, i As Long
Dim icol As Long
Dim myarr As Variant
Dim title As String
Dim titlerow As Long
Dim xTRg As Range
Dim xVRg As Range
Dim xWSTRg As Worksheet
On Error Resume Next

Set xTRg = Application.InputBox("Please select the header rows:", "Kutools for Excel", "", Type:=8)
If TypeName(xTRg) = "Nothing" Then Exit Sub
Set xVRg = Application.InputBox("Please select the column you want to split data based on:", "Kutools for Excel", "", Type:=8)


If TypeName(xVRg) = "Nothing" Then Exit Sub
Application.ScreenUpdating = False
vcol = xVRg.Column

Set ws = xTRg.Worksheet

lr = ws.Cells(ws.Rows.Count, vcol).End(xlUp).Row
title = xTRg.AddressLocal
titlerow = xTRg.Cells(1).Row
icol = ws.Columns.Count
ws.Cells(1, icol) = "Unique"
Application.DisplayAlerts = False
If Not Evaluate("=ISREF('xTRgWs_Sheet!A1')") Then
Sheets.Add(after:=Worksheets(Worksheets.Count)).Name = "xTRgWs_Sheet"
Else
Sheets("xTRgWs_Sheet").Delete
Sheets.Add(after:=Worksheets(Worksheets.Count)).Name = "xTRgWs_Sheet"
End If
Set xWSTRg = Sheets("xTRgWs_Sheet")
xTRg.Copy
xWSTRg.Paste Destination:=xWSTRg.Range("A1")
ws.Activate
For i = (titlerow + xTRg.Rows.Count) To lr
On Error Resume Next
If ws.Cells(i, vcol) <> "" And Application.WorksheetFunction.Match(ws.Cells(i, vcol), ws.Columns(icol), 0) = 0 Then
ws.Cells(ws.Rows.Count, icol).End(xlUp).Offset(1) = ws.Cells(i, vcol)
End If
Next
myarr = Application.WorksheetFunction.Transpose(ws.Columns(icol).SpecialCells(xlCellTypeConstants))
ws.Columns(icol).Clear
For i = 2 To UBound(myarr)
ws.Range(title).AutoFilter field:=vcol, Criteria1:=myarr(i) & ""
If Not Evaluate("=ISREF('" & myarr(i) & "'!A1)") Then
Sheets.Add(after:=Worksheets(Worksheets.Count)).Name = myarr(i) & ""
Else
Sheets(myarr(i) & "").Move after:=Worksheets(Worksheets.Count)
End If
xWSTRg.Range(title).Copy
Sheets(myarr(i) & "").Paste Destination:=Sheets(myarr(i) & "").Range("A1")
ws.Range("A" & (titlerow + xTRg.Rows.Count) & ":A" & lr).EntireRow.Copy Sheets(myarr(i) & "").Range("A" & (titlerow + xTRg.Rows.Count))
Sheets(myarr(i) & "").Columns.AutoFit
Next
xWSTRg.Delete
ws.AutoFilterMode = False
ws.Activate
Application.DisplayAlerts = True
Application.ScreenUpdating = True

MsgBox "Completed successfully."

Else
End If
ER:

End Sub
