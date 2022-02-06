Attribute VB_Name = "Filter_Batch_Values"
Sub bFi(control As IRibbonControl)
Call batchFilter
End Sub

Sub batchFilter()

question = "Warning : Once you done can't undo the task,Please backup your main file" & vbCrLf & vbCrLf & "Are you sure you want to run this Macro? "
If MsgBox(question, vbYesNo + vbQuestion) = vbYes Then
On Error GoTo ER

    Dim v As Variant
    Dim val As Range
    Dim tb As Range
    Set val = Application.InputBox("Please select the values you want to filter  :", "PYGEMS", "", Type:=8)
    
    v = Application.Transpose(val)
    Set tb = Application.InputBox("Please select the single column data to be filtered :", "PYGEMS", "", Type:=8)
    tb.Select
    Selection.AutoFilter field:=1, Criteria1:=v, Operator:=xlFilterValues
MsgBox "Completed successfully."

Else
End If
ER:
End Sub
