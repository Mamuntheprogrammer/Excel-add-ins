Attribute VB_Name = "Match_multiple_values"
'Callback for customButton2 onAction
Sub vMa(control As IRibbonControl)
Call matchvalues
End Sub
Sub matchvalues()
question = "Warning : Once you done can't undo the task,Please backup your main file" & vbCrLf & vbCrLf & "Are you sure you want to run this Macro? "
If MsgBox(question, vbYesNo + vbQuestion) = vbYes Then
On Error GoTo ER
   Dim nA As Long, nB As Long, v As Variant
   Dim a As Long, b As Long
   Dim valtoFind As Range
   Dim valWhertoFind As Range
   Dim vcol As Integer
   Dim xcol As Integer
   Dim cel As Range
   Dim dcel As Range
   
   Dim xWSTRg As Worksheet



   Set valtoFind = Application.InputBox("Please select values you want to match:", "PyGems", "", Type:=8)
   Set valWhertoFind = Application.InputBox("Please select column where you want to match:", "PyGems", "", Type:=8)


   vcol = valtoFind.Column
   Set ws = valtoFind.Worksheet

   xcol = valWhertoFind.Column
   Set wws = valWhertoFind.Worksheet

For Each cel In valtoFind.Cells
    v = cel.Value
    For Each dcel In valWhertoFind.Cells
        If dcel.Value = v Then
        dcel.Interior.ColorIndex = 6
                cel.Interior.ColorIndex = 4
    

         End If
      Next dcel
   Next cel




   nA = ws.Cells(ws.Rows.Count, vcol).End(xlUp).Row

   nB = wws.Cells(wws.Rows.Count, xcol).End(xlUp).Row



   For a = 1 To nA
      v = Cells(a, "A").Value
      For b = 1 To nB
         If Cells(b, "B").Value = v Then
            Cells(b, "B").Interior.ColorIndex = 6
            Cells(a, "A").Interior.ColorIndex = 6
    
         End If
      Next b
   Next a
MsgBox "Completed successfully."
Else
End If
ER:

End Sub

