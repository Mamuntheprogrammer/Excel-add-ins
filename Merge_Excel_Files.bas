Attribute VB_Name = "Merge_Excel_Files"
'Callback for customButton6 onAction
Sub fwMe(control As IRibbonControl)
Call mergeFiles
End Sub
Sub mergeFiles()
question = "Warning : Once you done can't undo the task,Please backup your main file" & vbCrLf & vbCrLf & "Are you sure you want to run this Macro? "
If MsgBox(question, vbYesNo + vbQuestion) = vbYes Then
On Error GoTo ER
    Dim FileFold As String
    Dim FileSpec As String
    Dim FileName As String
    Dim ShtCnt As Long
    Dim RowCnt As Long
    Dim Merged As Workbook
    Dim wb As Workbook
    Dim ws As Worksheet
    
    
    FileFold = Application.InputBox("Please Enter The Folder Path")
    


    FileSpec = FileFold & Application.PathSeparator & "*.xlsx*"
    FileName = Dir(FileSpec)
    
    'Exit if no files found
    If FileName = vbNullString Then
        MsgBox Prompt:="No files were found that match " & FileSpec, Buttons:=vbCritical, title:="Error"
        Exit Sub
    End If
    
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
    End With
    
    ShtCnt = 0
    RowCnt = 1
    
    Set Merged = Workbooks.Add
    
    Do While FileName <> vbNullString
        ShtCnt = ShtCnt + 1
        Set wb = Workbooks.Open(FileName:=FileFold & Application.PathSeparator & FileName, UpdateLinks:=False)
        Set ws = wb.Worksheets("Sheet1")
        With ws
            If .FilterMode Then .ShowAllData
            If ShtCnt > 1 Then .Rows(1).EntireRow.Delete Shift:=xlUp
            .Range("A1").CurrentRegion.Copy Destination:=Merged.Worksheets(1).Cells(RowCnt, 1)
        End With
        wb.Close SaveChanges:=False
        RowCnt = Application.WorksheetFunction.CountA(Merged.Worksheets(1).Columns("A:A")) + 1
        FileName = Dir
    Loop
    
    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
    End With
    
MsgBox "Completed successfully."

Else
End If
ER:

End Sub


