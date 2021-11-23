Sub ExportByCol()

Dim LastRow As Long, LastCol As Integer, i As Long, iStart As Long, iEnd As Long

Dim ws As Worksheet, r As Range, iCol As Integer, t As Date, Prefix As String

Dim sh As Worksheet, Master As String

On Error Resume Next

Set r = Application.InputBox("Click in the column to extract by", Type:=8)

On Error GoTo 0

If r Is Nothing Then Exit Sub

iCol = r.Column

t = Now

Application.ScreenUpdating = False

With ActiveSheet
    Master = .Name
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
    LastCol = .Cells(1, Columns.Count).End(xlToLeft).Column
    .Range(.Cells(2, 1), Cells(LastRow, LastCol)).Sort Key1:=Cells(2, iCol), Order1:=xlAscending, _
        Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    iStart = 2
    For i = 2 To LastRow
        If .Cells(i, iCol).Value <> .Cells(i + 1, iCol).Value Then
            iEnd = i
            Sheets.Add after:=Sheets(Sheets.Count)
            Set ws = ActiveSheet
            On Error Resume Next
            ws.Name = .Cells(iStart, iCol).Value
            On Error GoTo 0
            ws.Range(Cells(1, 1), Cells(1, LastCol)).Value = .Range(.Cells(1, 1), .Cells(1, LastCol)).Value
            .Range(.Cells(iStart, 1), .Cells(iEnd, LastCol)).Copy Destination:=ws.Range("A2")
            iStart = iEnd + 1
        End If
    Next i
End With

Application.CutCopyMode = False
Application.ScreenUpdating = True
MsgBox "Completed in " & Format(Now - t, "hh:mm:ss.00"), vbInformation
If MsgBox("Do you want to save the separated sheets as workbooks", vbYesNo + vbQuestion) = vbYes Then
    Prefix = InputBox("Enter a prefix (or leave blank)")
    Application.ScreenUpdating = False
    For Each sh In ThisWorkbook.Worksheets
        If sh.Name <> Master Then
            sh.Copy
            ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\" & Prefix & sh.Name & ".xls"
            ActiveWorkbook.Close
        End If
     Next sh
     Application.ScreenUpdating = True
End If
End Sub

