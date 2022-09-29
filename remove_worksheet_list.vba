Sub Deletenotinlist()
    Dim i As Long
    Dim cnt As Long
    Dim xWb, actWs As Worksheet
    Set actWs = ThisWorkbook.ActiveSheet
    cnt = 0
    Application.DisplayAlerts = False
    For i = Sheets.Count To 1 Step -1
        If Not ThisWorkbook.Sheets(i) Is actWs Then
            xWb = Application.Match(Sheets(i).Name, actWs.Range("A2:A6"), 0)
            If IsError(xWb) Then
                ThisWorkbook.Sheets(i).Delete
                cnt = cnt + 1
            End If
        End If
    Next
    Application.DisplayAlerts = True
    If cnt = 0 Then
        MsgBox "Not find the sheets to be seleted", vbInformation, "Kutools for Excel"
    Else
        MsgBox "Have deleted" & cnt & "worksheets"
    End If
End Sub

'https://www.extendoffice.com/documents/excel/4093-excel-delete-sheet-if-not-in-list.html
