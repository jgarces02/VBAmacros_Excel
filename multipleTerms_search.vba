Option Explicit

Sub ListHits()
    Dim Cell2Find As Range
    Dim FoundCell As Range
    Dim FirstFound As Range
    Dim Ct As Long
    Dim KeepLooking As Boolean
    'here you must rename the sheet with multiple search elements or adjust original name
    For Each Cell2Find In Worksheets("Keywords").UsedRange
        If Len(Cell2Find.Value) > 0 Then
            Set FirstFound = Nothing
            'the same for sheet with searchable data...
            With Worksheets("Data")
                Ct = 0
                Set FoundCell = Nothing
                Set FoundCell = .UsedRange.Find(What:=Cell2Find.Value, After:=.Range("A1"), LookIn:=xlValues, _
                                                LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                                MatchCase:=False, SearchFormat:=False)

                If FoundCell Is Nothing Then
                    KeepLooking = False
                Else
                    Set FirstFound = FoundCell
                    KeepLooking = True
                End If
                Do While KeepLooking
                    Ct = Ct + 1
                    Cell2Find.Offset(, Ct).Hyperlinks.Add Cell2Find.Offset(, Ct), "", FoundCell.Address(external:=True), FoundCell.Address(external:=True)
                    'Set FoundCell = Nothing
                    Set FoundCell = .UsedRange.FindNext(FoundCell)
                    If FoundCell Is Nothing Then KeepLooking = False
                    If FoundCell.Address = FirstFound.Address Then KeepLooking = False
                Loop
            End With
        End If
    Next
End Sub

'https://techcommunity.microsoft.com/t5/excel/how-do-i-find-multiple-search-terms-in-excel/m-p/1307297
