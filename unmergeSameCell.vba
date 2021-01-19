Sub UnMergeSameCell()
Dim Rng As Range, xCell As Range
xTitleId = "Select a range to unmerge and fill with duplicates"
Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
Application.ScreenUpdating = False
Application.DisplayAlerts = False
For Each Rng In WorkRng
    If Rng.MergeCells Then
        With Rng.MergeArea
            .UnMerge
            .Formula = Rng.Formula
        End With
    End If
Next
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
