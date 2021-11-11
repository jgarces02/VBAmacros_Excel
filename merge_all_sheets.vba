'option 1 --> sheets must to have the same structure
Sub Combine()
Dim J As Integer
On Error Resume Next
Sheets(1).Select
Worksheets.Add
Sheets(1).Name = "Combined"
Sheets(2).Activate
Range("A1").EntireRow.Select 'Aquí puedes cambiar la celda desde donde empieza a unir los datos'
Selection.Copy Destination:=Sheets(1).Range("A1")
For J = 2 To Sheets.Count
Sheets(J).Activate
Range("A1").Select 'cuidado, también habría que cambiar aquí la celda!'
Selection.CurrentRegion.Select
Selection.Offset(1, 0).Resize(Selection.Rows.Count - 1).Select
Selection.Copy Destination:=Sheets(1).Range("A65536").End(xlUp)(2)
Next
End Sub

'https://www.extendoffice.com/excel/1184-excel-merge-multiple-worksheets-into-one.html?page_comment=26&PageSpeed=noscript

'-------------------------------------------
'option 2 --> same structure is not needed
Sub Merge_Sheets()
    'Insert a new worksheet
    Sheets.Add
    
    'Rename the new worksheet
    ActiveSheet.Name = "Combined"
    
    'Loop through worksheets and copy the to your new worksheet
    For Each ws In Worksheets
        ws.Activate
        
        'Don't copy the merged sheet again
        If ws.Name <> "Combined" Then
            ws.UsedRange.Select
            Selection.Copy
            Sheets("Combined").Activate
            
            'Select the last filled cell
            ActiveSheet.Range("A1048576").Select
            Selection.End(xlUp).Select
            
            'For the first worksheet you don't need to go down one cell
            If ActiveCell.Address <> "$A$1" Then
                ActiveCell.Offset(1, 0).Select
            End If
            
            'Instead of just paste, you can also paste as link, as values etc.
            ActiveSheet.Paste
        
        End If
        
    Next
End Sub

'https://professor-excel.com/merge-sheets/
