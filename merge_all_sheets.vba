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
'Sheets must to have the same structure!
