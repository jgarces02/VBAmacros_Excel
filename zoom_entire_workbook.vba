Sub DbZoom()
    Dim I As Long
    Dim xActSheet As Worksheet
    Set xActSheet = ActiveSheet
    For I = 1 To ThisWorkbook.Sheets.Count
        Sheets(I).Activate
        ActiveWindow.Zoom = 150 'change zoom level
    Next
    xActSheet.Select
End Sub

'https://www.extendoffice.com/documents/excel/4951-excel-zoom-all-tabs.html
