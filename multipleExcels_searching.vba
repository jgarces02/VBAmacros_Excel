Sub SearchFolders()
    Dim fso As Object
    Dim fld As Object
    Dim strSearch As String
    Dim strPath As String
    Dim strFile As String
    Dim wOut As Worksheet
    Dim wbk As Workbook
    Dim wks As Worksheet
    Dim lRow As Long
    Dim rFound As Range
    Dim strFirstAddress As String

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    'Change as desired
    strPath = "c:\MyFolder"
    strSearch = "Specific text"

    Set wOut = Worksheets.Add
    lRow = 1
    With wOut
        .Cells(lRow, 1) = "Workbook"
        .Cells(lRow, 2) = "Worksheet"
        .Cells(lRow, 3) = "Cell"
        .Cells(lRow, 4) = "Text in Cell"
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set fld = fso.GetFolder(strPath)

        strFile = Dir(strPath & "\*.xls*")
        Do While strFile <> ""
            Set wbk = Workbooks.Open _
              (Filename:=strPath & "\" & strFile, _
              UpdateLinks:=0, _
              ReadOnly:=True, _
              AddToMRU:=False)

            For Each wks In wbk.Worksheets
                Set rFound = wks.UsedRange.Find(strSearch)
                If Not rFound Is Nothing Then
                    strFirstAddress = rFound.Address
                End If
                Do
                    If rFound Is Nothing Then
                        Exit Do
                    Else
                        lRow = lRow + 1
                        .Cells(lRow, 1) = wbk.Name
                        .Cells(lRow, 2) = wks.Name
                        .Cells(lRow, 3) = rFound.Address
                        .Cells(lRow, 4) = rFound.Value
                    End If
                    Set rFound = wks.Cells.FindNext(After:=rFound)
                Loop While strFirstAddress <> rFound.Address
            Next

            wbk.Close (False)
            strFile = Dir
        Loop
        .Columns("A:D").EntireColumn.AutoFit
    End With
    MsgBox "Done"

ExitHandler:
    Set wOut = Nothing
    Set wks = Nothing
    Set wbk = Nothing
    Set fld = Nothing
    Set fso = Nothing
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbExclamation
    Resume ExitHandler
End Sub

'https://excel.tips.net/T005598_Searching_Through_Many_Workbooks.html
