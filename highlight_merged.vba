Sub ColorMergedCells()  
Dim c As Range  
For Each c In ActiveSheet.UsedRange  
If c.MergeCells Then  
c.Interior.ColorIndex = 28  
End If  
Next  
End Sub  

'https://www.exceltrick.com/how_to/find-merged-cells-in-excel/
