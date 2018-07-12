# VBA
VBA related codes collection

#查找并返回第一个空行
Function getEmptyRow(sheetName As String, col As Long) As Long
    Dim rng As Range
    Set rng = ThisWorkbook.Sheets(sheetName).Cells(65536, col).End(xlUp)
    getEmptyRow = rng.Row + 1
    Set rng = Nothing
End Function

Function test()
    MsgBox getEmptyRow("sheetname", 1)'指定表的名称，指定列号
End Function
