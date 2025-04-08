Sub AddBorders(sheetName As String, startCell As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    Dim firstCell As Range
    Set firstCell = ws.Range(startCell) ' Celda inicial
    
    ' Encontrar la última fila con datos en la columna de la celda inicial
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, firstCell.Column).End(xlUp).Row
    
    ' Encontrar la última columna con datos en la fila de la celda inicial
    Dim lastCol As Long
    lastCol = ws.Cells(firstCell.Row, ws.Columns.Count).End(xlToLeft).Column
    
    ' Crear el rango completo desde la celda inicial hasta la última celda con datos
    Dim tableRange As Range
    Set tableRange = ws.Range(firstCell, ws.Cells(lastRow, lastCol))
    
    ' Aplicar bordes de color azul (sin diagonales)
    Dim bordersArray As Variant
    bordersArray = Array(xlEdgeLeft, xlEdgeRight, xlEdgeTop, xlEdgeBottom, xlInsideVertical, xlInsideHorizontal)
    
    Dim borderType As Variant
    For Each borderType In bordersArray
        With tableRange.Borders(borderType)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(48, 84, 150) ' Azul
        End With
    Next borderType
    
    ' Centrar contenido horizontal y verticalmente
    With tableRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Sub FormatExcel(sheetName As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    
    AddBorders sheetName, "B7"
    
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    
    ws.Range("E7:E" & lastRow).HorizontalAlignment = xlLeft
    
    ws.Range("G7:G" & lastRow).HorizontalAlignment = xlRight ' Valor factura
    ws.Range("J7:J" & lastRow).HorizontalAlignment = xlRight ' Total pagado
    ws.Range("K7:K" & lastRow).HorizontalAlignment = xlRight ' Total de saldo
    
    ws.Range("H8:H" & lastRow).Font.Color = RGB(255, 0, 0)
    
    Set ws = Nothing
End Sub