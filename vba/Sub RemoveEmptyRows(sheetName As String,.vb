Sub AgregarFilaAntesDeTotal()
    Dim ws As Worksheet
    Dim ultimaFila As Integer
    Dim i As Integer, j As Integer
    Dim filaTotal As Integer
    
    ' Establecer la hoja
    Set ws = ThisWorkbook.Sheets("FACTURA VENTA")
    
    ' Obtener la última fila con datos
    ultimaFila = ws.UsedRange.Rows.Count
    filaTotal = 0
    
    ' Buscar la fila que contiene "TOTAL CARTERA PENDIENTE:"
    For i = ultimaFila To 1 Step -1
        For j = 1 To ws.UsedRange.Columns.Count
            If InStr(1, Trim(ws.Cells(i, j).Value), "TOTAL CARTERA PENDIENTE:", vbTextCompare) > 0 Then
                filaTotal = i
                Exit For
            End If
        Next j
        If filaTotal > 0 Then Exit For
    Next i
    
    ' Si encontramos la fila, agregamos una antes con formato
    If filaTotal > 0 Then
        ' Insertar fila antes de TOTAL CARTERA PENDIENTE
        ws.Rows(filaTotal).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        
        ' Copiar el formato de la fila superior
        ws.Rows(filaTotal - 1).Copy
        ws.Rows(filaTotal).PasteSpecial Paste:=xlPasteFormats
        
        ' Limpiar portapapeles
        Application.CutCopyMode = False
        
    Else
        MsgBox "No se encontró la fila de TOTAL CARTERA PENDIENTE.", vbExclamation
    End If
End Sub

Sub LimpiarTabla1()
    Dim ws As Worksheet
    Dim tabla As ListObject
    Dim ultimaFila As Long
    Dim ultimaColumna As Long
    Dim rangoLimpiar As Range

    ' Establecer la hoja
    Set ws = ThisWorkbook.Sheets("FACTURA VENTA")

    ' Verificar si la tabla existe
    On Error Resume Next
    Set tabla = ws.ListObjects("Tabla1")
    On Error GoTo 0
    
    If tabla Is Nothing Then
      '  MsgBox "La tabla 'Tabla1' no existe en la hoja 'FACTURA VENTA'.", vbExclamation
        Exit Sub
    End If

    ' Encontrar la última fila de la tabla
    ultimaFila = tabla.ListRows.Count

    ' Si la tabla tiene menos de 2 filas de datos, no hacer nada
    If ultimaFila < 2 Then
       ' MsgBox "La tabla tiene menos de dos filas, no hay datos para limpiar.", vbExclamation
        Exit Sub
    End If
    
    ' Encontrar la última columna de la tabla
    ultimaColumna = tabla.Range.Columns.Count

    ' Definir el rango a limpiar (todas las filas excepto la última)
    Set rangoLimpiar = tabla.DataBodyRange.Resize(ultimaFila - 1)
    rangoLimpiar.ClearContents  ' Borra solo el contenido, mantiene formatos y estructura

   ' MsgBox "Datos de la tabla 'Tabla1' limpiados con éxito, manteniendo la cabecera y la última fila.", vbInformation
End Sub



Function GuardarArchivoConNombre() As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim nombreCliente As String
    Dim rutaCarpeta As String
    Dim rutaCompleta As String
    
    ' Definir el archivo activo
    Set wb = ActiveWorkbook
    Set ws = wb.Sheets(1) ' Asumimos que la tabla está en la primera hoja
    
    ' Buscar la Tabla1
    On Error Resume Next
    Set tbl = ws.ListObjects("Tabla1")
    On Error GoTo 0

    ' Verificar si la tabla existe
    If tbl Is Nothing Then
        MsgBox "La Tabla1 no fue encontrada en la hoja.", vbExclamation, "Error"
        Exit Function
    End If

    ' Obtener el nombre del cliente desde la primera fila de la columna "Nombre"
    nombreCliente = tbl.ListColumns("Nombre").DataBodyRange.Cells(1, 1).Value
    
    ' Si el nombre está vacío, mostrar error y salir
    If Trim(nombreCliente) = "" Then
        MsgBox "No se encontró un nombre válido en la Tabla1.", vbExclamation, "Error"
        Exit Function
    End If

    ' Asegurar que el nombre del cliente no tenga caracteres inválidos
    nombreCliente = LimpiarNombreArchivo(nombreCliente)
    
    ' Definir la carpeta donde se guardará el archivo
    rutaCarpeta = "C:\Users\duvan.sanabria\Downloads\Novedades Cartera\"
    
    ' Construir la ruta completa con el nuevo nombre
    rutaCompleta = rutaCarpeta & nombreCliente & ".xlsx"
    
    ' Guardar el archivo con el nombre del cliente
    wb.SaveAs Filename:=rutaCompleta, FileFormat:=xlOpenXMLWorkbook
    
    ' Retornar la ruta completa del archivo guardado
    GuardarArchivoConNombre = rutaCompleta
End Function

' Función para limpiar el nombre del archivo y evitar caracteres inválidos
Function LimpiarNombreArchivo(nombre As String) As String
    Dim caracteresInvalidos As String
    Dim i As Integer
    
    ' Caracteres no permitidos en nombres de archivos
    caracteresInvalidos = "\/:*?""<>|"
    
    ' Reemplazar caracteres inválidos por "_"
    For i = 1 To Len(caracteresInvalidos)
        nombre = Replace(nombre, Mid(caracteresInvalidos, i, 1), "_")
    Next i
    
    LimpiarNombreArchivo = nombre
End Function


Sub EliminarFilasSinFacturaExceptoUltima()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Integer
    Dim ultimaFila As Integer
    Dim colFactura As Integer

    ' Desactivar actualizaciones de pantalla y cálculos para mejorar rendimiento
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Definir la hoja y la tabla
    Set ws = ThisWorkbook.Sheets("FACTURA VENTA")
    Set tbl = ws.ListObjects("Tabla1")
    
    ' Obtener la columna "Factura"
    colFactura = tbl.ListColumns("Factura").Index
    
    ' Obtener la última fila de la tabla
    ultimaFila = tbl.ListRows.Count

    ' Recorrer la tabla de abajo hacia arriba para eliminar filas vacías sin tocar la última fila
    For i = ultimaFila - 1 To 1 Step -1 ' Evitamos borrar la última fila
        If IsEmpty(tbl.DataBodyRange.Rows(i).Cells(1, colFactura).Value) Then
            tbl.ListRows(i).Delete
        End If
    Next i

    ' Restaurar actualizaciones y cálculos
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    'MsgBox "Filas vacías eliminadas sin afectar la última fila.", vbInformation
End Sub


Sub EliminarUltimasFilasDespuesDeTabla()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim ultimaFilaTabla As Integer
    Dim filasAEliminar As Integer

    ' Definir la hoja y la tabla
    Set ws = ThisWorkbook.Sheets("FACTURA VENTA")
    Set tbl = ws.ListObjects("Tabla1")

    ' Obtener la última fila de la tabla
    ultimaFilaTabla = tbl.Range.Rows(tbl.Range.Rows.Count).Row

    ' Definir cuántas filas eliminar (5 filas después de la tabla)
    filasAEliminar = 5

    ' Verificar que haya suficiente espacio en la hoja para eliminar filas
    If ultimaFilaTabla + filasAEliminar <= ws.Rows.Count Then
        ws.Rows(ultimaFilaTabla + 1 & ":" & ultimaFilaTabla + filasAEliminar).Delete
    End If

    ' Mensaje de confirmación (opcional)
    'MsgBox "Se eliminaron las últimas 5 filas después de la tabla.", vbInformation
End Sub













