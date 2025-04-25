Sub AutoFillColumns2(filePath As String, sheetName As String, lastRow As Long)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim columnsToFill As Variant
    Dim col As Variant
    Dim colIndex As Long
    Dim wbAlreadyOpen As Boolean
    
    ' Define las columnas a realizar AutoFill (convertimos las letras a números)
    columnsToFill = Array("B", "AJ", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ", "EA", "EB", "EC", "ED", "EE", "EF", "EG", "EH", "EI", "EJ", "EK", "EL", "EM", "EN", "EO", "EP", "EQ", "ER", "EV")
    
    ' Verifica si el archivo ya está abierto
    wbAlreadyOpen = False
    On Error Resume Next
    Set wb = Workbooks(filePath)
    If wb Is Nothing Then
        ' Si el libro no está abierto, lo abre
        Set wb = Workbooks.Open(filePath)
    Else
        wbAlreadyOpen = True
    End If
    On Error GoTo 0
    
    ' Selecciona la hoja de trabajo
    Set ws = wb.Sheets(sheetName)
    
    ' Realiza el AutoFill en cada columna especificada
    For Each col In columnsToFill
        colIndex = ws.Columns(col).Column ' Convierte la letra de columna en número
        
        ' Verifica si la celda tiene datos para hacer el AutoFill
        If ws.Cells(2, colIndex).Value <> "" Then
            ws.Range(ws.Cells(2, colIndex), ws.Cells(lastRow, colIndex)).FillDown
        End If
    Next col
    
    ' Guarda y cierra el libro solo si fue abierto por la macro
    If Not wbAlreadyOpen Then
        wb.Close SaveChanges:=True
    End If
End Sub

Sub EjecutarAutoFill()
    Dim rutaArchivo As String
    Dim nombreHoja As String
    Dim ultimaFila As Long
    
    ' Especifica la ruta del archivo y el nombre de la hoja
    rutaArchivo = "\\boinfii10d09\RepositorioAA\R_RPAOPM-39_AgirVentMensSegPagoIncentivosComerciales\RutUsuario\clusterwinfs2fs\OBM_Banca_Minorista\Operaciones_Bancaseguros\2024\1. Radicación\10. Octubre\000MATRIZ DE RADICACIONES OCTUBRE.xlsx"
    nombreHoja = "PLANILLA"
    
    ' Especifica la última fila hasta donde quieres que llegue el AutoFill
    ultimaFila = 500 ' Cambia este valor según lo necesites
    
    ' Llama a la función AutoFillColumns2 con los parámetros especificados
    AutoFillColumns2 rutaArchivo, nombreHoja, ultimaFila
End Sub 