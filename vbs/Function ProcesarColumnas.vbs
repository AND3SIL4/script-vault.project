
'Script para remplazar todo lo que sea diferente a numeros en un columna, 
'El caso de uso, La columna contenía datos numericos. EJ "99812Aaa" La forma de revolver es dejando la columna, 
'Unicamente con numeros, sin datos como letras, puntos o comas. 
Function ProcesarColumnas(libro, hoja, columnas)
    Dim objExcel, objWorkbook, objWorksheet
    Dim regex, celda, columna, fila
    Dim mensajeSalida

    On Error Resume Next ' Habilitar manejo de errores

    ' Crear objeto de Excel
    Set objExcel = CreateObject("Excel.Application")
    If Err.Number <> 0 Then
        ProcesarColumnas = "Error al iniciar Excel: " & Err.Description
        Exit Function
    End If

    objExcel.Visible = False
    Set objWorkbook = objExcel.Workbooks.Open(libro)
    If Err.Number <> 0 Then
        ProcesarColumnas = "Error al abrir el libro: " & Err.Description
        Exit Function
    End If

    Set objWorksheet = objWorkbook.Sheets(hoja)
    If Err.Number <> 0 Then
        ProcesarColumnas = "Error al acceder a la hoja: " & Err.Descrgit iption
        Exit Function
    End If

    ' Crear la expresión regular
    Set regex = New RegExp
    regex.Pattern = "[^\d]" ' Coincide con todo lo que no sea un dígito (0-9)
    regex.Global = True

    ' Procesar cada columna y cada fila de la hoja
    For Each columna In columnas
        For fila = 1 To objWorksheet.UsedRange.Rows.Count ' Recorrer todas las filas de la columna
            Set celda = objWorksheet.Range(columna & fila) ' Obtener la celda
            If Not IsEmpty(celda.Value) Then
                ' Reemplazar todo lo que no sea números
                celda.Value = regex.Replace(celda.Value, "")
            End If
        Next
    Next

    ' Guardar y cerrar el archivo
    objWorkbook.Save
    objWorkbook.Close False
    objExcel.Quit

    If Err.Number <> 0 Then
        ProcesarColumnas = "Error al guardar o cerrar el libro: " & Err.Description
    Else
        ProcesarColumnas = "El procesamiento ha finalizado correctamente."
    End If

    ' Liberar objetos
    Set objWorksheet = Nothing
    Set objWorkbook = Nothing
    Set objExcel = Nothing
    Set regex = Nothing
End Function

' Ejecutar la función
Dim libro, hoja, columnas, resultado
libro = WScript.Arguments(0) ' Ruta del libro de Excel
hoja = WScript.Arguments(1) ' Nombre de la hoja
columnas = Split(WScript.Arguments(2), ",") ' Columnas a procesar (separadas por coma, ejemplo: "A,B,C")

resultado = ProcesarColumnas(libro, hoja, columnas)
WScript.Echo resultado
