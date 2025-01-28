Sub ActualizarOrigenDatosTablaDinamica(libro As String, hoja As String, celdaTablaDinamica As String, rangoOrigenDatos As String, password As String)
    ' Declarar variables
    Dim xlLibro As Workbook
    Dim xlHoja As Worksheet
    Dim xlTablaDinamica As PivotTable
    Dim xlApp As Object
    
    ' Crear una instancia de Excel
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False ' Mantener Excel oculto durante la ejecución
    
    ' Abrir el libro especificado en modo de solo lectura
    Set xlLibro = xlApp.Workbooks.Open(Filename:=libro, Password:=password, ReadOnly:=True)
    Set xlHoja = xlLibro.Sheets(hoja)
    
    ' Obtener la tabla dinámica
    Set xlTablaDinamica = xlHoja.Range(celdaTablaDinamica).PivotTable
    
    ' Actualizar el origen de datos de la tabla dinámica
    xlTablaDinamica.ChangePivotCache xlLibro.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rangoOrigenDatos)
    
    ' Actualizar la tabla dinámica
    xlTablaDinamica.RefreshTable
    
    ' Guardar y cerrar el libro
    xlLibro.Save
    xlLibro.Close
    xlApp.Quit
    
    ' Liberar objetos
    Set xlTablaDinamica = Nothing
    Set xlHoja = Nothing
    Set xlLibro = Nothing
    Set xlApp = Nothing
End Sub

' Ejemplo de uso
Sub LlamarActualizarOrigenDatosTablaDinamica()
    Call ActualizarOrigenDatosTablaDinamica( _
        "\\boinfii10d09\RepositorioAA\R_RPAOPM-197_LiquidacionDeIncentivos\RutaCompartida\clusterwinfs2fs\Gerencia_BancaSeguros\incentivos\Recaudos JUNIO 2024 Final.xlsx", _
        "TABLA", _
        "A3", _
        "RECAUDOS!A:BX", _
        "tu_contraseña" _
    )
End Sub