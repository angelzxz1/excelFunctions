Sub align()
'
' align Macro
'

'
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub

Sub Format()
'
' Format Macro
'

'
    Range("D3").Select
    Selection.Style = "Good"
    Range("E3").Select
    Selection.Style = "Neutral"
    Range("F3").Select
    Selection.Style = "Bad"
    Range("G3").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=33.4"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub

Sub CrearHojasYCalcularPorcentaje()
    Dim wsOriginal As Worksheet
    Dim wsNuevo As Worksheet
    Dim LastRow As Long
    Dim Agentes As New Collection
    Dim Agente As Variant
    Dim i As Long
    
    ' Definir la hoja original que contiene los datos
    Set wsOriginal = ThisWorkbook.Sheets("DSAT")
    
    ' Encontrar la última fila en la hoja original
    LastRow = wsOriginal.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Recorrer los datos y agregar agentes únicos a la colección
    On Error Resume Next
    For i = 2 To LastRow
        Agentes.Add wsOriginal.Cells(i, 2).Value, CStr(wsOriginal.Cells(i, 2).Value)
    Next i
    On Error GoTo 0
    
    ' Crear una hoja para cada agente en la colección
    For Each Agente In Agentes
        Set wsNuevo = ThisWorkbook.Sheets.Add
        wsNuevo.Name = Agente
        
        ' Copiar los encabezados
        wsOriginal.Rows(1).Copy wsNuevo.Rows(1)
        
        ' Filtrar y copiar las filas correspondientes al agente actual
        wsOriginal.Rows(1).AutoFilter Field:=2, Criteria1:=Agente
        wsOriginal.UsedRange.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Copy wsNuevo.Cells(2, 1)
        wsOriginal.AutoFilterMode = False
        
        ' Calcular el porcentaje de 1s en la columna "Rating" y colocarlo en la hoja
        wsNuevo.Cells(2, 5).Value = (WorksheetFunction.CountIf(wsNuevo.Columns(3), 1) / (wsNuevo.UsedRange.Rows.Count)) * 100
        wsNuevo.Cells(1, 5).Value = "DSAT"
        wsNuevo.Cells(1, 7).Value = "Total 'No' rated chats"
        wsNuevo.Cells(1, 8).Value = "Total rated chats"
        wsNuevo.Cells(2, 7).Value = WorksheetFunction.CountIf(wsNuevo.Columns(3), 1)
        wsNuevo.Cells(2, 8).Value = wsNuevo.UsedRange.Rows.Count
        wsNuevo.Cells(2, 8).Style = "Good"
        
    Next Agente
End Sub

