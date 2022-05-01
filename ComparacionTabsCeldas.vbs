Public Sub SeleccionarArchivoViejo()
    'variables
    Dim dialogBox As FileDialog
    Set dialogBox = Application.FileDialog(msoFileDialogOpen)
    
    'No permitir que se seleccione mas de un archivo
    dialogBox.AllowMultiSelect = False
    
    'Pregunta del DialogBox
    MsgBox "Selecciona el archivo 'VIEJO' a comparar"
    dialogBox.Title = "Selecciona el archivo 'VIEJO' a comparar"
    
    If dialogBox.Show = -1 Then
        MsgBox "You selected: " & dialogBox.SelectedItems(1)
    End If
    
    'abrir Workbook
    Workbooks.Open (dialogBox.SelectedItems(1))
End Sub

Public Sub SeleccionarArchivoNuevo()
    'variables
    Dim dialogBox As FileDialog
    Set dialogBox = Application.FileDialog(msoFileDialogOpen)
    
    'No permitir que se seleccione mas de un archivo
    dialogBox.AllowMultiSelect = False
    
    'Pregunta del DialogBox
    MsgBox "Selecciona el archivo 'NUEVO' a comparar"
    dialogBox.Title = "Selecciona el archivo 'NUEVO' a comparar"
    
    If dialogBox.Show = -1 Then
        MsgBox "You selected: " & dialogBox.SelectedItems(1)
    End If
    
    'abrir Workbook
    Workbooks.Open (dialogBox.SelectedItems(1))
End Sub


Public Sub IterarHojas()
    Dim ws As Worksheet
    Dim i As Integer
    Dim j As Integer
    Dim HojaN As Integer
    Dim HojaV As Integer
    'Conteo de Hojas
    HojaN = Workbooks(3).Worksheets.Count
    HojaV = Workbooks(2).Worksheets.Count
    'Esto evita que se vaya actualzando la pantalla
    Application.ScreenUpdating = False
    'No se va a ejcutar nada del codigo si los dos libros no tienen le mismo numero de hojas
    'If (Workbooks(2).Worksheets.Count = Workbooks(3).Worksheets.Count) Then
        'i = 1
        For i = 1 To HojaV Step 1
            For j = 1 To HojaN Step 1
                'Hojas = i
                'se hace la comparacion de hoja solo si la hoja tiene ZDC en el nombre y si las hojas de ambos libros se llaman igual
                If (InStr(1, Workbooks(2).Worksheets(i).Name, "ZDC", vbTextCompare) > 0 _
                    And Workbooks(2).Worksheets(i).Name = Workbooks(3).Worksheets(j).Name) Then
                    'MsgBox "Comparando las hojas: " & Workbooks(2).Worksheets(i).Name
                    Call ComparacionHoja(i, j)
                    Exit For
                End If
                'i = i + 1
            Next j
        Next i
    'Else
        'MsgBox "Los dos libros tiene diferente numero de hojas"
    'End If
    'Se reactiva la actualizacion de pantalla
    Application.ScreenUpdating = True
End Sub

Public Sub ComparacionHoja(HojaV As Integer, HojaN As Integer)
    'maximo numero de columnas en blanco presentes en las hojas analizadas
    'aveces son utilizadas para separar info
    Dim filasUse As Integer
    Dim ColUse As Integer
    Dim i As Integer
    Dim j As Integer
    filasUse = Workbooks(2).Worksheets(HojaV).UsedRange.Rows.Count
    ColUse = Workbooks(2).Worksheets(HojaV).UsedRange.Columns.Count
    Dim filas As Integer
    Dim columnas As Integer
    'Se definen las coordenadas para inicio de la comparacion
    columnas = 1
    filas = 1
    Dim ncambios As Integer
    ncambios = 0
    'Se va a comparar columna a columna
    'Se iterara como matriz (fila, columna), por ejemplo, si la informacion comienza en C7 osea (7,3)
    For j = 1 To ColUse Step 1
        'If (IsEmpty(Workbooks(2).Worksheets(HojaV).Cells(filas, j)) = True) Then
         '   GoTo ContinueDo
        'End If
        'Las filas se comienzan a contar desde 7 cada que se cambia una columna
        'Se deja de iterar sobre las filas cuando la celda se vuelve vacia
        For i = 1 To filasUse Step 1
            'Si las celdas son diferentes, se rellena en rojo la celda distinta en el libro nuevo
            If (Workbooks(2).Worksheets(HojaV).Cells(i, j).Value <> Workbooks(3).Worksheets(HojaN).Cells(i, j).Value) Then
                Workbooks(3).Worksheets(HojaN).Cells(i, j).Interior.Color = RGB(255, 0, 255)
                ncambios = ncambios + 1
            End If
        Next i
'ContinueDo:
        'Se selecciona la siguiente columna
    Next j
    If ncambios > 0 Then
        'MsgBox "Se realizaron " & ncabmios & " en la pestaña " & Workbooks(3).Worksheets(Hojas).Name
        Workbooks(3).Worksheets(HojaN).Tab.Color = RGB(0, 0, 0)
    End If
End Sub

Public Sub Completo()
    SeleccionarArchivoViejo
    SeleccionarArchivoNuevo
    'Workbooks(1).Worksheets("Hoja1").Range("A1").Select
    ConteoTabs
    IterarHojas
    MsgBox "Las celdas distintas fueron marcadas en el archivo " & Workbooks(3).Name
    Workbooks(1).Worksheets("Hoja1").Activate
End Sub

Public Sub ConteoTabs()
    'Variables
    Dim HojaN As Integer
    Dim HojaV As Integer
    Dim indicador As Integer
    'Conteo de Hojas
    HojaN = Workbooks(3).Worksheets.Count
    HojaV = Workbooks(2).Worksheets.Count
    'Aviso de comparacion de Hojas
    If (HojaN = HojaV) Then
        MsgBox "Los archivos tiene el mismo número de Hojas"
    End If
    If (HojaN < HojaV) Then
        MsgBox "El archivo 'VIEJO' tiene " & (HojaV - HojaN) & " más que el 'NUEVO'"
    End If
    If (HojaN > HojaV) Then
        MsgBox "El archivo 'NUEVO' tiene " & (HojaN - HojaV) & " más que el 'VIEJO'"
    End If
    
    'Verificacion para archivo viejo
    For i = 1 To HojaV Step 1
    indicador = 0
        For j = 1 To HojaN Step 1
            If (Workbooks(2).Worksheets(i).Name = Workbooks(3).Worksheets(j).Name) Then
                indicador = 1
                Exit For
            End If
        Next j
        If indicador = 0 Then
            'MsgBox Workbook(2).Worksheets(i).Name & " no se encuentra en el archivo 'NUEVO'"
            Workbooks(2).Worksheets(i).Tab.Color = RGB(0, 128, 128)
        End If
    Next i
    
    'verificacion para archivo nuevo
    For i = 1 To HojaN Step 1
    indicador = 0
        For j = 1 To HojaV Step 1
            If (Workbooks(2).Worksheets(j).Name = Workbooks(3).Worksheets(i).Name) Then
                indicador = 1
                Exit For
            End If
        Next j
        If indicador = 0 Then
            'MsgBox Workbook(3).Worksheets(i).Name & " no se encuentra en el archivo 'VIEJO'"
            
            Workbooks(3).Worksheets(i).Tab.Color = RGB(0, 128, 128)
        End If
    Next i
    MsgBox "Las hojas ya no existentes en el archivo 'NUEVO' pero si en el 'VIEJO' fueron marcados de color azul verdoso"
    MsgBox "Las hojas no existentes en el archivo 'VIEJO' pero si en el 'NUEVO' fueron marcados de color azul verdoso"
End Sub

