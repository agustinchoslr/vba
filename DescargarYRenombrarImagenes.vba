    ' Descargar imágenes desde excel, con una columna con el nombre que tendrá el archivo, otra con la extensión y otra con la URL desde donde debe descargarse.
Sub DescargarYRenombrarImagenes()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim SKU As String
    Dim imageURL As String
    Dim EXTENS As String
    Dim destinationPath As String
    Dim objXMLHTTP As Object
    Dim objStream As Object
    
    ' Establecer la hoja de trabajo activa
    Set ws = ActiveSheet
    
    ' Encontrar la última fila con datos
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Crear objetos para descargar y guardar archivos
    Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")
    Set objStream = CreateObject("ADODB.Stream")
    
    ' Iterar a través de las filas
    For i = 2 To lastRow ' Asumiendo que la fila 1 es el encabezado
        ' Obtener SKU (nombre) y URL de la imagen
        SKU = ws.Cells(i, 1).Value ' Asumiendo que SKU está en la columna A
        imageURL = ws.Cells(i, 2).Value ' Asumiendo que la URL está en la columna B
        EXTENS = ws.Cells(i, 3).Value ' Extensión de la imagen en la columna C
        
        ' Definir la ruta de destino (ajusta según tus necesidades)
        destinationPath = "C:\Users\yo\Pictures\backup img\oppen\redketchup\" & SKU & EXTENS
        
        ' Descargar y guardar la imagen
        objXMLHTTP.Open "GET", imageURL, False
        objXMLHTTP.send
        
        If objXMLHTTP.Status = 200 Then
            objStream.Open
            objStream.Type = 1 'adTypeBinary
            objStream.Write objXMLHTTP.responseBody
            objStream.SaveToFile destinationPath, 2 'adSaveCreateOverWrite
            objStream.Close
            
            ' Opcional: Actualizar el estado en la hoja de cálculo
            ws.Cells(i, 4).Value = "Descargado" ' Asumiendo que usamos la columna D para el estado
        Else
            ' Si hay un error, actualizar el estado
            ws.Cells(i, 4).Value = "Error: " & objXMLHTTP.Status
        End If
    Next i
    
    ' Liberar objetos
    Set objStream = Nothing
    Set objXMLHTTP = Nothing
    
    MsgBox "Proceso completado"
End Sub
