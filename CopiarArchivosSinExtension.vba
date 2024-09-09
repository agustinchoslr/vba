    ' Copiar archivos de una carpeta a otra mediante excel
Sub CopiarArchivosSinExtension()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim carpetaOrigen As String, carpetaDestino As String
    Dim nombreArchivoBase As String
    Dim fso As Object
    Dim file As Object
    Dim archivoEncontrado As Boolean

    Set wb = ThisWorkbook
    Set ws = ActiveSheet
    carpetaOrigen = "C:\Users\yo\Pictures\backup img\"
    carpetaDestino = "C:\Users\yo\Pictures\backup img\open\"

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Verifica si la carpeta destino existe, si no, la crea
    If Not fso.FolderExists(carpetaDestino) Then
        fso.CreateFolder carpetaDestino
    End If

    ' Asume que los nombres de archivo base (sin extensión) están en la columna A a partir de la fila 2
    For Each celda In ws.Range("A2:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)
        nombreArchivoBase = celda.Value
        archivoEncontrado = False

        ' Recorre todos los archivos en la carpeta origen
        For Each file In fso.GetFolder(carpetaOrigen).Files
            ' Compara el nombre del archivo sin la extensión
            If fso.GetBaseName(file.Name) = nombreArchivoBase Then
                file.Copy carpetaDestino & "\" & file.Name
                celda.Offset(0, 1).Value = "Éxito"
                archivoEncontrado = True
                Exit For
            End If
        Next file

        ' Si no se encontró el archivo
        If Not archivoEncontrado Then
            celda.Offset(0, 1).Value = "Archivo no encontrado"
        End If
    Next celda

    MsgBox "Proceso finalizado."
    Set fso = Nothing
End Sub

