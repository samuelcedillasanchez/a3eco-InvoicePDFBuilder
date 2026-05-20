Attribute VB_Name = "Módulo3"
Function GetFullPath(filePath As String) As String
    If Left(filePath, 2) = "\\" Or Mid(filePath, 2, 1) = ":" Then
        ' Ya es ruta absoluta (UNC o con unidad)
        GetFullPath = filePath
    Else
        ' Ruta relativa -> concatena con carpeta del libro Excel
        GetFullPath = ThisWorkbook.Path & Application.PathSeparator & filePath
    End If
End Function

Sub ConvertTIFFsToPDF_WordInsertImage()
    Dim filePath As String
    Dim pdfPath As String
    Dim wordApp As Object
    Dim doc As Object
    Dim extension As String
    Dim currentRow As Long

    Const START_ROW As Long = 9
    Const COLUMN_PATH As String = "K"
    Const COLUMN_RESULT As String = "R"

    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    If wordApp Is Nothing Then Set wordApp = CreateObject("Word.Application")
    On Error GoTo 0

    If wordApp Is Nothing Then
        MsgBox "Microsoft Word no está disponible.", vbCritical
        Exit Sub
    End If

    wordApp.Visible = False
    currentRow = START_ROW

    Do While Cells(currentRow, COLUMN_PATH).Value <> ""
        filePath = ""

        If Cells(currentRow, COLUMN_PATH).Hyperlinks.Count > 0 Then
            filePath = GetFullPath(Cells(currentRow, COLUMN_PATH).Hyperlinks(1).Address)
        End If

        If filePath = "" Then
            Cells(currentRow, COLUMN_RESULT).Value = "Sin hipervínculo"
            GoTo NextRow
        End If

        extension = LCase(Right(filePath, Len(filePath) - InStrRev(filePath, ".")))

        If extension = "pdf" Then
            Cells(currentRow, COLUMN_RESULT).Value = "Ya es PDF"
        ElseIf extension = "tif" Or extension = "tiff" Then
            pdfPath = Left(filePath, InStrRev(filePath, ".")) & "pdf"

            On Error Resume Next
            Set doc = wordApp.Documents.Add
            On Error GoTo 0

            If Not doc Is Nothing Then
                On Error Resume Next
                doc.InlineShapes.AddPicture filePath
                If Err.Number <> 0 Then
                    Cells(currentRow, COLUMN_RESULT).Value = "Error al insertar imagen"
                    doc.Close False
                    GoTo NextRow
                End If
                On Error GoTo 0

                On Error Resume Next
                doc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=17
                If Err.Number <> 0 Then
                    Cells(currentRow, COLUMN_RESULT).Value = "Error al exportar PDF"
                    doc.Close False
                    GoTo NextRow
                End If
                On Error GoTo 0

                doc.Close False
                Cells(currentRow, COLUMN_RESULT).Value = "Convertido a PDF"
            Else
                Cells(currentRow, COLUMN_RESULT).Value = "Error al crear doc Word"
            End If
        Else
            Cells(currentRow, COLUMN_RESULT).Value = "Formato no válido"
        End If

NextRow:
        currentRow = currentRow + 1
    Loop

    wordApp.Quit
    Set wordApp = Nothing

    MsgBox "Conversión finalizada."
End Sub

