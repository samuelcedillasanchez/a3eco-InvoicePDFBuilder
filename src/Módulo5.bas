Attribute VB_Name = "Módulo5"
Sub CopiarDatos()
    Dim carpeta As String
    Dim archivoVacio As String
    Dim archivoConDatos As String
    Dim wbVacio As Workbook
    Dim wbConDatos As Workbook
    Dim wsVacio As Worksheet
    Dim wsConDatos As Worksheet
    Dim archivo As Object  ' Cambiar a tipo Object
    Dim fso As Object
    Dim carpetaObj As Object
    Dim archivoEncontrado As Boolean
    Dim ultimaFila As Long
    Dim i As Long
    Dim j As Integer
    
    ' Establece la ruta de la carpeta tomando el valor de la celda L8
    carpeta = ThisWorkbook.Sheets(1).Range("L8").Value  ' Toma la ruta de la celda L8
    
    ' Verifica que la ruta no esté vacía y termine con una barra invertida
    If carpeta = "" Then
        MsgBox "La ruta en la celda L8 está vacía. Por favor, ingresa una ruta válida."
        Exit Sub
    End If
    
    ' Asegurarse de que la ruta termine con una barra invertida
    If Right(carpeta, 1) <> "\" Then
        carpeta = carpeta & "\"
    End If
    
    ' Crea el objeto FileSystemObject para gestionar archivos
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set carpetaObj = fso.GetFolder(carpeta)
    
    archivoEncontrado = False
    
    ' Recorre todos los archivos de la carpeta
    For Each archivo In carpetaObj.Files
        If InStr(archivo.Name, ".xls") > 0 Then  ' Filtra archivos .xls
            If Not archivoEncontrado Then
                archivoConDatos = archivo.Name
                archivoEncontrado = True
            Else
                archivoVacio = archivo.Name
                Exit For
            End If
        End If
    Next archivo
    
    ' Si se encuentran los dos archivos (uno vacío y otro con datos)
    If archivoVacio <> "" And archivoConDatos <> "" Then
        ' Abre el archivo vacío y el archivo con datos
        Set wbConDatos = Workbooks.Open(carpeta & archivoConDatos)
        Set wbVacio = Workbooks.Open(carpeta & archivoVacio)
        
        ' Asumiendo que los datos están en la primera hoja de cada archivo
        Set wsConDatos = wbConDatos.Sheets(1)
        Set wsVacio = wbVacio.Sheets(1)
        
        ' Copiar datos de A3, A4, A5, A6
        wsVacio.Range("A3").Value = wsConDatos.Range("A3").Value  ' Copia celda A3
        wsVacio.Range("A4").Value = wsConDatos.Range("A4").Value  ' Copia celda A4
        wsVacio.Range("A5").Value = wsConDatos.Range("A5").Value  ' Copia celda A5
        wsVacio.Range("A6").Value = wsConDatos.Range("A6").Value  ' Copia celda A6
        
        ' Encontrar la última fila con datos en la columna A
        ultimaFila = wsConDatos.Cells(wsConDatos.Rows.Count, "A").End(xlUp).Row
        
        ' Copiar datos desde A9 hasta K (última fila)
        For i = 9 To ultimaFila
            For j = 1 To 11  ' Columnas A a K (1 a 11)
                ' Copiar el valor de la celda (texto/número)
                wsVacio.Cells(i, j).Value = wsConDatos.Cells(i, j).Value
                
                ' Copiar el hipervínculo si existe
                If wsConDatos.Cells(i, j).Hyperlinks.Count > 0 Then
                    wsVacio.Hyperlinks.Add _
                        Anchor:=wsVacio.Cells(i, j), _
                        Address:=wsConDatos.Cells(i, j).Hyperlinks(1).Address, _
                        TextToDisplay:=wsConDatos.Cells(i, j).Hyperlinks(1).TextToDisplay
                End If
            Next j
        Next i
        
        ' No cerramos ni guardamos el archivo vacío, solo cerramos el archivo con datos
        wbConDatos.Close
        
        ' Mensaje indicando que los datos se han copiado
        MsgBox "Los datos se han copiado exitosamente. Ahora puedes continuar trabajando en el archivo vacío."
    Else
        MsgBox "No se han encontrado los archivos requeridos en la carpeta."
    End If
End Sub

