Attribute VB_Name = "Módulo2"
Sub ReemplazarNombre()

    Dim NombreActual As String
    Dim NombreNuevo As String
    Dim RutaOriginal As String
    Dim ExtensionArchivo As String
    Dim NuevaRuta As String
    Dim Celda As Range
    Dim Hoja As Worksheet
    Dim Fila As Long

    Set Hoja = Worksheets("Facturas")
    Fila = 9

    Do While Hoja.Cells(Fila, "N").Value <> ""

        NombreActual = Hoja.Cells(Fila, "N").Value
        NombreNuevo = Hoja.Cells(Fila, "O").Value

        ' Obtener ruta y extensión
        RutaOriginal = Left(NombreActual, InStrRev(NombreActual, "\"))
        ExtensionArchivo = Mid(NombreActual, InStrRev(NombreActual, ".") + 1)
        
        ' Si no hay extensión válida, saltamos
        If ExtensionArchivo = "" Then
            Fila = Fila + 1
            GoTo Continuar
        End If
        
        ' Construir nueva ruta con misma carpeta y misma extensión
        NuevaRuta = RutaOriginal & NombreNuevo & "." & ExtensionArchivo

        ' Intentar renombrar
        On Error Resume Next
        Name NombreActual As NuevaRuta
        On Error GoTo 0

Continuar:
        Fila = Fila + 1

    Loop

    MsgBox "Renombrado finalizado."

End Sub

