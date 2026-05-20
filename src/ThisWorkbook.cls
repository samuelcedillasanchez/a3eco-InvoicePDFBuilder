VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim nombreCompleto As String
    Dim fechaHora As String
    
    nombreCompleto = ObtenerNombreCompleto()
    fechaHora = Format(Now, "hh:mm:ss AM/PM") ' Aqu� usamos Now para obtener solo la hora
    
    ' Concatenar la fecha, nombre y hora en el formato deseado
    ThisWorkbook.Sheets("Facturas").Range("A1").Value = "�ltima modificaci�n: " & Format(Now, "dd/mm/yyyy") & " por " & nombreCompleto & " a las " & fechaHora
End Sub




Private Sub Workbook_Open()
    ' Llama a la subrutina CopiarDatos cuando se abre el archivo vac�o
    Call CopiarDatos
    Call ConvertTIFFsToPDF_WordInsertImage
    Call ReemplazarNombre
End Sub
