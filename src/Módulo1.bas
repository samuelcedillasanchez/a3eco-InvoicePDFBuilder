Attribute VB_Name = "Módulo1"
Function Extraer_Hipervinculo(Rango As Range)
Dim Hipervinculo As String
Hipervinculo = Rango.Hyperlinks(1).Address
Extraer_Hipervinculo = Hipervinculo
End Function

