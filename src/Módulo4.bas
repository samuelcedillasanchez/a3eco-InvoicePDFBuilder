Attribute VB_Name = "Módulo4"
Function ObtenerNombreCompleto() As String
    Dim objSysInfo As Object
    Dim objUser As Object

    On Error GoTo err_handler
    Set objSysInfo = CreateObject("ADSystemInfo")
    Set objUser = GetObject("LDAP://" & objSysInfo.UserName)
    
    ObtenerNombreCompleto = objUser.DisplayName
    Exit Function

err_handler:
    ObtenerNombreCompleto = Environ("USERNAME") ' Si falla, devuelve el nombre corto
End Function

