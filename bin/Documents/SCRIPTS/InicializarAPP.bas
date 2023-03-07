Attribute VB_Name = "InicializarAPP"
Dim miRuta As String
Public Registro As Object
Public AccesS_NomBase, AccesS_Ruta, pas, hash As String
Public Sub iniciar()
On Error GoTo Errores
    Call LeerRegistro
    Exit Sub
Errores:
    Call CreaRegistro
    Call LeerRegistro
    MsgBox "Acceso inical Configurado", vbInformation
End Sub
Public Sub CreaRegistro()
    Call RegistroCreaRuta("AccesS_Ruta", "C:\APLICATIVO ETPV - CERTIFICACIONES\BASE DE DATOS")
    Call RegistroCreaRuta("AccesS_NomBase", "BD_ETPV-CERTIFICADOS")
    Call RegistroCreaRuta("pass", "Masterkey15*")
End Sub
Public Sub RegistroCreaRuta(VNombre As String, VDato As String)
    miRuta = "HKEY_CURRENT_USER\SOFTWARE\ETPV_PMA\" & VNombre
    Set Registro = CreateObject("WScript.Shell")
    Registro.RegWrite miRuta, VDato
End Sub
Public Sub LeerRegistro()
    Set Registro = CreateObject("WScript.Shell")
    AccesS_NomBase = Registro.RegRead("HKEY_CURRENT_USER\SOFTWARE\ETPV_PMA\AccesS_NomBase")
    AccesS_Ruta = Registro.RegRead("HKEY_CURRENT_USER\SOFTWARE\ETPV_PMA\AccesS_Ruta")
    pas = Registro.RegRead("HKEY_CURRENT_USER\SOFTWARE\ETPV_PMA\pass")
End Sub
Sub cifrar(hash As String)
    Dim Texto As String
    Dim cCifrado As clsCifrado
    Set cCifrado = New clsCifrado
    cCifrado.Clave = pas
    hash = cCifrado.cifrar(hash)
End Sub
Sub descifrar(hash As String)
    Dim Texto As String
    Dim cCifrado As clsCifrado
    Set cCifrado = New clsCifrado
    cCifrado.Clave = pas
    hash = cCifrado.descifrar(hash)
End Sub
