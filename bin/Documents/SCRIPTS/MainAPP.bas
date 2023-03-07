Attribute VB_Name = "MainAPP"
Dim miRuta As String
Public Registro As Object
Public AccesS_NomBase, AccesS_Ruta, pas, hash As String
Public vRegistro As String
Public Sub CreaRutaCerti(Ruta As String)
    'CREO REGISTRO
    Call RegistroCreaRuta("RutaCerti", Ruta)
    Call LeerTodoRegistro("RutaCerti")
End Sub
Public Sub RegistroCreaRuta(VNombre As String, VDato As String)
    miRuta = "HKEY_CURRENT_USER\SOFTWARE\ETPV_PMA\" & VNombre
    Set Registro = CreateObject("WScript.Shell")
    Registro.RegWrite miRuta, VDato
End Sub
Public Sub LeerTodoRegistro(VRutaRgd As String)
    Set Registro = CreateObject("WScript.Shell")
    vRegistro = Registro.RegRead("HKEY_CURRENT_USER\SOFTWARE\ETPV_PMA\" & VRutaRgd)
End Sub
'NO USO
Public Sub CreaRegistro(Ruta As String)
    Call RegistroCreaRuta("AccesS_Ruta_Certi", Ruta)
    Call RegistroCreaRuta("AccesS_NomBase", "BD_ETPV-CERTIFICADOS")
    Call RegistroCreaRuta("pass", "Masterkey15*")
End Sub
Public Sub iniciar()
On Error GoTo Errores
    Call LeerRegistro
    Exit Sub
Errores:
    Call CreaRutaCerti
    Call LeerRegistro
    MsgBox "Acceso inical Configurado", vbInformation
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
