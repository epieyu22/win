Attribute VB_Name = "SQLConnect"
Public conn As ADODB.Connection
Public recordset1, recordset2, recordset3 As ADODB.Recordset
Sub abrirConexion()
On Error GoTo Errores
    Set conn = New ADODB.Connection
    Call InicializarAPP.LeerRegistro
    conn.Provider = "microsoft.ACE.OLEDB.12.0"
    conn.Properties("jet OLEDB:Database Password") = pas
    conn.Open CStr(AccesS_Ruta + "\" + AccesS_NomBase + ".accdb")
    Debug.Print "Connexion Up"
    Exit Sub
Errores:
    MsgBox Err.Description, vbCritical
    Call CerrarConexion
    Exit Sub
End Sub
Sub CerrarConexion()
On Error GoTo Errores
    If conn Is Nothing Then Exit Sub
        conn.Close: Set conn = Nothing
    Debug.Print "Connexion Down"
    Exit Sub
Errores:
    MsgBox Err.Description, vbCritical
End Sub
Sub RScriptSQl(script As String)
        Call abrirConexion
        conn.Execute (script)
        Call CerrarConexion
End Sub
