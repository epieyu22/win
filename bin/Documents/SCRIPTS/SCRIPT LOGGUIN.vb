Public Uid, Tuser, hashDB As String
Private Sub CommandButton1_Click()
On Error GoTo Errores
    Call InicializarAPP.iniciar
    Call logueo
    Exit Sub
Errores:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'CERRAR TODO SI PRECIONAN LA X ROJA
    If CloseMode <> 1 Then
        ThisWorkbook.Application.Visible = True
        Application.DisplayAlerts = False
        Call Cerrar_todo
    End If
End Sub
Public Sub Cerrar_todo()
    ThisWorkbook.Activate
    Sheets("ActMacro").Visible = True
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> "ActMacro" Then
            ws.Visible = False
        End If
    Next ws
ThisWorkbook.Close SaveChanges:=True
End Sub
Sub logueo()
    Select Case TextBox1.Text
    Case "admin"
        If TextBox1.Text = "admin" And TextBox2.Text = "admin" Then
            ThisWorkbook.Application.Visible = True
            ThisWorkbook.Activate
            Sheets("DASHBOARD").Visible = True
            Sheets("ActMacro").Visible = False
            Sheets("DASHBOARD").Activate
            Call varios.PerfilAdmin
            ActiveSheet.Range("B1").value = "Administrador"
            Unload Frm_Login
            Exit Sub
        Else
            MsgBox "Credenciales incorrectas", vbCritical
            Unload Frm_Login
            Frm_Login.Show
            Exit Sub
        End If
    Case Is <> "admin"
        Call SQLConnect.abrirConexion
        Set recordset1 = conn.Execute("SELECT TBL_USERS.id_user, TBL_USERS.pass_hash, TBL_USERS.tipo_user, TBL_USERS.Nom_comple FROM TBL_USERS WHERE TBL_USERS.cedula='" & TextBox1.Text & "' ;")
        If recordset1.EOF = True Then
            MsgBox "Credenciales incorrectas", vbCritical
            Call SQLConnect.CerrarConexion
        Else
            hashDB = RTrim(recordset1.Fields(1))
            'Call InicializarAPP.cifrar(CStr(TextBox2.Text))
            If hashDB = TextBox2.Text Then
                Uid = RTrim(recordset1.Fields(0))
                'ACTIVAR PERFILES DE USUARIOS
                Tuser = RTrim(recordset1.Fields(2))
                If Tuser = "admin" Then
                    Call varios.PerfilAdmin
                Else
                    Call varios.PerfilUser
                End If
                ThisWorkbook.Application.Visible = True
                ThisWorkbook.Activate
                Sheets("DASHBOARD").Visible = True
                Sheets("ActMacro").Visible = False
                Sheets("DASHBOARD").Select
                ActiveSheet.Range("B1").value = recordset1.Fields(3)
                Call varios.dashboardActua
                Unload Frm_Login
                Call SQLConnect.CerrarConexion
                Set recordset1 = Nothing
                Exit Sub
            Else
                MsgBox "Credenciales incorrectas", vbCritical
                Unload Frm_Login
                Frm_Login.Show
                Call SQLConnect.CerrarConexion
                recordset1.ClearFields
                Exit Sub
            End If
        End If
    End Select
End Sub