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
'1.
'BOTON IMPORTAR DATOS
Public Sub btnImport_Click()
    Call ImportarController.ENCONTRAR_RUTA 'FUNCION QUE ENCUENTRA RUTA
        If Ruta <> "" Then
            
        End If
    Exit Sub
End Sub
'BOTON DE SALIR
Private Sub btnSalir_Click()
mv = MsgBox("Esta seguro que desea cancelar?", vbYesNo, "Cancelar proceso")
mensaje = ""
Select Case mv
    Case 6  'Yes
        mensaje = "cancel"
        DoCmd.Close acForm, "Frm_encabezados", acSaveYes
        MsgBox "Favor Importar nuevamente", vbCritical, ""
    Case 7 'No
        mensaje = ""
End Select
End Sub


'2.IMPORTAR
Option Explicit
'ENCUENTRA RUTA
Public Sub ENCONTRAR_RUTA()
   Dim fDialog, fso, fsoFile As Object
   Set fDialog = Application.FileDialog(3)
   With fDialog
    .AllowMultiSelect = True
    .Title = "Selecciones las solicitudes"
    .Filters.Clear
    .Filters.Add "All Files", "*.xlsx"
    If .Show = True Then
        Ruta = .SelectedItems(1)
        Call TEMP_6A
    Else
        Exit Sub
    End If
   End With
End Sub
Sub TEMP_6A()
Dim mv As Integer
'On Error GoTo Errores
    CurrentDb.CreateTableDef ("6A_TEMPORAL")
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "TEMP_6A", Ruta
    Call VALIDA_6A
    Exit Sub
'Errores:
'MsgBox "Error 1: No es posible importar, El documento no comple con requisitos minimos", vbCritical
'DoCmd.DeleteObject acTable, "TEMP_6A"
End Sub
Sub VALIDA_6A()
Dim contCampos As Integer
Dim mv As String
'On Error GoTo Errores
    'ESTABLECER BASE DE DATOS Y RECORSET
    Set BaseDeDatos = CurrentDb
    Set RSgeneral = BaseDeDatos.OpenRecordset("TEMP_6A", dbOpenDynaset)
    'ENCONTRAR EL ENCABEZADO
    Do While Not RSgeneral.EOF
        If IsNumeric(RSgeneral!F1) Then
            RSgeneral.MoveNext
        Else
            Exit Do
        End If
        DoEvents
    Loop
    'ENCONTRAR PARAMETROS: MAPEO PARA IMPORTAR
    Set RSParametros = BaseDeDatos.OpenRecordset("TBL_PARAMETROS", dbOpenDynaset)
    Call ParametrosController.GetValorParametros("TIPO_DOCUMENTO", 0) 'TIPO DE DOCUMENTO
    Call ParametrosController.GetValorParametros("NUMERO_DOCUMENTO", 1) 'NÚMERO DE DOCUMENTO
    Call ParametrosController.GetValorParametros("PRIMER_NOMBRE", 2) 'PRIMER NOMBRE
    Call ParametrosController.GetValorParametros("SEGUNDO_NOMBRE", 3) 'SEGUNDO NOMBRE
    Call ParametrosController.GetValorParametros("PRIMER_APELLIDO", 4) 'PRIMER APELLIDO
    Call ParametrosController.GetValorParametros("SEGUNDO_APELLIDO", 5) 'SEGUNDO APELLIDO
    Call ParametrosController.GetValorParametros("FECHA_NACIMIENTO", 6) 'FECHA DE NACIMIENTO
    Call ParametrosController.GetValorParametros("ETNIA", 7) 'ÉTNIA
    Call ParametrosController.GetValorParametros("SEXO", 8) 'SEXO
    Call ParametrosController.GetValorParametros("GRADO", 9) 'GRADO
    Call ParametrosController.GetValorParametros("GRUPO", 10) 'GRUPO
    Call ParametrosController.GetValorParametros("TIPO_JORNADA", 11) 'TIPO_JORNADA
    'INSTITUCIONES
'    Call ParametrosController.GetValorParametros("TIPO_DOCUMENTO", 12) 'NOMBRE_SEDE
'    Call ParametrosController.GetValorParametros("TIPO_DOCUMENTO", 13) 'NOMBRE_EE
'    Call ParametrosController.GetValorParametros("TIPO_DOCUMENTO", 14) 'CODIGO_DANE
'    Call ParametrosController.GetValorParametros("TIPO_DOCUMENTO", 15) 'DANE_ANTERIOR
'    Call ParametrosController.GetValorParametros("TIPO_DOCUMENTO", 16) 'CONS_SEDE
'    Call ParametrosController.GetValorParametros("TIPO_DOCUMENTO", 17) 'EXP_MUN
'    Call ParametrosController.GetValorParametros("TIPO_DOCUMENTO", 18) 'EXP_DEPTO
'    Call ParametrosController.GetValorParametros("TIPO_DOCUMENTO", 19)
'    Call ParametrosController.GetValorParametros("TIPO_DOCUMENTO", 20)
    
    'VALIDACION: PARAMETROS DE ASISTENCIAS
    contCampos = 0
    For cont = 0 To (RSgeneral.Fields.Count - 1)
        Select Case Trim(RSgeneral.Fields(cont))
            Case VecParametros(0, 1)
                VecParametros(0, 1) = RSgeneral.Fields(cont) 'TIPO DE DOCUMENTO
                contCampos = contCampos + 1
            Case VecParametros(1, 1)
                VecParametros(1, 1) = RSgeneral.Fields(cont) 'NÚMERO DE DOCUMENTO
                contCampos = contCampos + 1
            Case VecParametros(2, 1)
                VecParametros(2, 1) = RSgeneral.Fields(cont) 'PRIMER NOMBRE
                contCampos = contCampos + 1
            Case VecParametros(3, 1)
                VecParametros(3, 1) = RSgeneral.Fields(cont) 'SEGUNDO NOMBRE
                contCampos = contCampos + 1
            Case VecParametros(4, 1)
                VecParametros(4, 1) = RSgeneral.Fields(cont) 'PRIMER APELLIDO
                contCampos = contCampos + 1
            Case VecParametros(5, 1)
                VecParametros(5, 1) = RSgeneral.Fields(cont) 'SEGUNDO APELLIDO
                contCampos = contCampos + 1
            Case VecParametros(6, 1)
                VecParametros(6, 1) = RSgeneral.Fields(cont) 'FECHA DE NACIMIENTO
                contCampos = contCampos + 1
            Case VecParametros(7, 1)
                VecParametros(7, 1) = RSgeneral.Fields(cont) 'ÉTNIA
                contCampos = contCampos + 1
            Case VecParametros(8, 1)
                VecParametros(8, 1) = RSgeneral.Fields(cont) 'SEXO
                contCampos = contCampos + 1
            Case VecParametros(9, 1)
                VecParametros(9, 1) = RSgeneral.Fields(cont) 'GRADO
                contCampos = contCampos + 1
            Case VecParametros(10, 1)
                VecParametros(10, 1) = RSgeneral.Fields(cont) 'GRUPO
                contCampos = contCampos + 1
            Case VecParametros(11, 1)
                VecParametros(11, 1) = RSgeneral.Fields(cont) 'TIPO_JORNADA
                contCampos = contCampos + 1
        End Select
        cont = cont + 1
        DoEvents
    Next
    
    If contCampos < 12 Then
        mv = MsgBox("Desea establecerla manualmente?", vbYesNo, "Estructura no coincide")
        Select Case mv
            Case 6  'Yes
'                Forms![Frm_asistencias].Visible = False
                DoCmd.OpenForm "Frm_encabezados", acNormal, , , acFormReadOnly, acWindowNormal
                Exit Sub
            Case 7 'No
                Set RSgeneral = Nothing
                Set RSParametros = Nothing
                MsgBox "Favor intentelo nuevamente", vbCritical, ""
                DoCmd.DeleteObject acTable, "TEMP_6A"
                Exit Sub
        End Select
    Else
        Set RSgeneral = Nothing
        Set RSParametros = Nothing
        
        Forms![Form_Frm_encabezado].TxtNomArch.Caption = Ruta
        Exit Sub
    End If

'Errores:
'mv = MsgBox("No se pudo establecer el encabezado dentro del archivo. Si desea establecer el encabezado manualmente presione Si. de lo contrario presione No, para salir", vbYesNo)
'Select Case mv
'    Case 6  'Yes
'
'    Case 7 'No
'End Select
End Sub

'PARAMETROS CONTROLER
Sub GetValorParametros(ByVal val As String, ByVal positions As String)
    RSParametros.Filter = "Modulo = 'ASISTENCIAS' AND Grupo = 'ANEXO6A_13A' AND NomParametro = '" & val & "'"
    Set RSParametrosFilt = RSParametros.OpenRecordset
    If Not RSParametros.EOF And Not IsNull(RSParametrosFilt!ValorParametro) Then
        VecParametros(positions, 0) = RSParametrosFilt!ValorParametro
    Else
        VecParametros(positions, 0) = ""
    End If
End Sub
Option Explicit
Public Ruta, mDiaNacimiento, Idfrmlo, itemDec As String
Public mRegAnterior, mIdSolic As Integer
Public DataBase As DAO.DataBase
Public rsSOLICITUDES, rsSOLICITUDESblanco, rsESTADOSOLICITUDES, rsESTADOSOLICITUDESfilt As DAO.Recordset
Public Sub IMPORTARDOC(ByVal rutaFamilias As String, ByVal rutaFormularios As String)
    'IMPORTA EL DOCUMENTO DE FAMILIA
    Call ADECUAR_CSV(rutaFamilias, "flma")
    'IMPORTA EL DOCUMENTO DE FORMULARIO
    Call ADECUAR_CSV(rutaFormularios, "frmlo")
    'GENERAR CAMPOS NECESARIO EN LA TABLA
    Call DATOSGENERADOS
    MsgBox "Cargue Exitoso", vbInformation
End Sub

Public Sub ADECUAR_CSV(ByVal rutaDoc As String, ByVal TipoDoc As String)
    'CREA OBJETO, ABRE Y ALMACENA EN VARIABLE EL DOCUMENTO CSV
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fsoFile = fso.OpenTextFile(rutaDoc, 1)
    txt = fsoFile.ReadAll
    fsoFile.Close
    'LE INDICO QUE BASE DE DATOS USO
    Set DataBase = CurrentDb
    'DIVIDE EN LINEAS EL CONTENIDO DEL CSV
    txt = Split(txt, vbNewLine)
    i = 1
    'RECORRE TODAS LAS LINEAS DEL CSV
    Do While txt(i) <> ""
        'VALIDA QUE EL DOCUMENTO ESTE DIVIDI POR COMAS
        If InStr(txt(i), ";") = 0 Then
            mC = ","
        Else
            MsgBox "Error: Este documento debe estar separado por coma", vbCritical
            Exit Sub
        End If
        'DIVIDO POR COMA LA LINEA Y ALMACENO EN UN VECTOR
        txtDato = Split(txt(i), mC)
        'ALMACENAR BENEFICIARIO
        Select Case TipoDoc
        Case "flma"
            'VALIDA DUPLICIDAD CARGUES ANTERIORES, Si esta duplicado no lo almacena
            Set rsSOLICITUDES = DataBase.OpenRecordset("tbl_solicitudes", dbOpenDynaset)
            rsSOLICITUDES.Filter = "IdFormulario = '" & CStr(txtDato(17)) & "' AND CodBinario <> ''"
            Set rsESTADOSOLICITUDESfilt = rsSOLICITUDES.OpenRecordset
            If rsESTADOSOLICITUDESfilt.EOF Then
                'GUARDO DATO
                j = 0
                For j = 0 To 19
                    'LIMPIAR ESPACIOS EN BLANCO
                    txtDato(j) = Replace(txtDato(j), " ", "")
                    'CONVERTIR CARACTERES UTF 8
                    txtDato(j) = SpecialCharReplace(txtDato(j))
    '                txtDato(j) = StrConv(txtDato(j), vbUnicode)
                    'ALMACENAR
                    Call guardarFamilia(i, j, txtDato(j), TipoDoc)
                Next j
                    'ESTADO DE SOLICITUD
                    Call CambiarEstadoSolicitud(1, mIdSolic)
            Else
                MsgBox "La solicitud de: " & txtDato(2) & " " & txtDato(4) & " Ya fue cargada. Por tal motivo no sera almacenada", vbInformation
            End If
        Case "frmlo"
            'IDENTIFICA Y RECORRE INVERSAMENTE EL RECORSET
            Set rsSOLICITUDES = DataBase.OpenRecordset("tbl_solicitudes", dbOpenDynaset)
            If Not rsSOLICITUDES.EOF Then
                rsSOLICITUDES.MoveLast
                Do While Not IsEmpty(rsSOLICITUDES!CodBinario) And Not rsSOLICITUDES.BOF
                'MIENTRAS ENCUENTRE VALORES VACIOS ADICIONE
                    If rsSOLICITUDES!IdFormulario = txtDato(18) Then
                        rsSOLICITUDES.Edit
                        rsSOLICITUDES!fechaSolicitud = txtDato(0) & " " & txtDato(1)
                        rsSOLICITUDES.Update
                    End If
                rsSOLICITUDES.MovePrevious
                DoEvents
                Loop
            End If
        rsSOLICITUDES.Close
        Set rsSOLICITUDES = Nothing
        End Select
    i = i + 1
    Loop
DataBase.Close
Set DataBase = Nothing
fsoFile.Close
Set fsoFile = Nothing
End Sub
Public Sub DATOSGENERADOS()
    Dim mNumBin, mNomComp, strInput As String
    Dim txtDato As Variant
    'ESTABLECER BASE DE DATOS Y RECORSET
    Set DataBase = CurrentDb
    Set rsSOLICITUDES = DataBase.OpenRecordset("tbl_solicitudes", dbOpenDynaset)
    If Not rsSOLICITUDES.EOF Then
        'MUEVO AL FIN E INICIO POR SI A CASO
        rsSOLICITUDES.MoveLast
        rsSOLICITUDES.MoveFirst
        'RECORRO TODA LA TABLA E INSERTO
        Do While Not rsSOLICITUDES.EOF
            If IsNull(rsSOLICITUDES!CodBinario) Then
                rsSOLICITUDES.Edit
                'CIUDAD DONDE SE GENERA
                rsSOLICITUDES!Ciudad_genera = "Bogotá D.C"
                'FECHA DE SOLICITUD EN FORMATO
                If Not IsNull(rsSOLICITUDES!fechaSolicitud) Then
                    txtDato = Replace(rsSOLICITUDES!fechaSolicitud, """", "")
                    txtDato = Split(txtDato, " ")
                    'MESES
                    txtDato(0) = Replace(txtDato(0), "ene", "1")
                    txtDato(0) = Replace(txtDato(0), "feb", "2")
                    txtDato(0) = Replace(txtDato(0), "mar", "3")
                    txtDato(0) = Replace(txtDato(0), "abr", "4")
                    txtDato(0) = Replace(txtDato(0), "may", "5")
                    txtDato(0) = Replace(txtDato(0), "jun", "6")
                    txtDato(0) = Replace(txtDato(0), "jul", "7")
                    txtDato(0) = Replace(txtDato(0), "ago", "8")
                    txtDato(0) = Replace(txtDato(0), "sep", "9")
                    txtDato(0) = Replace(txtDato(0), "oct", "10")
                    txtDato(0) = Replace(txtDato(0), "nov", "11")
                    txtDato(0) = Replace(txtDato(0), "dic", "12")
                    rsSOLICITUDES!fechaSolicitudFormato = txtDato(1) & "/" & txtDato(0) & "/" & txtDato(3)
                End If
                'NUMERO BINARIO
                mNumBin = DecToBin(rsSOLICITUDES!id, 20)
                rsSOLICITUDES!CodBinario = mNumBin
                'NOMBRE DEL PDF
                rsSOLICITUDES!NomPDF = "COL-ETPV-" & rsSOLICITUDES!TipoDoc & "-" & rsSOLICITUDES!NroDocumento & "_" & rsSOLICITUDES!PrimerNombre & "_" & rsSOLICITUDES!SegundoNombre & "_" & rsSOLICITUDES!PrimerApellido & "_" & rsSOLICITUDES!SegundoApellido & "_" & rsSOLICITUDES!Edad
                'NOMBRE COMPLETO
                mNomComp = rsSOLICITUDES!PrimerNombre & " " & rsSOLICITUDES!SegundoNombre & " " & rsSOLICITUDES!PrimerApellido & " " & rsSOLICITUDES!SegundoApellido
                rsSOLICITUDES!NombreCompleto = mNomComp
                'CODIGO CERTIFICADO
                rsSOLICITUDES!codigoCertificado = "COL-ETPV-" & rsSOLICITUDES!TipoDoc & "-" & rsSOLICITUDES!NroDocumento & "/" & mNumBin & "/" & mNomComp
                rsSOLICITUDES.Update
            End If
        rsSOLICITUDES.MoveNext
        DoEvents
        Loop
    End If
rsSOLICITUDES.Close
Set rsSOLICITUDES = Nothing
DataBase.Close
Set DataBase = Nothing
End Sub
Public Sub guardarFamilia(ByVal i As Integer, ByVal j As Integer, ByVal dato As String, ByVal TipoDoc As String)
    Set rsSOLICITUDES = DataBase.OpenRecordset("tbl_solicitudes", dbOpenDynaset)
If i <> mRegAnterior Then
    rsSOLICITUDES.AddNew
Else
    rsSOLICITUDES.MoveLast
    rsSOLICITUDES.Edit
End If

If TipoDoc = "flma" Then
    mRegAnterior = i
    Select Case j
        Case "0"
            rsSOLICITUDES!KEY = dato
        Case "2"
            rsSOLICITUDES!PrimerNombre = dato
        Case "3"
            rsSOLICITUDES!SegundoNombre = dato
        Case "4"
            rsSOLICITUDES!PrimerApellido = dato
        Case "5"
            rsSOLICITUDES!SegundoApellido = dato
        Case "6"
            rsSOLICITUDES!TipoDoc = dato
        Case "7"
            rsSOLICITUDES!NroDocumento = dato
        Case "8"
            rsSOLICITUDES!CabezaHogar = dato
        Case "9"
            mDiaNacimiento = dato & " "
        Case "10"
            rsSOLICITUDES!FechaNacimiento = mDiaNacimiento & dato
        Case "11"
            rsSOLICITUDES!Edad = dato
        Case "12"
            rsSOLICITUDES!TelefonoMovil = dato
        Case "13"
            rsSOLICITUDES!TelefonoWS = dato
        Case "14"
            rsSOLICITUDES!CorreoElectronico = dato
        Case "15"
            rsSOLICITUDES!TiposAsistencia = dato
        Case "16"
            rsSOLICITUDES!PeriodoAsistencia = dato
        Case "17"
            rsSOLICITUDES!IdFormulario = dato
        Case "18"
            mIdSolic = rsSOLICITUDES!id
    End Select
    rsSOLICITUDES.Update
Else
End If
rsSOLICITUDES.Close
Set rsSOLICITUDES = Nothing
End Sub
Public Sub ENCONTRAR_RUTA()
   Dim fDialog, fso, fsoFile As Object
   Dim txt, txtDato As Variant
   Dim i As Integer
   Set fDialog = Application.FileDialog(3)
   
   With fDialog
    .AllowMultiSelect = True
    .Title = "Selecciones las solicitudes"
    .Filters.Clear
    .Filters.Add "All Files", "*.csv"
    If .Show = True Then
        Ruta = .SelectedItems(1)
    Else
        Exit Sub
    End If
   End With
    'CREA OBJETO, ABRE Y ALMACENA EN VARIABLE EL DOCUMENTO CSV
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fsoFile = fso.OpenTextFile(Ruta, 1)
    txt = fsoFile.ReadAll
    fsoFile.Close
    'DIVIDE EN LINEAS EL CONTENIDO DEL CSV
    txt = Split(txt, vbNewLine)
    i = 1
    'VALIDA QUE EL DOCUMENTO ESTE DIVIDI POR COMAS
    If InStr(txt(i), ";") <> 0 Then
        MsgBox "Error: Este documento debe estar separado por coma", vbCritical
        Ruta = "Vuelva a buscar el documento"
        Exit Sub
    End If
End Sub
Function SpecialCharReplace(ByVal strInput As String) As String
    strInput = Replace(strInput, "Ã€", "À")
    strInput = Replace(strInput, "Ã‚", "Â")
    strInput = Replace(strInput, "Ãƒ", "Ã")
    strInput = Replace(strInput, "Ã„", "Ä")
    strInput = Replace(strInput, "Ã…", "Å")
    strInput = Replace(strInput, "Ã†", "Æ")
    strInput = Replace(strInput, "Ã‡", "Ç")
    strInput = Replace(strInput, "Ãˆ", "È")
    strInput = Replace(strInput, "Ã‰", "É")
    strInput = Replace(strInput, "ÃŠ", "Ê")
    strInput = Replace(strInput, "Ã‹", "Ë")
    strInput = Replace(strInput, "ÃŒ", "Ì")
    strInput = Replace(strInput, "ÃŽ", "Î")
    strInput = Replace(strInput, "Ã‘", "Ñ")
    strInput = Replace(strInput, "Ã’", "Ò")
    strInput = Replace(strInput, "Ã“", "Ó")
    strInput = Replace(strInput, "Ã”", "Ô")
    strInput = Replace(strInput, "Ã•", "Õ")
    strInput = Replace(strInput, "Ã–", "Ö")
    strInput = Replace(strInput, "Ã—", "×")
    strInput = Replace(strInput, "Ã™", "Ù")
    strInput = Replace(strInput, "Ãš", "Ú")
    strInput = Replace(strInput, "Ãœ", "Ü")
    strInput = Replace(strInput, "Ãž", "Þ")
    strInput = Replace(strInput, "ÃŸ", "ß")
    strInput = Replace(strInput, "Ã¡", "á")
    strInput = Replace(strInput, "Ã¢", "â")
    strInput = Replace(strInput, "Ã£", "ã")
    strInput = Replace(strInput, "Ã¤", "ä")
    strInput = Replace(strInput, "Ã¥", "å")
    strInput = Replace(strInput, "Ã¦", "æ")
    strInput = Replace(strInput, "Ã§", "ç")
    strInput = Replace(strInput, "Ã¨", "è")
    strInput = Replace(strInput, "Ã©", "é")
    strInput = Replace(strInput, "Ãª", "ê")
    strInput = Replace(strInput, "Ã«", "ë")
    strInput = Replace(strInput, "Ã¬", "ì")
    strInput = Replace(strInput, "Ã­", "í")
    strInput = Replace(strInput, "Ã®", "î")
    strInput = Replace(strInput, "Ã¯", "ï")
    strInput = Replace(strInput, "Ã°", "ð")
    strInput = Replace(strInput, "Ã±", "ñ")
    strInput = Replace(strInput, "Ã²", "ò")
    strInput = Replace(strInput, "Ã³", "ó")
    strInput = Replace(strInput, "Ã´", "ô")
    strInput = Replace(strInput, "Ãµ", "õ")
    strInput = Replace(strInput, "Ã¶", "ö")
    strInput = Replace(strInput, "Ã·", "÷")
    strInput = Replace(strInput, "Ã¸", "ø")
    strInput = Replace(strInput, "Ã¹", "ù")
    strInput = Replace(strInput, "Ãº", "ú")
    strInput = Replace(strInput, "Ã»", "û")
    strInput = Replace(strInput, "Ã¼", "ü")
    strInput = Replace(strInput, "Ã›", "Û")
    SpecialCharReplace = strInput
End Function
Function DecToBin(ByVal DecimalIn As String, Optional NumberOfBits As Variant) As String
  DecToBin = ""
  DecimalIn = CDec(DecimalIn)
  Do While DecimalIn <> 0
    DecToBin = Trim$(Str$(DecimalIn - 2 * Int(DecimalIn / 2))) & DecToBin
    DecimalIn = Int(DecimalIn / 2)
  Loop
  If Not IsMissing(NumberOfBits) Then
    If Len(DecToBin) > NumberOfBits Then
      DecToBin = "Error - Number too large for bit size"
    Else
      DecToBin = Right$(String$(NumberOfBits, "0") & _
      DecToBin, NumberOfBits)
    End If
  End If
  itemDec = CStr(DecToBin)
End Function

Public Sub CambiarEstadoSolicitud(ByVal etapa As Integer, ByVal mIdSolic As Integer)
Set rsESTADOSOLICITUDES = DataBase.OpenRecordset("tbl_HistorialEstatusSolicitud", dbOpenDynaset)

rsESTADOSOLICITUDES.AddNew
rsESTADOSOLICITUDES!Solicitud = mIdSolic
rsESTADOSOLICITUDES!EtapaSolicitud = etapa
rsESTADOSOLICITUDES.Update

rsESTADOSOLICITUDES.Close
Set rsESTADOSOLICITUDES = Nothing
End Sub
'Public Sub ValidacionDuplicidadAnt(ByVal dato As String)
''DUPLICIDAD FRENTE A CARGUES ANTERIORES
'
'                'VALIDACION DE DUPLICIDAD SOLICITUDES PASADAS
''                    If j = 7 Then
''                        R
''                        DataBase.Execute ("SELECT * FROM tbl_solicitudes WHERE NroDocumento ='" & txtDato(j) & "' ")
''                    End If
'
'
'End Sub
'Public Sub ValidacionDuplicidadAnt(ByVal dato As String)
''DUPLICICDAD FRENTE AL MISMO CARGUE
'End Sub
