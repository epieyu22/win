Attribute VB_Name = "Importar"
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
    Dim fName, mC As String, i, j As Integer, fso As Object, fsoFile As Object, txt, txtDato As Variant
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
                    'LIMPIAR COMILLAS
                    txtDato(j) = Replace(txtDato(j), Chr(34), "")
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
                mNumBin = DecToBin(rsSOLICITUDES!Id, 20)
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
            mIdSolic = rsSOLICITUDES!Id
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
Sub limpieza()
Set DataBase = CurrentDb
Set rsSOLICITUDES = DataBase.OpenRecordset("tbl_solicitudes", dbOpenDynaset)

rsSOLICITUDES.MoveLast
rsSOLICITUDES.MoveFirst

Do While Not rsSOLICITUDES.EOF
    rsSOLICITUDES.Edit
    If Not IsNull(rsSOLICITUDES!PrimerNombre) Then rsSOLICITUDES!PrimerNombre = Replace(rsSOLICITUDES!PrimerNombre, Chr(34), "")
    If Not IsNull(rsSOLICITUDES!SegundoNombre) Then rsSOLICITUDES!SegundoNombre = Replace(rsSOLICITUDES!SegundoNombre, Chr(34), "")
    If Not IsNull(rsSOLICITUDES!PrimerApellido) Then rsSOLICITUDES!PrimerApellido = Replace(rsSOLICITUDES!PrimerApellido, Chr(34), "")
    If Not IsNull(rsSOLICITUDES!SegundoApellido) Then rsSOLICITUDES!SegundoApellido = Replace(rsSOLICITUDES!SegundoApellido, Chr(34), "")
    If Not IsNull(rsSOLICITUDES!codigoCertificado) Then rsSOLICITUDES!codigoCertificado = Replace(rsSOLICITUDES!codigoCertificado, Chr(34), "")
    
    If Not IsNull(rsSOLICITUDES!NombreCompleto) Then rsSOLICITUDES!NombreCompleto = Replace(rsSOLICITUDES!NombreCompleto, Chr(34), "")
    If Not IsNull(rsSOLICITUDES!NomPDF) Then rsSOLICITUDES!NomPDF = Replace(rsSOLICITUDES!NomPDF, Chr(34), "")
    If Not IsNull(rsSOLICITUDES!PeriodoAsistencia) Then rsSOLICITUDES!PeriodoAsistencia = Replace(rsSOLICITUDES!PeriodoAsistencia, Chr(34), "")
    If Not IsNull(rsSOLICITUDES!TiposAsistencia) Then rsSOLICITUDES!TiposAsistencia = Replace(rsSOLICITUDES!TiposAsistencia, Chr(34), "")
    rsSOLICITUDES.Update
    rsSOLICITUDES.MoveNext
Loop

DataBase.Close
Set rsSOLICITUDES = Nothing
Set DataBase = Nothing

End Sub
