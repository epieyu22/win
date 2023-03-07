VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCifrado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Binary

Private clsClave As String

Property Get Clave() As String
    Clave = clsClave
End Property

Property Let Clave(value As String)
    clsClave = value
End Property


Function cifrar(Frase As String) As String

    Dim Cachos() As Byte
    Dim LaClave() As Byte
    Dim i As Integer
    Dim Largo As Integer

    If Frase <> "" Then
        Cachos() = StrConv(Frase, vbFromUnicode)
        LaClave() = StrConv(clsClave, vbFromUnicode)
        Largo = Len(clsClave)

        For i = LBound(Cachos) To UBound(Cachos)
            Cachos(i) = (Cachos(i) Xor LaClave(i Mod Largo)) + 34
        Next i

        cifrar = StrConv(Cachos(), vbUnicode)
    Else
        cifrar = ""
    End If

End Function

Function descifrar(Frase As String) As String

    Dim Cachos() As Byte
    Dim LaClave() As Byte
    Dim i As Integer
    Dim Largo As Integer

    If Frase <> "" Then
        Cachos() = StrConv(Frase, vbFromUnicode)
        LaClave() = StrConv(clsClave, vbFromUnicode)
        Largo = Len(clsClave)

        For i = LBound(Cachos) To UBound(Cachos)
            Cachos(i) = Cachos(i) - 34
            Cachos(i) = (Cachos(i) Xor LaClave(i Mod Largo))
        Next i

        descifrar = StrConv(Cachos(), vbUnicode)
    Else
        descifrar = ""
    End If

End Function
