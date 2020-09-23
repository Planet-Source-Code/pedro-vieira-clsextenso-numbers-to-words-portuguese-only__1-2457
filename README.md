<div align="center">

## clsExtenso \(numbers to words \- Portuguese only\)


</div>

### Description

Convert numbers to words (Portuguese only!)
 
### More Info
 
a number (double)

The number coverted to Portuguese words.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Pedro Vieira](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/pedro-vieira.md)
**Level**          |Unknown
**User Rating**    |5.0 (5 globes from 1 user)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/pedro-vieira-clsextenso-numbers-to-words-portuguese-only__1-2457/archive/master.zip)

### API Declarations

```
Private Declare Function GetLocaleInfo& _
Lib "kernel32" Alias "GetLocaleInfoA" ( _
 ByVal Locale As Long, _
 ByVal LCType As Long, _
 ByVal lpLCData As String, _
 ByVal cchData As Long)
Private Const LOCALE_USER_DEFAULT& = &H400
Private Const LOCALE_SDECIMAL& = &HE
Private Const LOCALE_SCURRENCY& = &H14
Private Const LOCALE_SMONDECIMALSEP& = &H16
Public Enum enmFormat
 Maiusculas
 Minusculas
 PrimeiraMaiuscula
End Enum
Private arrGrupo() As String
'2 dimensoes
'1º -> [0]=valor numérico do grupo; [1]=extenso
'2ª -> contador
Private Const E = "e "
Private Const Virgula = ", "
Private Const ZERO = "Zero "
Private Const UM = "Um "
Private Const DOIS = "Dois "
Private Const TRES = "Três "
Private Const QUATRO = "Quatro "
Private Const CINCO = "Cinco "
Private Const SEIS = "Seis "
Private Const SETE = "Sete "
Private Const OITO = "Oito "
Private Const NOVE = "Nove "
Private Const DEZ = "Dez "
Private Const ONZE = "Onze "
Private Const DOZE = "Doze "
Private Const TREZE = "Treze "
Private Const CATORZE = "Catorze "
Private Const QUINZE = "Quinze "
Private Const DEZASSEIS = "Dezasseis "
Private Const DEZASSETE = "Dezassete "
Private Const DEZOITO = "Dezoito "
Private Const DEZANOVE = "Dezanove "
Private Const VINTE = "Vinte "
Private Const TRINTA = "Trinta "
Private Const QUARENTA = "Quarenta "
Private Const CINQUENTA = "Cinquenta "
Private Const SESSENTA = "Sessenta "
Private Const SETENTA = "Setenta "
Private Const OITENTA = "Oitenta "
Private Const NOVENTA = "Noventa "
Private Const CEM = "Cem "
Private Const CENTO = "Cento "
Private Const DUZENTOS = "Duzentos "
Private Const TREZENTOS = "Trezentos "
Private Const QUATROCENTOS = "Quatrocentos "
Private Const QUINHENTOS = "Quinhentos "
Private Const SEISCENTOS = "Seiscentos "
Private Const SETECENTOS = "Setecentos "
Private Const OITOCENTOS = "Oitocentos "
Private Const NOVECENTOS = "Novecentos "
Private Const MIL = "Mil "
Private Const MILHAO = "Milhao "
Private Const MILHOES = "Milhoes "
Private Const BILIAO = "Biliao "
Private Const BILIOES = "Bilioes "
Private strUnidades(9) As String
Private strTeens(99) As String
Private strDezenas(9) As String
Private strCentenas(9) As String
Private strMilhares(9) As String
Private mstrDecSep As String * 1
Private mstrDefaultErrorMsgOverflow As String
Private Const ERR_OVERF = "Overflow"
'singular
Private mstrDefaultSufixoInteiro1 As String
Private Const SUF_INT1 = "Escudo "
Private mstrDefaultSufixoDecimal1 As String
Private Const SUF_DEC1 = "Centavo "
'plural
Private mstrDefaultSufixoInteiro2 As String
Private Const SUF_INT2 = "Escudos "
Private mstrDefaultSufixoDecimal2 As String
Private Const SUF_DEC2 = "Centavos "
Private Const MAX_NUMBER As Double = 999999999999.99
```


### Source Code

```
Private Sub msEncher()
'strUnidades(0) = ZERO ' deve ser um empty string!
strUnidades(1) = UM
strUnidades(2) = DOIS
strUnidades(3) = TRES
strUnidades(4) = QUATRO
strUnidades(5) = CINCO
strUnidades(6) = SEIS
strUnidades(7) = SETE
strUnidades(8) = OITO
strUnidades(9) = NOVE
'strTeens(0) = ZERO ' deve ser um empty string!
strTeens(1) = UM
strTeens(2) = DOIS
strTeens(3) = TRES
strTeens(4) = QUATRO
strTeens(5) = CINCO
strTeens(6) = SEIS
strTeens(7) = SETE
strTeens(8) = OITO
strTeens(9) = NOVE
strTeens(10) = DEZ
strTeens(11) = ONZE
strTeens(12) = DOZE
strTeens(13) = TREZE
strTeens(14) = CATORZE
strTeens(15) = QUINZE
strTeens(16) = DEZASSEIS
strTeens(17) = DEZASSETE
strTeens(18) = DEZOITO
strTeens(19) = DEZANOVE
strDezenas(0) = ""
strDezenas(1) = "-"
strDezenas(2) = VINTE
strDezenas(3) = TRINTA
strDezenas(4) = QUARENTA
strDezenas(5) = CINQUENTA
strDezenas(6) = SESSENTA
strDezenas(7) = SETENTA
strDezenas(8) = OITENTA
strDezenas(9) = NOVENTA
strCentenas(0) = ""
strCentenas(1) = CEM
strCentenas(2) = DUZENTOS
strCentenas(3) = TREZENTOS
strCentenas(4) = QUATROCENTOS
strCentenas(5) = QUINHENTOS
strCentenas(6) = SEISCENTOS
strCentenas(7) = SETECENTOS
strCentenas(8) = OITOCENTOS
strCentenas(9) = NOVECENTOS
End Sub
Private Function mfTraduzir(xGrupo%, xstr$) As String
'traduz um grupo de 3 algarismos
'(right pad)
On Error GoTo erro
Dim blnAnteriorRedondo As Boolean  'quando grupo anterior = '*00'
Dim ret$, xlen%
xlen = Len(xstr$)
Dim Unid As Byte, strUnid$
Dim Teen As Byte, strTeen$
Dim Dezena As Byte, strDezn$
Dim Centena As Byte, strCent$
 Unid = CByte(Right(xstr$, 1))
 Teen = CByte(Right(xstr$, 2))
 Dezena = CByte(Mid(xstr$, xlen - 1, 1))
 Centena = CByte(Mid(xstr$, xlen - 2, 1))
If Centena Then
strCent = IIf(Teen = 0, strCentenas(Centena), _
 IIf(Centena = 1, CENTO, strCentenas(Centena)) & _
 IIf(Teen = 0, "", E)) & " "
End If
strDezn = IIf(Teen > 19, strDezenas(Dezena), strTeens(Teen)) & _
 IIf(Unid And Teen > 19, E, "")
strUnid = IIf(Teen > 19, strUnidades(Unid), "")
ret = strCent & strDezn & strUnid
 Dim strNumAnterior$, strExtAnterior$
 On Error Resume Next
 strNumAnterior = arrGrupo(0, xGrupo - 1) 'grupo anterior
 strExtAnterior = arrGrupo(1, xGrupo - 1)
 blnAnteriorRedondo = Val(Right(strNumAnterior, 2)) = 0
 On Error GoTo erro
 Select Case xGrupo
  Case 0        '  000
  Case 1 'mil      '  000xxx
   arrGrupo(1, xGrupo - 1) = _
   IIf(blnAnteriorRedondo, _
   IIf(Val(strNumAnterior) = 0, "", E) & strExtAnterior, _
   E & strExtAnterior)
  ret = IIf(Val(xstr) = 0, "", _
   IIf(Val(xstr) = 1, MIL, ret & MIL))
  Case 2 'milhão     ' 000xxxxxx
   arrGrupo(1, xGrupo - 1) = _
   IIf(Val(strNumAnterior) = 0 And Val(arrGrupo(0, xGrupo - 2)) = 0, _
    "", IIf(Val(strNumAnterior) > 0, IIf(Val(arrGrupo(0, xGrupo - 2)) = 0, _
    E, Virgula), "") & strExtAnterior)
  ret = IIf(Val(xstr) = 0, "", _
   IIf(Val(xstr) = 1, ret & MILHAO, ret & MILHOES))
  Case 3 'bilião     ' 000xxxxxxxxx
   arrGrupo(1, xGrupo - 1) = _
   IIf(Val(strNumAnterior) = 0 And Val(arrGrupo(0, xGrupo - 2)) = 0 _
   And Val(arrGrupo(0, xGrupo - 3)) = 0, _
    "", IIf(Val(strNumAnterior) = 0, "", _
    IIf(Val(arrGrupo(0, xGrupo - 2)) = 0, E, Virgula)) & strExtAnterior)
  ret = IIf(Val(xstr) = 0, "", _
   IIf(Val(xstr) = 1, ret & BILIAO, ret & BILIOES))
 End Select
mfTraduzir = Trim(ret) & " "
Exit Function
erro:
 If Err = 5 Then
 Resume Next
 Else
 MsgBox Err & vbCrLf & Err.Description
 Resume Next
 End If
End Function
Private Sub Class_Initialize()
msEncher
mstrDecSep = mfstrGetDecimalSep
mstrDefaultErrorMsgOverflow = ERR_OVERF
mstrDefaultSufixoInteiro1 = SUF_INT1
mstrDefaultSufixoDecimal1 = SUF_DEC1
mstrDefaultSufixoInteiro2 = SUF_INT2
mstrDefaultSufixoDecimal2 = SUF_DEC2
End Sub
Public Function gfGet( _
 ByVal dblX As Double, _
 Optional ByVal lngFormat As Long = PrimeiraMaiuscula) As String
On Error GoTo erro
If dblX > MAX_NUMBER Then
 gfGet = mstrDefaultErrorMsgOverflow
 Exit Function
End If
dblX = Format(dblX, ".00")
Dim strInteiro$, strDecimal$
 msGetParts CStr(dblX), strInteiro, strDecimal
 Dim ret$, retInt$, retDec$
  If strInteiro <> "" Then
   If CDbl(strInteiro) > 0 Then
    retInt = mfstrProcessar(strInteiro)
   Else
    retInt = ZERO
   End If
   retInt = retInt & IIf(CDbl(strInteiro) = 1, mstrDefaultSufixoInteiro1, mstrDefaultSufixoInteiro2)
  End If
  If strDecimal <> "" Then
   If CDbl(strInteiro) = 0 Then
    retInt = ""
   Else
    retInt = retInt & E
   End If
   retDec = mfstrProcessar(strDecimal)
   retDec = retDec & IIf(CDbl(strDecimal) = 1, mstrDefaultSufixoDecimal1, mstrDefaultSufixoDecimal2)
  End If
  ret = retInt & retDec
 gfGet = IIf(lngFormat = Minusculas, LCase(ret), _
       IIf(lngFormat = Maiusculas, UCase(ret), _
       ret))
Exit Function
erro:
 gfGet = Err.Number & "; " & Err.Description
End Function
Public Property Get VersionInfo() As String
Dim ret$
ret = "Números Por Extenso" & vbCrLf & _
"Versão " & App.Major & "." & _
Format(App.Minor, "00") & "." & _
Format(App.Revision, "00") & vbCrLf & vbCrLf & _
"Pedro Vieira, [Bilógica, Lda]" & vbCrLf & vbCrLf & _
"bfe03116@mail.telepac.pt" & vbCrLf & _
"bilogica@mail.telepac.pt" & vbCrLf & vbCrLf & _
"Novembro de 1998"
VersionInfo = ret
End Property
Private Sub msGetParts(ByVal strAll$, ByRef strInt$, ByRef strDec$)
 Dim intVirgLoc%
 intVirgLoc = InStr(1, strAll, mstrDecSep)
  If intVirgLoc > 0 Then
   strInt = Mid(strAll, 1, intVirgLoc% - 1)
   strDec = Mid(strAll, intVirgLoc% + 1)
    If Len(strDec) = 1 Then strDec = strDec & "0"
  Else
   strInt = strAll$
   strDec = ""
  End If
End Sub
Private Function mfstrProcessar(strPart$) As String
Dim lp%, xlen%, cnt%, ret$, buf$
Dim xstart%
xlen = Len(strPart$)
 For lp = 1 To xlen Step 3
 'enviar o número em grupos de 3 algarismos
 xstart = xlen - (3 * cnt)
 xstart = IIf(xstart <= 0, 1, xstart)
 buf = Right(Left(strPart$, xstart), 3)
 ReDim Preserve arrGrupo(1, cnt)
 arrGrupo(0, cnt) = CDbl(buf)
 arrGrupo(1, cnt) = mfTraduzir(cnt, Format(buf, "000"))
  cnt = cnt + 1
 Next
 'obter a frase juntando os grupos traduzidos
 Dim xtemp As String
 For lp = UBound(arrGrupo, 2) To 0 Step -1
  xtemp = xtemp & arrGrupo(1, lp)
 Next
 'retirar espaços redundantes
 Dim red1$, inred1%, red2$, inred2%
 Dim tempA$, tempB$
 inred1 = 999: inred2 = 999
 red1 = " ": red2 = " ,"
 Do Until inred1 + inred2 = 0
  inred1 = InStr(1, xtemp, red1)
  inred2 = InStr(1, xtemp, red2)
  If inred1 > 0 Then
   xtemp = Trim(Left(xtemp, inred1) & Right(xtemp, Len(xtemp) - (inred1 + 1)))
  End If
  If inred2 > 0 Then Mid(xtemp, inred2, 2) = ", "
 Loop
 ret = xtemp & IIf(Right(xtemp, 1) <> " ", " ", "")
 mfstrProcessar = ret
End Function
Private Function mfstrGetDecimalSep() As String
Dim ret&
Dim buf As String * 10
ret = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, buf, Len(buf))
mfstrGetDecimalSep = Left(buf, InStr(1, buf, vbNullChar) - 1)
End Function
'   //////////////   PROPS   /////////////////////
Public Property Get DecimalSep() As String
 DecimalSep = mstrDecSep
End Property
Public Property Let DecimalSep(x As String)
 mstrDecSep = x
End Property
Public Property Get OverflowMsg() As String
 OverflowMsg = mstrDefaultErrorMsgOverflow
End Property
Public Property Let OverflowMsg(x As String)
 mstrDefaultErrorMsgOverflow = x
End Property
Public Property Get MaxNumber() As Double
 MaxNumber = MAX_NUMBER
End Property
Public Property Get SufixoInteiroSingular() As String
 SufixoInteiroSingular = mstrDefaultSufixoInteiro1
End Property
Public Property Let SufixoInteiroSingular(x As String)
 mstrDefaultSufixoInteiro1 = x & IIf(Right(x, 1) = "", "", " ")
End Property
Public Property Get SufixoInteiroPlural() As String
 SufixoInteiroPlural = mstrDefaultSufixoInteiro2
End Property
Public Property Let SufixoInteiroPlural(x As String)
 mstrDefaultSufixoInteiro2 = x & IIf(Right(x, 1) = "", "", " ")
End Property
Public Property Get SufixoDecimalSingular() As String
 SufixoDecimalSingular = mstrDefaultSufixoDecimal1
End Property
Public Property Let SufixoDecimalSingular(x As String)
 mstrDefaultSufixoDecimal1 = x & IIf(Right(x, 1) = "", "", " ")
End Property
Public Property Get SufixoDecimalPlural() As String
 SufixoDecimalPlural = mstrDefaultSufixoDecimal2
End Property
Public Property Let SufixoDecimalPlural(x As String)
 mstrDefaultSufixoDecimal2 = x & IIf(Right(x, 1) = "", "", " ")
End Property
```

