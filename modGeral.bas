Attribute VB_Name = "modGeral"
Option Explicit
Public cnnConnection            As New Connection
Private strconnectionString     As String

Sub Main()

    MDIVendas.Show
    
End Sub

Public Sub AbreBanco()
    strconnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Vendas;Data Source=DESC\SQLEXPRESS"
    
With cnnConnection
    cnnConnection.ConnectionString = strconnectionString
    cnnConnection.Open
End With

End Sub


Public Sub FechaConexao()
  If cnnConnection.State = adStateOpen Then
      cnnConnection.Close
    End If
End Sub

Function ValidarCPF(CPF As String) As Boolean

On Error GoTo Err_CPF
Dim i As Integer 'utilizada nos FOR... NEXT
Dim strcampo As String 'armazena do CPF que ser� utilizada para o c�lculo
Dim strCaracter As String 'armazena os digitos do CPF da direita para a esquerda
Dim intNumero As Integer 'armazena o digito separado para c�lculo (uma a um)
Dim intMais As Integer 'armazena o digito espec�fico multiplicado pela sua base
Dim lngSoma As Long 'armazena a soma dos digitos multiplicados pela sua base(intmais)
Dim dblDivisao As Double 'armazena a divis�o dos digitos*base por 11
Dim lngInteiro As Long 'armazena inteiro da divis�o
Dim intResto As Integer 'armazena o resto
Dim intDig1 As Integer 'armazena o 1� digito verificador
Dim intDig2 As Integer 'armazena o 2� digito verificador
Dim strConf As String 'armazena o digito verificador

lngSoma = 0
intNumero = 0
intMais = 0
strcampo = Left(CPF, 9)

'Inicia c�lculos do 1� d�gito
For i = 2 To 10
    strCaracter = Right(strcampo, i - 1)
    intNumero = Left(strCaracter, 1)
    intMais = intNumero * i
    lngSoma = lngSoma + intMais
Next i
dblDivisao = lngSoma / 11

lngInteiro = Int(dblDivisao) * 11
intResto = lngSoma - lngInteiro
If intResto = 0 Or intResto = 1 Then
    intDig1 = 0
Else
    intDig1 = 11 - intResto
End If

strcampo = strcampo & intDig1 'concatena o CPF com o primeiro digito verificador
lngSoma = 0
intNumero = 0
intMais = 0
'Inicia c�lculos do 2� d�gito
For i = 2 To 11
    strCaracter = Right(strcampo, i - 1)
    intNumero = Left(strCaracter, 1)
    intMais = intNumero * i
    lngSoma = lngSoma + intMais
Next i
dblDivisao = lngSoma / 11
lngInteiro = Int(dblDivisao) * 11
intResto = lngSoma - lngInteiro
If intResto = 0 Or intResto = 1 Then
    intDig2 = 0
Else
    intDig2 = 11 - intResto
End If
strConf = intDig1 & intDig2
'Caso o CPF esteja errado dispara a mensagem
If strConf <> Right(CPF, 2) Then
    ValidarCPF = False
Else
    ValidarCPF = True
End If
Exit Function

Exit_CPF:
    Exit Function
Err_CPF:
    MsgBox Error$
    Resume Exit_CPF
End Function



