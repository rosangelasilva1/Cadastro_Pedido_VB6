VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private intCodigoPedido           As Integer
Private strDataPedido             As Date
Private strSolicitantePedido      As String
Private strCPFCliente             As String
Private strSituacaoPedido         As String


Dim rsAux As ADODB.Recordset

'-------------------------------------------------------------------------------------------
'Property CodigoPedido
Public Property Get CodigoPedido() As Integer
    CodigoPedido = intCodigoPedido
End Property

Public Property Let CodigoPedido(ByVal vCodigoPedido As Integer)
    intCodigoPedido = vCodigoPedido
End Property

'-------------------------------------------------------------------------------------------
'Property DataPedido

Public Property Get DataPedido() As Date
    DataPedido = strDataPedido
End Property

Public Property Let DataPedido(ByVal vDataPedido As Date)
    strDataPedido = vDataPedido
End Property

'-------------------------------------------------------------------------------------------
'Property SolicitantePedido
Public Property Get SolicitantePedido() As String
    SolicitantePedido = strSolicitantePedido
End Property

Public Property Let SolicitantePedido(ByVal vSolicitantePedido As String)
    strSolicitantePedido = vSolicitantePedido
End Property


'-------------------------------------------------------------------------------------------
'Property CPFCliente
Public Property Get CPFCliente() As Variant
    CPFCliente = strCPFCliente
End Property

Public Property Let CPFCliente(ByVal vCPFCliente As Variant)
    strCPFCliente = vCPFCliente
End Property

'-------------------------------------------------------------------------------------------
'Property SituacaoPedido
Public Property Get SituacaoPedido() As String
    SituacaoPedido = strSituacaoPedido
End Property

Public Property Let SituacaoPedido(ByVal vSituacaoPedido As String)
    strSituacaoPedido = vSituacaoPedido
End Property

Public Function getProximoNumeroPedido() As Integer
    
    Dim strSql          As String
    Dim rsPedido As ADODB.Recordset
    
On Error GoTo TrataErro

    getProximoNumeroPedido = 1
    
    strSql = "Select max(codigo) + 1 as Codigo from Pedido"
    Set rsPedido = Nothing
    Set rsPedido = New ADODB.Recordset
    
    rsPedido.Open strSql, cnnConnection, adOpenKeyset, adLockOptimistic
    
    If Not rsPedido.EOF Then
        getProximoNumeroPedido = rsPedido!Codigo
    End If
    Set rsPedido = Nothing
  Exit Function
TrataErro:
       Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Public Sub SalvarPedido(aclsPedido As clsPedido)
    
    'Verifica se Pedido existe na base de Dados. Se sim, alterar, senão,Incluir.
    
    Dim intCodigoPedido As Integer
    Dim strSql          As String
    
On Error GoTo TrataErro
    
    strSql = "Select Codigo from Pedido where Codigo = " & aclsPedido.CodigoPedido
    Set rsAux = Nothing
    Set rsAux = New ADODB.Recordset
    rsAux.Open strSql, cnnConnection, adOpenKeyset, adLockOptimistic
    
    If rsAux.EOF Then
        
        intCodigoPedido = getProximoNumeroPedido
        
        If intCodigoPedido <> 0 Then
            strSql = "INSERT INTO PEDIDO (Codigo, CPF, Solicitante, Data, Situacao, ValorTotal)" & _
                    "VALUES (" & intCodigoPedido & _
                    ",'" & Replace(Replace(aclsPedido.CPFCliente, ".", ""), "-", "") & "'" & _
                    ",'" & aclsPedido.SolicitantePedido & "'" & _
                    ",'" & aclsPedido.DataPedido & "'" & _
                    ",'" & aclsPedido.SituacaoPedido & "'" & _
                    ", 0" & ")"
                    
                    
        End If
    Else
        
       
        strSql = " UPDATE  PEDIDO " & _
                   " Set CPF ='" & Replace(Replace(aclsPedido.CPFCliente, ".", ""), "-", "") & "'," & _
                   "     Solicitante = '" & aclsPedido.SolicitantePedido & "'," & _
                   "     Data = '" & aclsPedido.DataPedido & "'," & _
                   "     Situacao = '" & aclsPedido.SituacaoPedido & "'" & _
                   " WHERE Codigo = " & aclsPedido.CodigoPedido
                
       
    End If
            
    cnnConnection.Execute strSql
    Set rsAux = Nothing
Exit Sub
TrataErro:
       Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function rsComboSituacao() As ADODB.Recordset
    
    Dim strSql          As String
    
On Error GoTo TrataErro

    strSql = "Select * from SituacaoPedido  order by Codigo"
    Set rsAux = Nothing
    Set rsAux = New ADODB.Recordset
    rsAux.Open strSql, cnnConnection, adOpenKeyset, adLockReadOnly
    
    Set rsComboSituacao = rsAux
    Set rsAux = Nothing
    
    Exit Function
TrataErro:
       Err.Raise Err.Number, Err.Source, Err.Description
    
End Function


Public Function rsLocalizarPedido(ByVal aCodigo As Integer) As ADODB.Recordset
    
    Dim strSql As String
    
On Error GoTo TrataErro
    
    strSql = "Select * from Pedido Where Codigo =  " & aCodigo
    Set rsAux = Nothing
    Set rsAux = New ADODB.Recordset
    rsAux.Open strSql, cnnConnection, adOpenKeyset, adLockReadOnly
    
    Set rsLocalizarPedido = rsAux
    
    Exit Function
TrataErro:
       Err.Raise Err.Number, Err.Source, Err.Description
    
End Function


Public Sub ExcluirPedido(ByVal aCodigoPedido As Integer)

    Dim strSql As String

On Error GoTo TrataErro

    strSql = "Delete FROM PEDIDO WHERE codigo= " & aCodigoPedido
    
    cnnConnection.Execute strSql
 
    Exit Sub
TrataErro:
       Err.Raise Err.Number, Err.Source, Err.Description
        

End Sub


'Public Sub CancelarPedido(ByVal aCodigoPedido As Integer)
'
'    Dim strSql As String
'
'On Error GoTo TrataErro
'
'    strSql = "UPDATE  PEDIDO  SET SITUACAO = 3 WHERE codigo= " & aCodigoPedido
'
'    cnnConnection.Execute strSql
'
'    Exit Sub
'TrataErro:
'       Err.Raise Err.Number, Err.Source, Err.Description
'
'
'End Sub
'
'Public Function rsRetornaSituacaoPedido(ByVal aCodigo As Integer) As ADODB.Recordset
'
'    Dim strSql As String
'
'On Error GoTo TrataErro
'
'    strSql = "Select Situacao from Pedido Where Codigo =  " & aCodigo
'    Set rsAux = Nothing
'    Set rsAux = New ADODB.Recordset
'    rsAux.Open strSql, cnnConnection, adOpenKeyset, adLockReadOnly
'
'    Set rsRetornaSituacaoPedido = rsAux
'
'    Exit Function
'TrataErro:
'       Err.Raise Err.Number, Err.Source, Err.Description
'
'End Function

