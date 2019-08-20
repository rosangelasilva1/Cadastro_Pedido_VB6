VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItemPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private intCodigoItem        As Integer
Private intCodigoPedido      As Integer
Private intQuantidade        As Integer
Private dblValorUnitarioItem As Double
Private dblValorTotalItem    As Double

Dim rsItemPedido As ADODB.Recordset
Dim rsAux As ADODB.Recordset

'------------------------------------------------------------------------------------------------------------
'Property CodigoItem
Public Property Get CodigoItem() As Integer
    CodigoItem = intCodigoItem
End Property

Public Property Let CodigoItem(ByVal vCodigoItem As Integer)
    intCodigoItem = vCodigoItem
End Property

'------------------------------------------------------------------------------------------------------------
'Property CodigoPedido
Public Property Get CodigoPedido() As Integer
    CodigoPedido = intCodigoPedido
End Property

Public Property Let CodigoPedido(ByVal vCodigoPedido As Integer)
    intCodigoPedido = vCodigoPedido
End Property

'------------------------------------------------------------------------------------------------------------
'Property  Quantidade
Public Property Get Quantidade() As Integer
Quantidade = intQuantidade
End Property

Public Property Let Quantidade(ByVal vQuantidade As Integer)
intQuantidade = vQuantidade
End Property


'------------------------------------------------------------------------------------------------------------
'Property  ValorUnitarioItem
Public Property Get ValorUnitarioItem() As Double
    ValorUnitarioItem = dblValorUnitarioItem
End Property

Public Property Let ValorUnitarioItem(ByVal vValorUnitarioItem As Double)
    dblValorUnitarioItem = vValorUnitarioItem
End Property


'------------------------------------------------------------------------------------------------------------
'Property  ValorTotalItem
Public Property Get ValorTotalItem() As Double
    ValorTotalItem = dblValorTotalItem
End Property

Public Property Let ValorTotalItem(ByVal vValorTotalItem As Double)
    dblValorTotalItem = vValorTotalItem
End Property

Public Sub SalvarItemPedido(aclsItemPedido As clsItemPedido)
 
    Dim rsItemPedido As ADODB.Recordset
    Dim strSQLItem As String
    
On Error GoTo TrataErro
    
    'Verifica se Pedido existe na base
    strSQLItem = "Select Codigo from Pedido where Codigo = " & aclsItemPedido.CodigoPedido
    Set rsAux = Nothing
    Set rsAux = New ADODB.Recordset
    rsAux.Open strSQLItem, cnnConnection, adOpenKeyset, adLockOptimistic
    
    'Se Pedido existe
    If Not rsAux.EOF Then
            
            'cnnConnection.BeginTrans
            Set rsItemPedido = Nothing
            Set rsItemPedido = rsLocalizarItemPedido(aclsItemPedido.CodigoPedido, aclsItemPedido.CodigoItem)
            If Not rsItemPedido.EOF Then
                    'Altera somente a quantidade do ítem
                    strSQLItem = "UPDATE  ITEMPEDIDO  SET  QUANTIDADE =" & aclsItemPedido.Quantidade & ",ValorUnitarioItem=" & Replace(aclsItemPedido.ValorUnitarioItem, ",", ".") & ", ValorTotalItem=" & Replace(aclsItemPedido.ValorTotalItem, ",", ".") & _
                    " WHERE CodigoPedido = " & aclsItemPedido.CodigoPedido & _
                    " AND CodigoItem = " & aclsItemPedido.CodigoItem
                
            Else
                'Insere o Item de Pedido
                strSQLItem = "INSERT INTO ITEMPEDIDO ( CodigoItem, CodigoPedido, Quantidade, ValorUnitarioItem, ValorTotalItem)" & _
                        "VALUES (" & aclsItemPedido.CodigoItem & _
                         "," & aclsItemPedido.CodigoPedido & _
                        "," & aclsItemPedido.Quantidade & _
                        "," & Replace(aclsItemPedido.ValorUnitarioItem, ",", ".") & _
                        "," & Replace(aclsItemPedido.ValorTotalItem, ",", ".") & _
                        ")"
            End If
            cnnConnection.Execute strSQLItem
            
            'Atualiza o Valor Total do Pedido
            strSQLItem = " UPDATE  PEDIDO " & _
                   " Set ValorTotal =" & Replace(getValorTotalItemPedido(aclsItemPedido.CodigoPedido), ",", ".") & _
                   " WHERE Codigo = " & aclsItemPedido.CodigoPedido
            
            cnnConnection.Execute strSQLItem
            
            'cnnConnection.CommitTrans
  
    End If
      Exit Sub
TrataErro:
       Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub


Public Function getValorTotalItemPedido(ByVal aCodigoPedido As Integer) As Double
    
    Dim strSql                     As String
    
On Error GoTo TrataErro

    getValorTotalItemPedido = 0
    
   strSql = "SELECT isnull(SUM( i.ValorTotalItem),0) as ValorTotalItem" & _
            " FROM ItemPedido i" & _
            " WHERE i.CodigoPedido =" & aCodigoPedido
    Set rsItemPedido = Nothing
    Set rsItemPedido = New ADODB.Recordset
    
    rsItemPedido.Open strSql, cnnConnection, adOpenKeyset, adLockOptimistic
    
    If Not rsItemPedido.EOF Then
            getValorTotalItemPedido = rsItemPedido!ValorTotalItem
    End If
    
      Exit Function
TrataErro:
       Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Public Function rsLocalizarItemPedido(ByVal aCodigoPedido As Integer, acodigoItem As Integer) As ADODB.Recordset
    
    Dim strSql                     As String
On Error GoTo TrataErro
    strSql = "Select CodigoItem from ItemPedido Where CodigoPedido =  " & aCodigoPedido & " and  CodigoItem = " & acodigoItem
    Set rsAux = Nothing
    Set rsAux = New ADODB.Recordset
    rsAux.Open strSql, cnnConnection, adOpenKeyset, adLockReadOnly
    
    Set rsLocalizarItemPedido = rsAux
    Exit Function
TrataErro:
       Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function rsLocalizarItemPedidoCodigo(ByVal acodigoItem As Integer) As ADODB.Recordset
    
    Dim strSql                     As String

On Error GoTo TrataErro
    strSql = "Select CodigoItem from ItemPedido Where   CodigoItem = " & acodigoItem
    Set rsAux = Nothing
    Set rsAux = New ADODB.Recordset
    rsAux.Open strSql, cnnConnection, adOpenKeyset, adLockReadOnly
    
    Set rsLocalizarItemPedidoCodigo = rsAux
      Exit Function
TrataErro:
       Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function rsRetornaItensPedido(ByVal aCodigoPedido As Integer) As ADODB.Recordset
    
    Dim strSql                     As String
 
 On Error GoTo TrataErro
 
    strSql = "SELECT  CodigoItem," & vbCrLf & _
    "   Descricao," & vbCrLf & _
    "   Quantidade," & vbCrLf & _
    "   ValorUnitarioItem ," & vbCrLf & _
    "   ValorTotalItem" & vbCrLf & _
    "   From ItemPedido" & vbCrLf & _
    "   inner join Item  on ItemPedido.CodigoItem  = Item.Codigo " & vbCrLf & _
    "   Where ItemPedido.CodigoPedido = " & aCodigoPedido & vbCrLf & _
    "   order by CodigoItem"
            
    Set rsItemPedido = Nothing
    Set rsItemPedido = New ADODB.Recordset
    rsItemPedido.Open strSql, cnnConnection, adOpenKeyset, adLockOptimistic
    
    Set rsRetornaItensPedido = rsItemPedido
  Exit Function
TrataErro:
       Err.Raise Err.Number, Err.Source, Err.Description
    
End Function


Public Sub ExcluirItemPedido(ByVal aCodigoPedido As Integer, Optional acodigoItem As Integer)
    
    Dim strSql As String
    
On Error GoTo TrataErro

    strSql = "Delete FROM ITEMPEDIDO WHERE codigoPedido= " & aCodigoPedido
    
    If acodigoItem > 0 Then
        strSql = strSql & " AND codigoItem= " & acodigoItem
    End If
    
    cnnConnection.Execute strSql
    
      Exit Sub
TrataErro:
       Err.Raise Err.Number, Err.Source, Err.Description
End Sub

