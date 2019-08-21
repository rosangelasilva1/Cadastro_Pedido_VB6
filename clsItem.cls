VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private intCodigoItem           As Integer
Private strDescricaoItem        As String
Private dblVlUnitarioItem       As Double

Dim rsAux As ADODB.Recordset


'-------------------------------------------------------------------------------------------
'Property Codigo
Public Property Get CodigoItem() As Integer
    CodigoItem = intCodigoItem
End Property

Public Property Let CodigoItem(ByVal vCodigoItem As Integer)
    intCodigoItem = vCodigoItem
End Property


'-------------------------------------------------------------------------------------------
'Property Descrição
Public Property Get DescricaoItem() As String
    DescricaoItem = strDescricaoItem
End Property

Public Property Let DescricaoItem(ByVal vDescricaoItem As String)
    strDescricaoItem = vDescricaoItem
End Property


'-------------------------------------------------------------------------------------------
'Property Valor unitário

Public Property Get VlUnitarioItem() As Double
    VlUnitarioItem = dblVlUnitarioItem
    
End Property

Public Property Let VlUnitarioItem(ByVal vVlUnitarioItem As Double)
    dblVlUnitarioItem = vVlUnitarioItem
    
End Property


Public Function rsLocalizarItem(ByVal acodigoItem As Integer) As ADODB.Recordset
    
    Dim strSql                      As String
    
On Error GoTo TrataErro
    
    strSql = "Select * from Item Where Codigo =  " & acodigoItem
    Set rsAux = Nothing
    Set rsAux = New ADODB.Recordset
    rsAux.Open strSql, cnnConnection, adOpenKeyset, adLockReadOnly
    
    Set rsLocalizarItem = rsAux
    
Exit Function
TrataErro:
       Err.Raise Err.Number, Err.Source, Err.Description
     
    
End Function

Public Function rsRetornaItens() As ADODB.Recordset
    
    Dim strSql                     As String
    Dim rsItem                     As ADODB.Recordset
    
On Error GoTo TrataErro
    
    strSql = "SELECT  Codigo," & vbCrLf & _
    "   Descricao," & vbCrLf & _
    "   ValorUnitario" & vbCrLf & _
    "   From Item " & vbCrLf & _
    "   order by Codigo"
            
    Set rsItem = Nothing
    Set rsItem = New ADODB.Recordset
    rsItem.Open strSql, cnnConnection, adOpenKeyset, adLockOptimistic
    
    Set rsRetornaItens = rsItem
    
 Exit Function
TrataErro:
       Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Public Function rsRetornaItensCodigo(ByVal acodigoItem As Integer) As ADODB.Recordset
    
    Dim strSql                     As String
    Dim rsItem                     As ADODB.Recordset
    
On Error GoTo TrataErro
    
    strSql = "SELECT  Codigo," & vbCrLf & _
    "   Descricao," & vbCrLf & _
    "   ValorUnitario" & vbCrLf & _
    "   From Item " & vbCrLf & _
    "   WHERE Codigo=" & acodigoItem
            
    Set rsItem = Nothing
    Set rsItem = New ADODB.Recordset
    rsItem.Open strSql, cnnConnection, adOpenKeyset, adLockOptimistic
    
    Set rsRetornaItensCodigo = rsItem
    
 Exit Function
TrataErro:
       Err.Raise Err.Number, Err.Source, Err.Description
            
    
End Function


Public Sub SalvarItem(ByVal aclsItem As clsItem)
    
    Dim rsItem      As ADODB.Recordset
    Dim strSQLItem  As String
    
On Error GoTo TrataErro

    'Verifica se Ítem existe
    strSQLItem = "Select Codigo from Item where Codigo = " & aclsItem.CodigoItem
    Set rsAux = Nothing
    Set rsAux = New ADODB.Recordset
    rsAux.Open strSQLItem, cnnConnection, adOpenKeyset, adLockOptimistic
    
    'Se Item  existe
    
    If Not rsAux.EOF Then
        strSQLItem = "UPDATE  ITEM  SET  DESCRICAO ='" & aclsItem.DescricaoItem & "',ValorUnitario=" & Replace(aclsItem.VlUnitarioItem, ",", ".") & _
        " where Codigo = " & aclsItem.CodigoItem
    Else
        'Insere o Item
        strSQLItem = "INSERT INTO ITEM ( Codigo, Descricao,  ValorUnitario)" & _
                "VALUES (" & aclsItem.CodigoItem & _
                ",'" & Replace(aclsItem.DescricaoItem, ",", ".") & _
                "'," & Replace(aclsItem.VlUnitarioItem, ",", ".") & _
                ")"
    End If
    cnnConnection.Execute strSQLItem
    Set rsAux = Nothing
    
    Exit Sub
TrataErro:
       Err.Raise Err.Number, Err.Source, Err.Description
        
    
    
End Sub

Public Sub ExcluirItem(ByVal acodigoItem As Integer)

    Dim strSql As String
    
On Error GoTo TrataErro
    
    strSql = "Delete FROM Item WHERE codigo= " & acodigoItem
    
    cnnConnection.Execute strSql

Exit Sub
TrataErro:
       Err.Raise Err.Number, Err.Source, Err.Description
        
End Sub
