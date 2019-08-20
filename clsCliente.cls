VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim strCodigo       As Integer
Dim strCPF          As String
Dim strNome         As String
Dim strEndereco      As String
Dim strTelCelular    As String
Dim strEmail       As String
Dim strCidade      As String
Dim strUF         As String


'O CPF do Cliente seria criado e validado nesta classe de cadastro de Clientes  para
'disponibilizar, talvez, um combobox  na tela de Pedidos , para que o usuário escolha o cliente.


