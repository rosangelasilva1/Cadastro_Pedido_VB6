VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItens 
   Caption         =   "Cadastro de Ítens"
   ClientHeight    =   3105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10305
   Icon            =   "frmItens.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   10305
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar ToolbarItens 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   1164
      ButtonWidth     =   1111
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salvar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Fechar"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraItens 
      Caption         =   "Cadastrar Itens"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   10095
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1920
         TabIndex        =   0
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtDescricao 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1920
         TabIndex        =   1
         Top             =   1320
         Width           =   7815
      End
      Begin VB.TextBox txtVlUnitario 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1920
         TabIndex        =   2
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label lblCodigo 
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblDescricao 
         Caption         =   "Descrição:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Valor Unitário:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   1920
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmItens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsItem As New clsItem
Dim clsItemPedido As New clsItemPedido
Dim rsLocalizarItem As ADODB.Recordset

Private Sub cmdExcluir_Click()
    If MsgBox("Deseja relamente excluir o Item ", vbQuestion + vbYesNo, "Exclusão de Ítens") = vbYes Then
        MsgBox "Item excluído com sucesso", vbInformation, "Exclusão de Itens"
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Top = 5500
    Me.Left = 6800
End Sub

Private Sub ToolbarItens_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo TrataErro
 
 Select Case Button.Index
      Case 1
         Call SalvarItem
      Case 2
         Call ExcluirItem
         
      Case 3 'Sair
         Unload Me
   End Select
Exit Sub
    
TrataErro:
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "Toolbar1_ButtonClick"
End Sub

Private Sub SalvarItem()

   Dim strMensagem As String
   
On Error GoTo TrataErro
    
    If ValidarCampos = False Then
        Exit Sub
    End If
    Set rsLocalizarItem = Nothing
    Set rsLocalizarItem = clsItem.rsLocalizarItem(txtCodigo.Text)
    If rsLocalizarItem.EOF Then
        strMensagem = "Ítem cadastrado com sucesso!"
    Else
        strMensagem = "Ítem alterado com sucesso!"
    End If
    Set rsLocalizarItem = Nothing
    
    
    clsItem.CodigoItem = txtCodigo.Text
    clsItem.DescricaoItem = txtDescricao.Text
    clsItem.VlUnitarioItem = txtVlUnitario.Text
    clsItem.SalvarItem clsItem
    
    MsgBox strMensagem, vbInformation, "Cadastro de Pedidos"
    
    Unload Me
Exit Sub
    
TrataErro:
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "SalvarItem"
End Sub


Private Function ValidarCampos() As Boolean
    
    ValidarCampos = True
    
On Error GoTo TrataErro

    If Not IsNumeric(txtCodigo.Text) Then
        MsgBox "Código de Inválido - Somente números", vbCritical, "Cadastro de Ítem"
        ValidarCampos = False
    ElseIf Trim(txtDescricao.Text) = "" Then
        MsgBox "Favor preencher o campo Descrição", vbCritical, "Cadastro de Ítem"
        ValidarCampos = False
    ElseIf Not IsNumeric(txtVlUnitario.Text) Then
        MsgBox "Valor Unitário Inválido", vbCritical, "Cadastro de Ítem"
        ValidarCampos = False
    End If
    
 Exit Function
    
TrataErro:
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "ValidarCampos"
End Function


Private Sub ExcluirItem()

    Dim rsLocalizarItem As New ADODB.Recordset
    
On Error GoTo TrataErro
    
    If Trim(txtCodigo.Text) = "" Then
        MsgBox "Código do Ítem inválido"
        Exit Sub
    End If
    
    Set rsLocalizarItem = Nothing
    Set rsLocalizarItem = clsItem.rsLocalizarItem(txtCodigo.Text)
    If rsLocalizarItem.EOF Then
        MsgBox "Este ítem não existe no cadastro!", vbCritical, "Exclusão de Ítem"
        Exit Sub
    End If
    
    Set rsLocalizarItem = Nothing
    Set rsLocalizarItem = clsItemPedido.rsLocalizarItemPedidoCodigo(txtCodigo.Text)
    If Not rsLocalizarItem.EOF Then
        MsgBox "O Ítem encontra-se cadastrado em Pedidos e não poderá ser excluído", vbCritical, "Exclusão de Ítem"
        Exit Sub
    End If
    
    
    If MsgBox("Deseja realmente excluir o Ítens de  nº " & txtCodigo.Text & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Exclusão de Pedido") = vbYes Then
        clsItem.ExcluirItem (txtCodigo.Text)
        MsgBox "Ítem excluído com sucesso", vbCritical, "Exclusão de Ítem"
        Unload Me
    End If
    
Exit Sub
    
TrataErro:
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "ExcluirItem"
End Sub


