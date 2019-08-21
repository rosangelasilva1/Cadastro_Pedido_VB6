VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPedidos 
   Caption         =   "Cadastro de Pedidos"
   ClientHeight    =   7980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13485
   Icon            =   "frmPedidos.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   13485
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageListPedidos 
      Left            =   480
      Top             =   7920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   413
      ImageHeight     =   310
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPedidos.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   13485
      _ExtentX        =   23786
      _ExtentY        =   1164
      ButtonWidth     =   1402
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Novo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salvar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Fechar"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraValorTotal 
      Height          =   975
      Left            =   240
      TabIndex        =   20
      Top             =   6840
      Width           =   12855
      Begin VB.TextBox txtValorTotalPedido 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   9720
         TabIndex        =   11
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Valor Total do Pedido"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   7080
         TabIndex        =   21
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame fraItensPedido 
      Caption         =   "Itens do Pedido"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3615
      Left            =   240
      TabIndex        =   17
      Top             =   3120
      Width           =   12855
      Begin VB.TextBox txtValorTotalItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   10320
         TabIndex        =   8
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox txtValorUnitario 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8520
         TabIndex        =   7
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox txtQuantidadeItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7200
         MaxLength       =   4
         TabIndex        =   6
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtDescricaoItem 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1560
         TabIndex        =   5
         Top             =   2520
         Width           =   5535
      End
      Begin VB.TextBox txtCodigoItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         MaxLength       =   4
         TabIndex        =   4
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton cmdNovoItemPedido 
         Caption         =   "Novo Ítem"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Limpa os campos do Ítem para Inclusão de novo Ítem"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "Excluir Ítem"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11040
         TabIndex        =   10
         ToolTipText     =   "Exclui um ítem selecionado"
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdSalvarItem 
         Caption         =   "Salvar Ítem"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9720
         TabIndex        =   9
         ToolTipText     =   "Salva Novo ítem ou Altera ìtem selecionado"
         Top             =   3000
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGridItens 
         Height          =   2055
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Dê um clique para selecionar um ítem para Alteração ou Exclusão"
         Top             =   360
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
   End
   Begin VB.Frame fraPedido 
      Caption         =   "Cadastrar Pedidos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2295
      Left            =   240
      TabIndex        =   12
      Top             =   720
      Width           =   12855
      Begin VB.ComboBox cboSituacao 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8880
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   720
         Width           =   2775
      End
      Begin MSMask.MaskEdBox mskCPF 
         Height          =   375
         Left            =   2400
         TabIndex        =   1
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###.###.###-##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtSolicitante 
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
         Height          =   390
         Left            =   240
         MaxLength       =   200
         TabIndex        =   3
         Top             =   1680
         Width           =   9975
      End
      Begin MSComCtl2.DTPicker dtpData 
         Height          =   375
         Left            =   4800
         TabIndex        =   2
         Top             =   720
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   105316353
         CurrentDate     =   43695
      End
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
         Height          =   390
         Left            =   240
         MaxLength       =   4
         TabIndex        =   0
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Situação"
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
         Left            =   8880
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "CPF"
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
         Left            =   2400
         TabIndex        =   16
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Solicitante"
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
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblData 
         Caption         =   "Data"
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
         Left            =   4800
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblCodigo 
         Caption         =   "Código"
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
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim clsPedido               As New clsPedido
Dim clsItemPedido           As New clsItemPedido
Dim clsItem                 As New clsItem
Dim rsLocalizarPedido       As New ADODB.Recordset
Dim rsLocalizarItemPedido   As New ADODB.Recordset
'Obs.: Tive uma dúvida referente o Botão Cancelar.



Private Sub ConfigurarGrid()
    
On Error GoTo TrataErro

    With MSHFlexGridItens
        
        .ColWidth(0) = 1500
        .ColWidth(1) = 5500
        .ColWidth(2) = 1300
        .ColWidth(3) = 1800
        .ColWidth(4) = 2000
        
        .Row = 0
        .Col = 0
        .Text = "Código"
        .Col = 1
        .Text = "Descrição do Ítem"
        .Col = 2
        .Text = "Quantidade"
        .Col = 3
        .Text = "Valor Unitário"
        .Col = 4
        .Text = "Valor Total Ítem"
        
    End With
 Exit Sub
    
TrataErro:
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "ConfigurarGrid"

End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub

Private Sub cmdExcluir_Click()

On Error GoTo TrataErro

 If MsgBox("Deseja realmente excluir o ítem de Pedido   nº " & txtCodigoItem.Text & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Exclusão de Pedido") = vbYes Then
        Call clsItemPedido.ExcluirItemPedido(txtCodigo.Text, txtCodigoItem.Text)
        LimparCamposItem
        Call PreencherGridItemPedidos
        Call AtualizarValorTotalPedido
    End If
  Exit Sub
    
TrataErro:
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "cmdExcluir_Click"
End Sub

Private Sub cmdNovoItemPedido_Click()
    
      
On Error GoTo TrataErro
    
    Set rsLocalizarPedido = Nothing
    Set rsLocalizarPedido = clsPedido.rsLocalizarPedido(txtCodigo.Text)
    
    If rsLocalizarPedido.EOF Then
        MsgBox "O pedido nº " & txtCodigo.Text & " não foi cadastrado." & vbCrLf & _
        "Para cadastrar novo ítem, é necessário primeiro cadastrar o Pedido.", vbInformation, "Cadastrar Ítem"
         Exit Sub
    End If
    Set rsLocalizarPedido = Nothing
    
    Call LimparCamposItem
    txtCodigoItem.SetFocus

    Exit Sub
    
TrataErro:
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "cmdNovoItemPedido_Click"
    
End Sub

Private Sub cmdSalvarItem_Click()

    Dim strMensagem As String
    
On Error GoTo TrataErro

    Set rsLocalizarPedido = clsPedido.rsLocalizarPedido(txtCodigo.Text)
    
    If rsLocalizarPedido.EOF Then
        MsgBox "O pedido nº " & txtCodigo.Text & " não foi cadastrado." & vbCrLf & _
        "Para cadastrar novo ítem, é necessário primeiro cadastrar o Pedido.", vbInformation, "Cadastrar Ítem"
         Set rsLocalizarPedido = Nothing
         Exit Sub
    End If
    Set rsLocalizarPedido = Nothing

    
    If ValidarCamposItemPedido = False Then
       txtCodigoItem.SetFocus
       Exit Sub
    End If
    
    Dim rsLocalizarItem As New ADODB.Recordset
   If txtCodigoItem.Text <> "" Then
        Set rsLocalizarItem = clsItemPedido.rsLocalizarItemPedido(txtCodigo.Text, txtCodigoItem.Text)
        If Not rsLocalizarItem.EOF Then
           strMensagem = "Item alterado com sucesso "
        Else
            strMensagem = "Item cadastrado com sucesso "
        End If
    End If
    
    Set rsLocalizarItem = Nothing
        
    clsItemPedido.CodigoItem = txtCodigoItem
    clsItemPedido.CodigoPedido = txtCodigo.Text
    clsItemPedido.Quantidade = txtQuantidadeItem.Text
    clsItemPedido.ValorUnitarioItem = CDbl(txtValorUnitario.Text)
    clsItemPedido.ValorTotalItem = CDbl(txtValorTotalItem.Text)
    
    clsItemPedido.SalvarItemPedido clsItemPedido
    
    Call LimparCamposItem
    
    Call AtualizarValorTotalPedido
    
    Call PreencherGridItemPedidos
    
    MsgBox strMensagem, vbInformation, "Ítens de Pedido"
    
    txtCodigoItem.SetFocus
    
 Exit Sub
    
TrataErro:
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "cmdSalvarItem_Click"
        
End Sub

Private Sub Form_Load()

On Error GoTo TrataErro

   CarregarComboSituacao
   ConfigurarGrid
   
   NovoPedido
   
  Exit Sub
    
TrataErro:
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "Form_Load"
   
End Sub

Private Sub MSHFlexGridItens_Click()
    
    Dim linha              As Integer
    
On Error GoTo TrataErro
    
    linha = MSHFlexGridItens.Row
    
    If linha = 0 Then
        Exit Sub
    End If
    
    txtCodigoItem.Text = MSHFlexGridItens.TextMatrix(linha, 0)
    txtDescricaoItem.Text = MSHFlexGridItens.TextMatrix(linha, 1)
    txtQuantidadeItem.Text = MSHFlexGridItens.TextMatrix(linha, 2)
    txtValorUnitario.Text = MSHFlexGridItens.TextMatrix(linha, 3)
    txtValorTotalItem.Text = MSHFlexGridItens.TextMatrix(linha, 4)
       
    Exit Sub
    
TrataErro:
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "MSHFlexGridItens_Click"

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 
On Error GoTo TrataErro
 
 Select Case Button.Index
      Case 1 'Novo
         Call NovoPedido
         
         Case 2 'Salvar
         Call SalvarPedido
         
        Case 3 'Excluir
         Call ExcluirPedido
        
        Case 4 'Cancelar Pedido
         'Call CancelarPedido


      Case 6 'Sair
         Unload Me
   End Select
Exit Sub
    
TrataErro:
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "Toolbar1_ButtonClick"
End Sub


Private Sub NovoPedido()
    
    Dim intNovoNumeroPedido As Integer
    
On Error GoTo TrataErro

    Call LimparCamposItem
    Call LimparCamposPedido
    
   
    'Busca Próximo número de Pedido
    intNovoNumeroPedido = clsPedido.getProximoNumeroPedido
    txtCodigo.Text = intNovoNumeroPedido
    'cboSituacao.ListIndex = 0
    
Exit Sub
    
TrataErro:
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "NovoPedido"
    
End Sub
Private Sub LimparCamposItem()

On Error GoTo TrataErro
    txtCodigoItem.Text = ""
    txtDescricaoItem.Text = ""
    txtQuantidadeItem.Text = ""
    txtValorUnitario.Text = ""
    txtValorTotalItem.Text = ""
    txtValorTotalPedido.Text = ""
Exit Sub
    
TrataErro:
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "LimparCamposItem"
End Sub


Private Sub LimparCamposPedido()
    
On Error GoTo TrataErro

    mskCPF.Text = "___.___.___-__"
    txtSolicitante.Text = ""
    MSHFlexGridItens.Rows = 1
    dtpData.Value = Now

Exit Sub
    
TrataErro:
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "LimparCamposPedido"

End Sub

Private Sub SalvarPedido()

    Dim strMensagem As String
    Dim rsSituacao As ADODB.Recordset
    
On Error GoTo TrataErro
    
    If ValidarCampos = False Then
        Exit Sub
    End If
    
    Set rsLocalizarPedido = Nothing
    Set rsLocalizarPedido = clsPedido.rsLocalizarPedido(txtCodigo.Text)
    If rsLocalizarPedido.EOF Then
        strMensagem = "Pedido cadastrado com sucesso!"
        clsPedido.SituacaoPedido = 1 ' Pendente
    Else
        Set rsSituacao = Nothing
        Set rsSituacao = clsPedido.rsRetornaSituacaoPedido(txtCodigo.Text)
        
        If Not rsSituacao.EOF Then
            If rsSituacao!Situacao = 3 Then
                Set rsSituacao = Nothing
                Set rsLocalizarPedido = Nothing
                MsgBox " O Pedido está cancelado, não poderá ser alterado", vbInformation, "Cadastro de Pedidos"
                Exit Sub
            End If
        End If
        clsPedido.SituacaoPedido = cboSituacao.ItemData(cboSituacao.ListIndex)
        strMensagem = "Pedido alterado com sucesso!"
    End If
    Set rsLocalizarPedido = Nothing
    Set rsSituacao = Nothing
    
    clsPedido.CodigoPedido = txtCodigo.Text
    clsPedido.CPFCliente = mskCPF
    clsPedido.DataPedido = dtpData
    
    clsPedido.SolicitantePedido = txtSolicitante.Text
    clsPedido.SalvarPedido clsPedido
    
    MsgBox strMensagem, vbInformation, "Cadastro de Pedidos"
        
    
Exit Sub
    
TrataErro:
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "SalvarPedido"

End Sub

Private Function ValidarCampos() As Boolean
    
    ValidarCampos = True
    
On Error GoTo TrataErro

    If Not IsNumeric(txtCodigo.Text) Then
        MsgBox "Código de Pedido Inválido", vbCritical, "Cadastro de Pedido"
        ValidarCampos = False
    ElseIf mskCPF.Text = "___.___.___-__" Then
        MsgBox "Favor Preencher o CPF", vbCritical, "Cadastro de Pedido"
        ValidarCampos = False
    ElseIf ValidarCPF(Replace(Replace(Replace(mskCPF.Text, ".", ""), "-", ""), "_", "")) = False Then
        MsgBox "CPF Inválido", vbCritical, "Cadastro de Pedido"
        ValidarCampos = False
    ElseIf Trim(txtSolicitante.Text) = "" Then
        MsgBox "Favor preencher o campo Solicitante", vbCritical, "Cadastro de Pedido"
        ValidarCampos = False
        
    End If
    
 Exit Function
    
TrataErro:
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "ValidarCampos"
End Function

Private Sub CarregarComboSituacao()

    Dim rsComboSituacao As New ADODB.Recordset

On Error GoTo TrataErro
    
    Set rsComboSituacao = Nothing
    Set rsComboSituacao = clsPedido.rsComboSituacao
    
    Do While Not rsComboSituacao.EOF
        
        With cboSituacao
            .AddItem rsComboSituacao!Codigo & "-" & rsComboSituacao!Descricao
            .ItemData(.NewIndex) = rsComboSituacao!Codigo
        End With
        rsComboSituacao.MoveNext
    Loop
    
    Set rsComboSituacao = Nothing
    cboSituacao.ListIndex = 0
    
     Exit Sub
    
TrataErro:
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "CarregarComboSituacao"
    
End Sub

Private Sub txtCodigo_LostFocus()
    
On Error GoTo TrataErro
    Set rsLocalizarPedido = Nothing
    Set rsLocalizarPedido = clsPedido.rsLocalizarPedido(txtCodigo.Text)
    LimparCamposPedido
    'Se Pedido existe, busca campos do BD e mostra na tela
    If Not rsLocalizarPedido.EOF Then
           
        If rsLocalizarPedido!CPF <> "___.___.___-__" Then
            mskCPF.Text = Mid(rsLocalizarPedido!CPF, 1, 3) & "." & Mid(rsLocalizarPedido!CPF, 4, 3) & "." & Mid(rsLocalizarPedido!CPF, 7, 3) & "-" & Mid(rsLocalizarPedido!CPF, 10, 2)
        End If
        dtpData = rsLocalizarPedido!Data
        'cboSituacao.ListIndex = rsLocalizarPedido!Situacao - 1
        txtSolicitante.Text = rsLocalizarPedido!Solicitante
        
        Call AtualizarValorTotalPedido
        
        Call PreencherGridItemPedidos

    Else
       Call NovoPedido
    End If
    Set rsLocalizarPedido = Nothing
    
     Exit Sub
    
TrataErro:
    Set rsLocalizarPedido = Nothing
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "txtCodigo_LostFocus"
    
End Sub

Private Sub txtCodigoItem_LostFocus()
        
   Dim rsLocalizarItem As New ADODB.Recordset

On Error GoTo TrataErro

    txtDescricaoItem = ""
    txtValorUnitario = ""
    txtQuantidadeItem = ""
    txtValorTotalItem = ""
    
    If txtCodigoItem.Text <> "" Then
        Set rsLocalizarItem = Nothing
        Set rsLocalizarItem = clsItem.rsLocalizarItem(txtCodigoItem.Text)
        If Not rsLocalizarItem.EOF Then
            txtDescricaoItem = rsLocalizarItem!Descricao
            txtValorUnitario = Format(rsLocalizarItem!ValorUnitario, "###,###,##0.00")
        End If
    End If
    Set rsLocalizarItem = Nothing
    
    txtQuantidadeItem.SetFocus
    
   Exit Sub
    
TrataErro:
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "txtCodigoItem_LostFocus"
    
End Sub

Private Sub txtQuantidadeItem_LostFocus()
    
On Error GoTo TrataErro
    
    If Trim(txtQuantidadeItem.Text) <> "" And Trim(txtValorUnitario.Text) <> "" Then
        txtValorTotalItem.Text = Format(CDbl(txtQuantidadeItem.Text) * CDbl(txtValorUnitario.Text), "###,###,##0.00")
    End If
    
   Exit Sub
    
TrataErro:
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "txtQuantidadeItem_LostFocus"
    
End Sub

Private Function ValidarCamposItemPedido() As Boolean
    
    ValidarCamposItemPedido = True
    
On Error GoTo TrataErro
    
    If Trim(txtCodigoItem.Text) = "" Then
        MsgBox "Favor preencher o campo Código do Ítem", vbCritical, "Cadastro de Ítens de Pedido"
        ValidarCamposItemPedido = False
    ElseIf Trim(txtDescricaoItem.Text) = "" Then
        MsgBox "Ítem não cadastrado", vbCritical, "Cadastro de Ítens de Pedido"
        ValidarCamposItemPedido = False
    ElseIf Trim(txtQuantidadeItem.Text) = "" Then
        MsgBox "Favor digitar a quantidade do Ítem", vbCritical, "Cadastro de Ítens de Pedido"
        ValidarCamposItemPedido = False
    ElseIf Trim(txtValorUnitario.Text) = "" Then
        MsgBox "Ocorreu um problema com o Valor Unitário do Ítem", vbCritical, "Cadastro de Ítens de Pedido"
        ValidarCamposItemPedido = False
    ElseIf Trim(txtValorTotalItem.Text) = "" Then
        MsgBox "Ocorreu um problema com o Valor Total do Ítem", vbCritical, "Cadastro de Ítens de Pedido"
        ValidarCamposItemPedido = False
    End If
    
    Exit Function
    
TrataErro:
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "ValidarCamposItemPedido"
    
End Function

Private Sub AtualizarValorTotalPedido()

On Error GoTo TrataErro
    
    txtValorTotalPedido.Text = Format(clsItemPedido.getValorTotalItemPedido(txtCodigo.Text), "###,###,##0.00")
    Exit Sub
TrataErro:
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "AtualizarValorTotalPedido"
    
End Sub

Private Sub PreencherGridItemPedidos()

    Dim j As Integer
    Dim i As Integer
    Dim rsItenPedido As New ADODB.Recordset
    
 On Error GoTo TrataErro
 
    Set rsItenPedido = Nothing
    Set rsItenPedido = clsItemPedido.rsRetornaItensPedido(txtCodigo.Text)
    
    i = 1
    MSHFlexGridItens.Rows = 1
    
    If Not rsItenPedido.EOF Then
        
        MSHFlexGridItens.Rows = rsItenPedido.RecordCount + 1
        Do While Not rsItenPedido.EOF
        
        For j = 0 To rsItenPedido.Fields.Count - 1
            If Not IsNull(rsItenPedido.Fields(j).Value) Then
                If (j = 3 Or j = 4) Then
                    MSHFlexGridItens.TextMatrix(i, j) = Format(rsItenPedido.Fields(j).Value, "###,###,##0.00")
                Else
                    MSHFlexGridItens.TextMatrix(i, j) = rsItenPedido.Fields(j).Value
                End If
            End If
        Next
        
        i = i + 1
        rsItenPedido.MoveNext
        Loop

    End If
    Exit Sub
TrataErro:
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "PreencherGridItemPedidos"
    
End Sub

Private Sub ExcluirPedido()
    
On Error GoTo TrataErro
    
    If Trim(txtCodigo.Text) = "" Then
        MsgBox "Favor digitar o código do Pedido!", vbCritical, "Exclusão de Pedidos"
        Exit Sub
    End If
    
    If Not IsNumeric(txtCodigo.Text) Then
        MsgBox "Código do Pedido inválido!", vbCritical, "Exclusão de Pedidos"
        Exit Sub
    End If

    Set rsLocalizarPedido = Nothing
    Set rsLocalizarPedido = clsPedido.rsLocalizarPedido(txtCodigo.Text)
    If rsLocalizarPedido.EOF Then
        MsgBox "Pedido não cadastrado!", vbCritical, "Exclusão de Pedidos"
        Exit Sub
    End If
    Set rsLocalizarPedido = Nothing
    
    If MsgBox("Deseja realmente excluir o Pedido e Ítens de  nº " & txtCodigo.Text & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Exclusão de Pedido") = vbYes Then
        If cnnConnection.State = adStateOpen Then
            cnnConnection.BeginTrans
        End If
        
        Call clsItemPedido.ExcluirItemPedido(txtCodigo.Text)
        Call clsPedido.ExcluirPedido(txtCodigo.Text)
        
        cnnConnection.CommitTrans
        
        LimparCamposPedido
        LimparCamposItem
        Call NovoPedido
    End If
    
    Exit Sub
TrataErro:
    
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "ExcluirPedido"
    cnnConnection.RollbackTrans
    If False Then
        Resume Next
    End If
End Sub

Private Sub CancelarPedido()
    
On Error GoTo TrataErro
    
    If Trim(txtCodigo.Text) = "" Then
        MsgBox "Favor digitar o código do Pedido!", vbCritical, "Exclusão de Pedidos"
        Exit Sub
    End If
    
    If Not IsNumeric(txtCodigo.Text) Then
        MsgBox "Código do Pedido inválido!", vbCritical, "Exclusão de Pedidos"
        Exit Sub
    End If

    Set rsLocalizarPedido = Nothing
    Set rsLocalizarPedido = clsPedido.rsLocalizarPedido(txtCodigo.Text)
    If rsLocalizarPedido.EOF Then
        MsgBox "Pedido não cadastrado!", vbCritical, "Exclusão de Pedidos"
        Exit Sub
    End If
    
    If MsgBox("Deseja realmente cancelar o Pedido e Ítens de  nº " & txtCodigo.Text & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Cancelar Pedido") = vbYes Then
        'Call clsPedido.CancelarPedido(txtCodigo.Text)
        LimparCamposPedido
        LimparCamposItem
        Call NovoPedido
    End If
    
    Exit Sub
TrataErro:
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "ExcluirPedido"
End Sub
