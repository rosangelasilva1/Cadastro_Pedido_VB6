VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGridItens 
   Caption         =   "Cadastro de Itens"
   ClientHeight    =   5685
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10350
   Icon            =   "frmGridItens.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar ToolbarCadastroItem 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   1164
      ButtonWidth     =   1111
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Novo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Fechar"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   " Itens Cadastrados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4695
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   10095
      Begin VB.TextBox txtPesquisarItem 
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
         Left            =   6360
         MaxLength       =   5
         TabIndex        =   0
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "Pesquisar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   8280
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGridItens 
         Height          =   3375
         Left            =   360
         TabIndex        =   2
         ToolTipText     =   "Duplo clique para alterar Item"
         Top             =   1080
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   5953
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label lblCodIem 
         Caption         =   "Código do Ítem:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmGridItens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsItem As New clsItem

Private Sub cmdExcluir_Click()
    
    frmItens.Caption = "Exclusão de Itens"
    frmItens.fraItens.Caption = "Excluir Item"
    frmItens.txtCodigo.Enabled = True
    frmItens.Show 1


End Sub

Private Sub cmdFechar_Click()
    Unload Me

End Sub

Private Sub AlterarItem()
    frmItens.Caption = "Alterar Ítem"
    frmItens.fraItens.Caption = "Alterar Item"
    frmItens.txtCodigo.Enabled = False
    frmItens.Show 1
End Sub
Private Sub NovoItem()
    frmItens.Caption = "Cadastrar Novo Ítem"
    frmItens.fraItens.Caption = "Cadastrar Novo Item"
    frmItens.txtCodigo.Enabled = True
    
    frmItens.Show 1
End Sub

Private Sub cmdPesquisar_Click()
    If Trim(txtPesquisarItem.Text) = "" Then
        Call PreencherGridItensCadastrados
    End If
    
    If IsNumeric(txtPesquisarItem.Text) Then
        Call PreencherGridItensCadastrados(txtPesquisarItem.Text)
    End If
    
    
End Sub

Private Sub Form_Load()
    
    Call ConfigurarGridItens
    Call PreencherGridItensCadastrados
    
End Sub

Private Sub MSHFlexGridItens_Click()
  Dim conteudo               As String
    Dim linha              As Integer
    Dim CodigoItem      As String
    Dim Descricao       As String
    Dim ValorUnitario   As String
    
    
    
On Error GoTo TrataErro
    
    linha = MSHFlexGridItens.Row
    
    If linha = 0 Then
        Exit Sub
    End If
    
    
    frmItens.txtCodigo.Text = MSHFlexGridItens.TextMatrix(linha, 0)
    frmItens.txtDescricao.Text = MSHFlexGridItens.TextMatrix(linha, 1)
    frmItens.txtVlUnitario.Text = MSHFlexGridItens.TextMatrix(linha, 2)
    frmItens.txtCodigo.Enabled = False
    frmItens.Show 1
    Call PreencherGridItensCadastrados
       
    Exit Sub
    
TrataErro:
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "MSHFlexGridItens_Click"
End Sub
Private Sub ConfigurarGridItens()

    
On Error GoTo TrataErro

    With MSHFlexGridItens
        
        .ColWidth(0) = 1500
        .ColWidth(1) = 5500
        .ColWidth(2) = 1800
        
        .Row = 0
        .Col = 0
        .Text = "Código"
        .Col = 1
        .Text = "Descrição do Ítem"
        .Col = 2
        .Text = "Valor Unitário"
        
    End With
 Exit Sub
    
TrataErro:
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "ConfigurarGridItens"

End Sub
Private Sub PreencherGridItensCadastrados(Optional ByVal acodigoItem As Integer)

    Dim j As Integer
    Dim i As Integer
    Dim rsItens As ADODB.Recordset
    
 On Error GoTo TrataErro
 
    Set rsItens = Nothing
    Set rsItens = New ADODB.Recordset
    If acodigoItem = 0 Then
        Set rsItens = clsItem.rsRetornaItens
    Else
        Set rsItens = clsItem.rsRetornaItensCodigo(acodigoItem)
    End If
    
    i = 1
    MSHFlexGridItens.Rows = 1
    
    If Not rsItens.EOF Then
    
        MSHFlexGridItens.Rows = rsItens.RecordCount + 1
        Do While Not rsItens.EOF
        
            For j = 0 To rsItens.Fields.Count - 1
                If Not IsNull(rsItens.Fields(j).Value) Then
                    If j = 2 Then
                        MSHFlexGridItens.TextMatrix(i, j) = Format(rsItens.Fields(j).Value, "###,###,##0.00")
                    Else
                        MSHFlexGridItens.TextMatrix(i, j) = rsItens.Fields(j).Value
                    End If
                End If
            Next
            
            i = i + 1
            rsItens.MoveNext
        Loop
    
    End If
    Exit Sub
TrataErro:
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "PreencherGridItensCadastrados"

End Sub

Private Sub ToolbarCadastroItem_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo TrataErro
 
 Select Case Button.Index
      Case 1 'Novo
         Call NovoItem
         Call PreencherGridItensCadastrados

      Case 2 'Sair
         Unload Me
   End Select
Exit Sub
    
TrataErro:
    MsgBox Err.Number & " -" & Err.Description, vbCritical, "ToolbarCadastroItem_ButtonClick"
End Sub


