VERSION 5.00
Begin VB.MDIForm MDIVendas 
   BackColor       =   &H8000000C&
   Caption         =   "Vendas"
   ClientHeight    =   6495
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   12225
   Icon            =   "MDIVendas.frx":0000
   LinkTopic       =   "Vendas"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuCadastro 
      Caption         =   "Cadastro"
      Begin VB.Menu mnuCadItens 
         Caption         =   "Itens"
      End
      Begin VB.Menu mnuCadPedidos 
         Caption         =   "Pedidos"
      End
   End
   Begin VB.Menu mnuSair 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "MDIVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
AbreBanco
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    FechaConexao
End Sub

Private Sub mnuCadItens_Click()
    frmGridItens.Show 1
End Sub

Private Sub mnuCadPedidos_Click()
    frmPedidos.Show 1
End Sub

Private Sub mnuSair_Click()
    Unload Me
End Sub
