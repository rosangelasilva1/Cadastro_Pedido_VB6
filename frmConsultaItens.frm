VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmConsultaItens 
   Caption         =   "Cadastro de Itens"
   ClientHeight    =   7665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14490
   Icon            =   "frmConsultaItens.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   14490
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Itens Cadastrados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13815
      Begin VB.CommandButton Command1 
         Caption         =   "Novo Ítem"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11280
         TabIndex        =   2
         Top             =   600
         Width           =   1935
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   4335
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   7646
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
End
Attribute VB_Name = "frmConsultaItens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExcluir_Click()

If MsgBox("Deseja relamente excluir o Item ", vbQuestion + vbYesNo, "Exclusão de Ítens") = vbYes Then
    MsgBox "Item excluído com sucesso", vbInformation, "Exclusão de Itens"
    
End If

End Sub

Private Sub cmdFechar_Click()
Unload Me

End Sub
