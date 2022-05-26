VERSION 5.00
Begin VB.Form Tela_Grupos 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   1584
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2532
   LinkTopic       =   "Form1"
   ScaleHeight     =   1584
   ScaleWidth      =   2532
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox LT_Grupos 
      Height          =   1584
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2532
   End
End
Attribute VB_Name = "Tela_Grupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LT_Grupos_Click()
    Tela_Click.BE.Panels(1).Text = LT_Grupos.Text
    Tela_Click.BE.Panels(2).Text = ""
    Tela_Click.BE.Panels(3).Text = ""
End Sub
Private Sub LT_Grupos_DblClick()
    If KeyAscii = vbKeyReturn Then Tela_Grupos.Hide
End Sub
Private Sub LT_Grupos_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Tela_Grupos.Hide
End Sub
