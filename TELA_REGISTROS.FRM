VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Tela_Registros 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registros das fotos"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   Icon            =   "Tela_Registros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BT_Novo 
      Caption         =   "&Novo"
      Height          =   732
      Left            =   2880
      Picture         =   "Tela_Registros.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5040
      Width           =   732
   End
   Begin VB.CommandButton BT_Voltar 
      Caption         =   "&Voltar"
      Height          =   732
      Left            =   6720
      Picture         =   "Tela_Registros.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5040
      Width           =   732
   End
   Begin VB.CommandButton BT_Apagar 
      Caption         =   "&Apagar"
      Height          =   732
      Left            =   3960
      Picture         =   "Tela_Registros.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5040
      Width           =   732
   End
   Begin VB.OptionButton RB_Foto 
      Caption         =   "Foto"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   5160
      Width           =   735
   End
   Begin VB.OptionButton RB_Grupo 
      Caption         =   "Grupo"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   5520
      Width           =   735
   End
   Begin VB.OptionButton RB_Catalogo 
      Caption         =   "Catálogo"
      Height          =   195
      Left            =   1200
      TabIndex        =   15
      Top             =   5160
      Width           =   975
   End
   Begin VB.OptionButton RB_Modalidade 
      Caption         =   "Modalidade"
      Height          =   195
      Left            =   1200
      TabIndex        =   14
      Top             =   5520
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   360
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.Frame Frame1 
      Caption         =   "Modalidade"
      Height          =   4815
      Left            =   5160
      TabIndex        =   11
      Top             =   120
      Width           =   2295
      Begin VB.ListBox LT_Modalidades 
         Height          =   2400
         Left            =   120
         TabIndex        =   12
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Image IMG_2 
         BorderStyle     =   1  'Fixed Single
         Height          =   1815
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame FR_Grupos 
      Caption         =   "Grupos"
      Height          =   4815
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Width           =   2295
      Begin VB.CommandButton BT_Tira 
         Caption         =   "Apaga Foto   <<<"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton BT_Passa 
         Caption         =   "Passa Foto   >>>"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin VB.ListBox LT_Catalogos 
         Height          =   1425
         Left            =   120
         TabIndex        =   6
         Top             =   3240
         Width           =   2055
      End
      Begin VB.ListBox LT_Grupos 
         Height          =   1425
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Grupo:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Catálogo:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   3000
         Width           =   675
      End
   End
   Begin VB.Frame FR_Foto 
      Caption         =   "Fotos"
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.ListBox LT_Fotos 
         Height          =   1815
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   3
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox TXT_Arq 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   4440
         Width           =   2055
      End
      Begin RichTextLib.RichTextBox RTB_1 
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"Tela_Registros.frx":0E98
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Nome do arquivo:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   4200
         Width           =   1260
      End
      Begin VB.Image IMG_1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1815
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Tela_Registros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub BT_Apagar_Click()
    If RB_Foto.Value = True Then 'Apaga Foto
        If LT_Fotos.SelCount < 1 Then
            MsgResp = MsgBox("Você deve primeiro selecionar a(s) foto(s) que deseja  deletar na lista.", vbInformation + vbOKOnly, "Apagar foto")
            LT_Fotos.SetFocus
            Exit Sub
        Else
            Dim cFot As String
            cFot = ""
            For I = 0 To LT_Fotos.ListCount - 1
                If LT_Fotos.Selected(I) = True Then
                    If cFot = "" Then
                        cFot = LT_Fotos.List(I)
                    Else
                        cFot = cFot & " , " & LT_Fotos.List(I)
                    End If
                End If
            Next I
            MsgResp = MsgBox("Você tem certeza que deseja apagar a(s) foto(s): " & cFot & " ?", vbYesNo + vbDefaultButton2 + vbQuestion, "Apagar foto(s)")
            If MsgResp = vbNo Then Exit Sub
            'Apaga fotos
            AtivaTela (True)
            For I = 0 To LT_Fotos.ListCount - 1
                If LT_Fotos.Selected(I) = True Then
                    BDCLI_TBFOT.Seek "=", LT_Fotos.List(I)
                    If Not BDCLI_TBFOT.NoMatch Then BDCLI_TBFOT.Delete
                End If
            Next I
            For I = LT_Fotos.ListCount - 1 To 0 Step -1
                If LT_Fotos.Selected(I) = True Then LT_Fotos.RemoveItem (I)
            Next I
            AtivaTela (False)
        End If
    ElseIf RB_Grupo.Value = True Then 'Apaga Grupo
        If LT_Grupos.ListIndex = -1 Then
            MsgResp = MsgBox("Você deve primeiro selecionar o Grupo que deseja deletar na lista.", vbInformation + vbOKOnly, "Apagar Grupo")
            LT_Grupos.SetFocus
            Exit Sub
        Else
            MsgResp = MsgBox("Você tem certeza que deseja apagar o Grupo " & LT_Grupos.List(I) & " ? Esta operação poderá demorar, além de afetar provavelmente vários registros, pois irá apagar o Grupo, os Catálogos deste Grupo e as Modalidades de fotos registradas neste Grupo. Você tem certeza ?", vbYesNo + vbDefaultButton2 + vbQuestion, "Apagar Grupo")
            If MsgResp = vbNo Then Exit Sub
            'Apaga grupo
            AtivaTela (True)

            AtivaTela (False)
        End If
    ElseIf RB_Catalogo.Value = True Then 'Apaga Catalogo
    
    End If
End Sub
Private Sub BT_Novo_Click()
    If RB_Foto.Value = True Then 'Nova Foto
        CD.DialogTitle = "Indique o caminho das novas fotos"
        CD.Filter = "Arquivos de Fotos (*.jpg;*.gif;*.tif;*.bmp;*.pcx)|*.jpg;*.gif;*.tif;*.bmp;*.pcx"
        CD.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames
        CD.ShowOpen
        'Inclui a foto
        If CD.filename <> "" Then
            Tela_Registros.Enabled = False
            AtivaTela (True)
            RTB_1.LoadFile (CD.filename), 1
            TXT_Arq.Text = CD.filename
            If BDCLI_TBFOT.RecordCount = 0 Then
                BDCLI_TBFOT.AddNew
                BDCLI_TBFOT_CPIND.Value = 1
                BDCLI_TBFOT_CPNOM.Value = TXT_Arq.Text
                BDCLI_TBFOT_CPFOT.Value = RTB_1.Text
                BDCLI_TBFOT.Update
                LT_Fotos.AddItem ("1")
            Else
                Dim nNum As Integer
                BDCLI_TBFOT.MoveLast
                nNum = BDCLI_TBFOT_CPIND.Value + 1
                BDCLI_TBFOT.AddNew
                BDCLI_TBFOT_CPIND.Value = nNum
                BDCLI_TBFOT_CPNOM.Value = TXT_Arq.Text
                BDCLI_TBFOT_CPFOT.Value = RTB_1.Text
                BDCLI_TBFOT.Update
                LT_Fotos.AddItem (nNum)
            End If
            RTB_1.Text = ""
            TXT_Arq.Text = ""
            Tela_Registros.Enabled = True
            AtivaTela (False)
        End If
    ElseIf RB_Grupo.Value = True Then 'Novo Grupo
        cVal = InputBox("Se você deseja incluir um novo grupo, digite o nome dele abaixo:", "Incluir Grupos", "")
        If cVal = "" Then
            Exit Sub
        Else
            For I = 0 To LT_Grupos.ListCount - 1
                If LT_Grupos.Text = cVal Then
                    MsgResp = MsgBox("O nome do grupo digitado já existe - tente outro nome.", vbInformation + vbOKOnly, "Grupo já existe")
                    Exit Sub
                End If
            Next I
            BDCLI_TBGRU.AddNew
            BDCLI_TBGRU_CPGRU.Value = cVal
            BDCLI_TBGRU.Update
            LT_Grupos.AddItem (cVal)
            LT_Catalogos.Clear
            LT_Modalidades.Clear
        End If
        LT_Grupos.ListIndex = -1
    ElseIf RB_Catalogo.Value = True Then 'Novo Catalogo
        If LT_Grupos.ListIndex = -1 Then
            MsgResp = MsgBox("Você deve primeiro selecionar um grupo deste catálogo.", vbInformation + vbOKOnly, "Informar Grupo")
            LT_Grupos.SetFocus
            Exit Sub
        End If
        cVal = InputBox("Se você deseja incluir um novo catálogo do grupo " & Trim(LT_Grupos.Text) & ", digite o nome dele abaixo:", "Incluir Catálogo", "")
        If cVal = "" Then
            Exit Sub
        Else
            For I = 0 To LT_Catalogos.ListCount - 1
                If LT_Catalogos.Text = cVal Then
                    MsgResp = MsgBox("O nome do catálogo do grupo " & Trim(LT_Grupos.Text) & " digitado já existe - tente outro nome.", vbInformation + vbOKOnly, "Catálogo já existe")
                    Exit Sub
                End If
            Next I
            BDCLI_TBCAT.AddNew
            BDCLI_TBCAT_CPGRU.Value = LT_Grupos.Text
            BDCLI_TBCAT_CPCAT.Value = cVal
            BDCLI_TBCAT.Update
            LT_Catalogos.AddItem (cVal)
            LT_Modalidades.Clear
        End If
        LT_Catalogos.ListIndex = -1
    Else
        LT_Grupos.ListIndex = -1
        LT_Catalogos.ListIndex = -1
        LT_Modalidades.Clear
    End If
    cVal = ""
End Sub
Private Sub BT_Voltar_Click()
    Unload Tela_Registros
End Sub
Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    'Carrega lista de fotos
    LT_Fotos.Clear
    If BDCLI_TBFOT.RecordCount > 0 Then
        BDCLI_TBFOT.MoveFirst
        Do While Not BDCLI_TBFOT.EOF
            LT_Fotos.AddItem (BDCLI_TBFOT_CPIND.Value)
            BDCLI_TBFOT.MoveNext
        Loop
    End If
    LT_Fotos.ListIndex = -1
    'Carrega lista de grupos
    LT_Grupos.Clear
    If BDCLI_TBGRU.RecordCount > 0 Then
        BDCLI_TBGRU.MoveFirst
        Do While Not BDCLI_TBGRU.EOF
            LT_Grupos.AddItem (BDCLI_TBGRU_CPGRU.Value)
            BDCLI_TBGRU.MoveNext
        Loop
    End If
    LT_Grupos.ListIndex = -1
    'apaga lista de catalogos
    LT_Catalogos.Clear

    
    RB_Foto.Value = True
    
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Terminate()
    Unload Tela_Registros
End Sub
Private Sub LT_Fotos_Click()
    If LT_Fotos.ListIndex = -1 Then Exit Sub
    'Procura a foto
    AtivaTela (True)
    BDCLI_TBFOT.Seek "=", LT_Fotos.Text
    If BDCLI_TBFOT.NoMatch Then
        MsgBox ("Ocorreu algum erro durante a procura da foto...")
        AtivaTela (False)
        Exit Sub
    End If
    RTB_1.Text = BDCLI_TBFOT_CPFOT.Value
    RTB_1.SaveFile Diretorio & "\Click.jpg", 1
    IMG_1.Picture = LoadPicture(Diretorio & "\Click.jpg")
    Kill (Diretorio & "\Click.jpg")
    RTB_1.Text = ""
    AtivaTela (False)
End Sub
Private Sub LT_Grupos_Click()
    If LT_Grupos.ListIndex = -1 Then Exit Sub
    AtivaTela (True)
    'Limpa listas
    LT_Catalogos.Clear
    LT_Modalidades.Clear
    'Procura catalogos
    If BDCLI_TBCAT.RecordCount > 0 Then
        BDCLI_TBCAT.MoveFirst
        Do While Not BDCLI_TBCAT.EOF
            If BDCLI_TBCAT_CPGRU.Value = LT_Grupos.Text Then LT_Catalogos.AddItem (BDCLI_TBCAT_CPCAT.Value)
            BDCLI_TBCAT.MoveNext
        Loop
    End If
    
    AtivaTela (False)
End Sub
Private Sub RB_Catalogo_Click()
    BT_Novo.Enabled = True
    BT_Apagar.Enabled = True
    BT_Voltar.Enabled = True
End Sub
Private Sub RB_Foto_Click()
    BT_Novo.Enabled = True
    BT_Apagar.Enabled = True
    BT_Voltar.Enabled = True
End Sub
Private Sub RB_Grupo_Click()
    BT_Novo.Enabled = True
    BT_Apagar.Enabled = True
    BT_Voltar.Enabled = True
End Sub
Private Sub RB_Modalidade_Click()
    BT_Novo.Enabled = True
    BT_Apagar.Enabled = False
    BT_Voltar.Enabled = True
End Sub
