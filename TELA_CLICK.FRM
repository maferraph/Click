VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Tela_Click 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Click - Fotos Incríveis !"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   7710
   Icon            =   "Tela_Click.frx":0000
   ScaleHeight     =   3855
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar BF 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   847
      ButtonWidth     =   572
      ButtonHeight    =   550
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "LI"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   21
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Fotos"
            Description     =   "Registro das fotos do sistema"
            Object.ToolTipText     =   "Registro das fotos do sistema"
            Object.Tag             =   "Fotos"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Grupos"
            Description     =   "Grupos de fotos"
            Object.ToolTipText     =   "Grupos de fotos"
            Object.Tag             =   "Grupos"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Catalogos"
            Object.ToolTipText     =   "Catálogos de fotos"
            Object.Tag             =   "Catalogos"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "S1"
            Object.Tag             =   "S1"
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "BVT"
            Description     =   "Vai para primeira foto"
            Object.ToolTipText     =   "Vai para primeira foto"
            Object.Tag             =   "BVT"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "BV1"
            Description     =   "Vai para foto anterior"
            Object.ToolTipText     =   "Vai para foto anterior"
            Object.Tag             =   "BV1"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "BA1"
            Description     =   "Vai para próxima foto"
            Object.ToolTipText     =   "Vai para próxima foto"
            Object.Tag             =   "BA1"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "BAT"
            Description     =   "Vai para última foto"
            Object.ToolTipText     =   "Vai para última foto"
            Object.Tag             =   "BAT"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "S2"
            Object.Tag             =   "S2"
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "1_IMG"
            Description     =   "Exibe uma imagem"
            Object.ToolTipText     =   "Exibe uma imagem"
            Object.Tag             =   "1_IMG"
            ImageIndex      =   17
            Style           =   2
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "IMG_FIC"
            Description     =   "Exibe uma imagem com ficha"
            Object.ToolTipText     =   "Exibe uma imagem com ficha"
            Object.Tag             =   "IMG_FIC"
            ImageIndex      =   18
            Style           =   2
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "4_IMG"
            Description     =   "Exibe quatro imagens"
            Object.ToolTipText     =   "Exibe quatro imagens"
            Object.Tag             =   "4_IMG"
            ImageIndex      =   19
            Style           =   2
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "IMG_TC"
            Description     =   "Exibe a imagem em tela cheia"
            Object.ToolTipText     =   "Exibe a imagem em tela cheia"
            Object.Tag             =   "IMG_TC"
            ImageIndex      =   20
            Style           =   2
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "S3"
            Object.Tag             =   "S3"
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Sequencia"
            Description     =   "Inicia sequência"
            Object.ToolTipText     =   "Inicia sequência"
            Object.Tag             =   "Sequencia"
            ImageIndex      =   26
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "S4"
            Object.Tag             =   "S4"
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Config"
            Description     =   "Configurações do Click"
            Object.ToolTipText     =   "Configurações do Click"
            Object.Tag             =   "Config"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "S5"
            Object.Tag             =   "S5"
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Imprimir"
            Description     =   "Imprime a foto"
            Object.ToolTipText     =   "Imprime a foto"
            Object.Tag             =   "Imprimir"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "S6"
            Object.Tag             =   "S6"
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button21 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Sair"
            Description     =   "Encerrar o Click"
            Object.ToolTipText     =   "Encerrar o Click"
            Object.Tag             =   "Sair"
            ImageIndex      =   24
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame FR_UmaFoto 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1572
      Left            =   960
      TabIndex        =   4
      Top             =   480
      Width           =   1692
      Begin VB.Image IMG_UmaFoto 
         Height          =   1092
         Left            =   240
         Top             =   240
         Width           =   1092
      End
   End
   Begin VB.ListBox LT_IMG 
      Height          =   255
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   372
   End
   Begin ComctlLib.StatusBar BE 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3480
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            Key             =   "Grupos"
            Object.Tag             =   "Grupos"
            Object.ToolTipText     =   "Grupos de fotos"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            Key             =   "Catalogo"
            Object.Tag             =   "Catalogo"
            Object.ToolTipText     =   "Catálogo de fotos"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            Key             =   "NumFotos"
            Object.Tag             =   "NumFotos"
            Object.ToolTipText     =   "Número de fotos"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   1200
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   327681
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   492
      Left            =   0
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   612
      _ExtentX        =   1085
      _ExtentY        =   873
      _Version        =   393217
      TextRTF         =   $"Tela_Click.frx":030A
   End
   Begin ComctlLib.ImageList LI 
      Left            =   0
      Top             =   600
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   29
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":03DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":06F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":0A0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":0D29
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":1043
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":135D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":1677
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":1991
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":1CAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":1FC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":22DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":25F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":2913
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":2C2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":2F47
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":3261
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":357B
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":3895
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":3BAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":3EC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":41E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":44FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":4817
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":4B31
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":4E4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":5165
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":547F
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":5799
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Tela_Click.frx":5AB3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu Menu_Principal 
      Caption         =   "Principal"
   End
   Begin VB.Menu Menu_Edicao 
      Caption         =   "Edição"
      Begin VB.Menu Menu_Edicao_Registros 
         Caption         =   "Registros"
      End
   End
   Begin VB.Menu Menu_Exibir 
      Caption         =   "Exibir"
   End
   Begin VB.Menu Menu_Ajuda 
      Caption         =   "Ajuda"
   End
End
Attribute VB_Name = "Tela_Click"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub BF_ButtonClick(ByVal Button As ComctlLib.Button)
    If Button.Key = "Sair" Then
        Fim
    ElseIf Button.Key = "Grupos" Then
        Tela_Grupos.Top = Button.Top + Button.Height + 700
        Tela_Grupos.Left = Button.Left
        Tela_Grupos.LT_Grupos.Top = 0
        Tela_Grupos.LT_Grupos.Left = 0
        Tela_Grupos.LT_Grupos.Visible = True
        Tela_Grupos.Show vbModal
    End If
End Sub
Private Sub Form_Load()
    Diretorio = "C:\Meus Documentos\Programas\Click"
    
    'Abrindo Workspaces
    Set WKSP = DBEngine.Workspaces(0)
    'Abrindo Banco de Dados
    AbreBD
    ' Abre Tabelas
    Set BDCLI_TBFOT = BDCLI.OpenRecordset("Fotos")
    Set BDCLI_TBMDL = BDCLI.OpenRecordset("Modalidades")
    Set BDCLI_TBGRU = BDCLI.OpenRecordset("Grupos")
    Set BDCLI_TBCAT = BDCLI.OpenRecordset("Catálogos")
    ' Abre Campos da Tabela Fotos
    Set BDCLI_TBFOT_CPIND = BDCLI_TBFOT.Fields("Índice")
    Set BDCLI_TBFOT_CPNOM = BDCLI_TBFOT.Fields("Nome")
    Set BDCLI_TBFOT_CPFOT = BDCLI_TBFOT.Fields("Foto")
    ' Abre Campos da Tabela Modalidades
    Set BDCLI_TBMDL_CPGRU = BDCLI_TBMDL.Fields("Grupo")
    Set BDCLI_TBMDL_CPCAT = BDCLI_TBMDL.Fields("Catálogo")
    Set BDCLI_TBMDL_CPIND = BDCLI_TBMDL.Fields("Índice da Foto")
    ' Abre Campos da Tabela Grupos
    Set BDCLI_TBGRU_CPGRU = BDCLI_TBGRU.Fields("Grupo")
    Set BDCLI_TBGRU_CPDES = BDCLI_TBGRU.Fields("Descrição")
    ' Abre Campos da Tabela Catálogos
    Set BDCLI_TBCAT_CPGRU = BDCLI_TBCAT.Fields("Grupo")
    Set BDCLI_TBCAT_CPCAT = BDCLI_TBCAT.Fields("Catálogo")
    Set BDCLI_TBCAT_CPDES = BDCLI_TBCAT.Fields("Descrição")
    'Abre índices
    BDCLI_TBFOT.Index = "Índice"
    BDCLI_TBMDL.Index = "Foto"
    BDCLI_TBGRU.Index = "Grupos"
    BDCLI_TBCAT.Index = "Catálogo"
    
    'Abre frame 1 foto
    FR_UmaFoto.Top = 0
    FR_UmaFoto.Left = 0
    FR_UmaFoto.Height = Tela_Click.Height
    FR_UmaFoto.Width = Tela_Click.Width
    FR_UmaFoto.Visible = True
    LT_IMG.Clear
    LT_IMG.Visible = False
    RTB.Visible = False
    
    
    BE.Panels(1).Text = "Grupos"
    BE.Panels(1).Picture = LI.ListImages(22).Picture
    BE.Panels(2).Text = "Catálogos"
    BE.Panels(2).Picture = LI.ListImages(10).Picture
    BE.Panels(3).Text = "1 de 100"
    BE.Panels(3).Picture = LI.ListImages(16).Picture



End Sub




'************************************************
'             FUNCOES DESTE MÓDULO
'************************************************
Private Sub Fim()
    BDCLI_TBFOT.Close
    BDCLI_TBMDL.Close
    BDCLI_TBGRU.Close
    BDCLI_TBCAT.Close
    BDCLI.Close
    End
End Sub
Private Sub AbreBD()
On Error GoTo ERRO_BD
    Set BDCLI = WKSP.OpenDatabase(Diretorio & "\Click.mdb")
    Exit Sub
ERRO_BD:
    CD.DialogTitle = "Indique o caminho do arquivo Click.mdb"
    CD.filename = "Click.mdb"
    CD.Filter = "Banco de Dados do Sistema Click|Click.mdb"
    CD.ShowOpen
On Error GoTo ERRO_BD2
    Set BDCLI = WKSP.OpenDatabase(CD.filename)
    Exit Sub
ERRO_BD2:
    MsgBox ("Infelizmente não foi possível abrir o arquivo: " & CD.filename & ", do qual faz parte do Sistema Click.")
    End
End Sub
Private Sub Form_Terminate()
    Fim
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Fim
End Sub
Private Sub LT_IMG_Click()
    Tela_Click.Enabled = False
    Tela_Click.MousePointer = vbHourglass
    Dim X, Y, W, A As Double
    'Frame 1 Foto
    If FR_UmaFoto.Visible = True Then
        BDCLI_TBFOT.Seek "=", LT_IMG.Text
        If BDCLI_TBFOT.NoMatch Then
            MsgBox ("Ocorreu algum erro durante a procura da foto...")
            Exit Sub
        End If
        RTB.Text = BDCLI_TBFOT_CPFOT.Value
        RTB.SaveFile Diretorio & "\Click.jpg", 1
        'If RB_Prop.Value = True Then
            IMG_UmaFoto.Visible = False
            IMG_UmaFoto.Stretch = False
            IMG_UmaFoto.Top = BF.Height + 100
            IMG_UmaFoto.Picture = LoadPicture(Diretorio & "\Click.jpg")
            If IMG_UmaFoto.Height > ((BF2.Top - 100) - (BF.Height + 100)) Then
                Y = IMG_UmaFoto.Height
                W = (BF2.Top - 100) - (BF.Height + 100)
                A = W / Y
                IMG_UmaFoto.Width = IMG_UmaFoto.Width * A
                IMG_UmaFoto.Height = IMG_UmaFoto.Height * A
                IMG_UmaFoto.Left = Tela_Click.Width / 2 - IMG_UmaFoto.Width / 2
            End If
            If IMG_UmaFoto.Width > Tela_Click.Width Then
                X = IMG_UmaFoto.Width
                W = Tela_Click.Width - 200
                A = W / X
                IMG_UmaFoto.Left = 100
                IMG_UmaFoto.Width = IMG_UmaFoto.Width * A
                IMG_UmaFoto.Height = IMG_UmaFoto.Height * A
            End If
            IMG_UmaFoto.Stretch = True
            IMG_UmaFoto.Visible = True
        'End If
    'Frame 4 fotos
    'ElseIf FR_QuatroFotos.Visible = True Then
    End If
    Tela_Click.Enabled = True
    Tela_Click.MousePointer = vbDefault
End Sub
Private Sub Menu_Edicao_Registros_Click()
    Tela_Registros.Show vbModal
End Sub
