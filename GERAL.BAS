Attribute VB_Name = "Geral"
'Vari�veis Diversas
Public TC As Click.Tela_Click
Public MsgResp
Public I As Integer
Public ModoEdicao As Boolean
Public X1, X2, Y1, Y2 As Long
Public Diretorio As String
Public cVal As String

' ***************   �REA DE TRABALHO   ****************
Global WKSP As Workspace

' ***************   BANCO DE DADOS    ****************
Global BDCLI As Database 'Click

' ******************   TABELAS    ******************
Global BDCLI_TBFOT As Recordset 'Tabela de Fotos
Global BDCLI_TBMDL As Recordset 'Tabela de Modalidades - Click
Global BDCLI_TBGRU As Recordset 'Tabela de Grupos
Global BDCLI_TBCAT As Recordset 'Tabela de Cat�logos

' ******************   CAMPOS    *******************
' Campos da Tabela Fotos
Global BDCLI_TBFOT_CPIND As Field '�ndice
Global BDCLI_TBFOT_CPNOM As Field 'Nome
Global BDCLI_TBFOT_CPFOT As Field 'Foto

' Campos da Tabela Modalidades
Global BDCLI_TBMDL_CPGRU As Field 'Grupo
Global BDCLI_TBMDL_CPCAT As Field 'Cat�logo
Global BDCLI_TBMDL_CPIND As Field '�ndice da Foto

' Campos da Tabela Grupos
Global BDCLI_TBGRU_CPGRU As Field 'Grupo
Global BDCLI_TBGRU_CPDES As Field 'Descri��o

' Campos da Tabela Cat�logos
Global BDCLI_TBCAT_CPGRU As Field 'Grupo
Global BDCLI_TBCAT_CPCAT As Field 'Cat�logo
Global BDCLI_TBCAT_CPDES As Field 'Descri��o


' ***************   FUNCOES PUBLICAS   ****************

Public Static Sub AtivaTela(Ponteiro_Ampulheta As Boolean)
    If Ponteiro_Ampulheta = True Then
        Screen.MousePointer = vbHourglass
    Else
        Screen.MousePointer = vbDefault
    End If
End Sub

    
