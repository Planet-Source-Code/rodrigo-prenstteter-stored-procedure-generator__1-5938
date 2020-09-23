VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SP01wxCamposRelac 
   Caption         =   "Gerador de Stored Procedure "
   ClientHeight    =   4725
   ClientLeft      =   1935
   ClientTop       =   1590
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4725
   ScaleWidth      =   8865
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   60
      TabIndex        =   6
      Top             =   4020
      Width           =   8715
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   3420
         TabIndex        =   7
         Top             =   180
         Width           =   1455
      End
   End
   Begin VB.Frame frmAtributoConsulta 
      Caption         =   "Campos da tabela relacionada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   8715
      Begin VB.CommandButton cmdAtualiza 
         Caption         =   "OK"
         Height          =   315
         Left            =   8160
         TabIndex        =   10
         Top             =   3540
         Width           =   435
      End
      Begin VB.TextBox txtAlias 
         Height          =   315
         Left            =   1140
         TabIndex        =   9
         Top             =   3540
         Width           =   6975
      End
      Begin VB.CommandButton cmdCamposRelacAcima 
         Caption         =   "^"
         Height          =   675
         Left            =   8160
         TabIndex        =   4
         Top             =   240
         Width           =   460
      End
      Begin VB.CommandButton cmdCamposRelacAbaixo 
         Caption         =   "v"
         Height          =   675
         Left            =   8160
         TabIndex        =   3
         Top             =   960
         Width           =   460
      End
      Begin VB.CommandButton cmdCamposRelacMarcarTodos 
         Caption         =   "xxx"
         Height          =   675
         Left            =   8160
         TabIndex        =   2
         Top             =   1680
         Width           =   460
      End
      Begin VB.CommandButton cmdCamposRelacDesmarcarTodos 
         Caption         =   "---"
         Height          =   675
         Left            =   8160
         TabIndex        =   1
         Top             =   2400
         Width           =   460
      End
      Begin MSComctlLib.ListView lvwCamposRelac 
         Height          =   3255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nome Coluna"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblSelect 
         Caption         =   "Alias coluna"
         Height          =   315
         Left            =   180
         TabIndex        =   8
         Top             =   3600
         Width           =   1575
      End
   End
End
Attribute VB_Name = "SP01wxCamposRelac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    lvwCamposRelac.SetFocus
    If lvwCamposRelac.ListItems.Count = 0 Then
        Call SP01wxCamposRelac.MontaListViewColunaRelac(lvwCamposRelac, SP01wxRelacionamentos.moRSetCamposTabelaRelac)
    Else
        txtAlias.Text = lvwCamposRelac.ListItems.Item(1).SubItems(2)
    End If
End Sub

Private Sub cmdOK_Click()
    Dim siLoop          As Integer
    Dim siQtdMarcado    As Integer
    
    siQtdMarcado = 0
    For siLoop = 1 To lvwCamposRelac.ListItems.Count
        If lvwCamposRelac.ListItems.Item(siLoop).Checked Then
            siQtdMarcado = siQtdMarcado + 1
            If lvwCamposRelac.ListItems.Item(siLoop).SubItems(1) = lvwCamposRelac.ListItems.Item(siLoop).SubItems(2) Then
                MsgBox "Erro na coluna: " & siLoop & ". Nome da coluna igual ao alias", vbOKOnly + vbInformation, "Mensagem"
                Exit Sub
            End If
        End If
    Next siLoop
    If siQtdMarcado = 0 Then
        MsgBox "Não assinalado nenhum campo da tabela relacionada", vbOKOnly + vbInformation, "Mensagem"
        Exit Sub
    End If
    If SP01wxRelacionamentos.msOpcao = "Alt" Then
        For siLoop = 1 To lvwCamposRelac.ListItems.Count
            SP01mxMain.paCamposTabRelacChk(SP01wxStoredProced.lvwTabelaRelac.SelectedItem.Index - 1, siLoop - 1) = lvwCamposRelac.ListItems.Item("K" & Format(siLoop, "00000")).Checked
            SP01mxMain.paCamposTabRelacAli(SP01wxStoredProced.lvwTabelaRelac.SelectedItem.Index - 1, siLoop - 1) = lvwCamposRelac.ListItems.Item("K" & Format(siLoop, "00000")).SubItems(2)
        Next siLoop
    End If
    
    Me.Hide
End Sub

Private Sub cmdCamposRelacAcima_Click()
    Call SP01wxStoredProced.MovimentoLinha(enumAcima, lvwCamposRelac)
End Sub

Private Sub cmdCamposRelacAbaixo_Click()
    Call SP01wxStoredProced.MovimentoLinha(enumAbaixo, lvwCamposRelac)
End Sub

Private Sub cmdCamposRelacMarcarTodos_Click()
    Call SP01wxStoredProced.MarcaColunas(enumMarcar, lvwCamposRelac)
End Sub

Private Sub cmdCamposRelacDesmarcarTodos_Click()
    Call SP01wxStoredProced.MarcaColunas(enumDesmarcar, lvwCamposRelac)
End Sub

Public Function MontaListViewColunaRelac(ByRef polvwFonte As ListView, _
                                         ByRef poRSetCamposTabela As ADODB.Recordset) As Boolean
    Dim siContador          As Integer
        
    On Error GoTo ErroMontaListViewColunaRelac
    polvwFonte.ColumnHeaders.Clear
    polvwFonte.ColumnHeaders.Add 1, , "Posição", 800
    polvwFonte.ColumnHeaders.Add 2, , "Nome Coluna"
    polvwFonte.ColumnHeaders.Add 3, , "Alias Coluna", 3600
    polvwFonte.ColumnHeaders.Item(2).Width = polvwFonte.Width - _
                                             polvwFonte.ColumnHeaders.Item(1).Width - _
                                             polvwFonte.ColumnHeaders.Item(3).Width
    polvwFonte.View = lvwReport
    
    polvwFonte.ListItems.Clear
    If SP01wxRelacionamentos.msOpcao = "Inc" Then
        If poRSetCamposTabela.RecordCount > 0 Then poRSetCamposTabela.MoveFirst
        siContador = 1
        While Not poRSetCamposTabela.EOF
            polvwFonte.ListItems.Add , "K" & Format(poRSetCamposTabela!ORDINAL_POSITION, "00000"), siContador
            
            polvwFonte.ListItems.Item("K" & Format(poRSetCamposTabela!ORDINAL_POSITION, "00000")).SubItems(1) = poRSetCamposTabela!COLUMN_NAME
            polvwFonte.ListItems.Item("K" & Format(poRSetCamposTabela!ORDINAL_POSITION, "00000")).SubItems(2) = "T" & SP01mxMain.piProxTabela + 2 & "_" & poRSetCamposTabela!COLUMN_NAME
            
            siContador = siContador + 1
            
            poRSetCamposTabela.MoveNext
        Wend
        Call SP01wxStoredProced.MarcaColunas(enumMarcar, polvwFonte)
    Else
        For siContador = 1 To piQtdCpoTabRelac
            If SP01mxMain.paCamposTabRelac(SP01wxStoredProced.lvwTabelaRelac.SelectedItem.Index - 1, siContador - 1) <> "" Then
                polvwFonte.ListItems.Add , "K" & Format(siContador, "00000"), siContador
                
                polvwFonte.ListItems.Item("K" & Format(siContador, "00000")).SubItems(1) = SP01mxMain.paCamposTabRelac(SP01wxStoredProced.lvwTabelaRelac.SelectedItem.Index - 1, siContador - 1)
                polvwFonte.ListItems.Item("K" & Format(siContador, "00000")).SubItems(2) = SP01mxMain.paCamposTabRelacAli(SP01wxStoredProced.lvwTabelaRelac.SelectedItem.Index - 1, siContador - 1)
                
                polvwFonte.ListItems.Item("K" & Format(siContador, "00000")).Checked = SP01mxMain.paCamposTabRelacChk(SP01wxStoredProced.lvwTabelaRelac.SelectedItem.Index - 1, siContador - 1)
            End If
        Next siContador
    End If
    MontaListViewColunaRelac = True
    Exit Function
ErroMontaListViewColunaRelac:
    MontaListViewColunaRelac = False
End Function

Private Sub lvwCamposRelac_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtAlias.Text = Item.SubItems(2)
End Sub

Private Sub lvwCamposRelac_KeyUp(KeyCode As Integer, Shift As Integer)
    txtAlias.Text = lvwCamposRelac.SelectedItem.SubItems(2)
End Sub

Private Sub cmdAtualiza_Click()
    lvwCamposRelac.SelectedItem.SubItems(2) = txtAlias.Text
    lvwCamposRelac.SetFocus
End Sub
