VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form SP01wxRelacionamentos 
   Caption         =   "Gerador de Stored Procedure "
   ClientHeight    =   5415
   ClientLeft      =   1560
   ClientTop       =   2010
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   60
      TabIndex        =   30
      Top             =   4680
      Width           =   7935
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Ok"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   31
         Top             =   180
         Width           =   1455
      End
   End
   Begin VB.Frame frmTipoRelac 
      Caption         =   "Tipo relacionamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   60
      TabIndex        =   25
      Top             =   3960
      Width           =   7935
      Begin VB.OptionButton optFullJoin 
         Caption         =   "Full Outer Join"
         Enabled         =   0   'False
         Height          =   255
         Left            =   6420
         TabIndex        =   29
         Top             =   300
         Width           =   1335
      End
      Begin VB.OptionButton optRightJoin 
         Caption         =   "Right Outer Join"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4140
         TabIndex        =   28
         Top             =   300
         Width           =   1455
      End
      Begin VB.OptionButton optLeftJoin 
         Caption         =   "Left Outer Join"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1980
         TabIndex        =   27
         Top             =   300
         Width           =   1335
      End
      Begin VB.OptionButton optInnerJoin 
         Caption         =   "Inner Join"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   300
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.Frame frmColunas 
      Caption         =   "Colunas relacionadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   60
      TabIndex        =   3
      Top             =   780
      Width           =   7935
      Begin VB.CommandButton cmdCancelaRelac 
         Caption         =   "x"
         Enabled         =   0   'False
         Height          =   315
         Index           =   6
         Left            =   7380
         TabIndex        =   38
         ToolTipText     =   "Cancela relacionamento"
         Top             =   2700
         Width           =   435
      End
      Begin VB.CommandButton cmdCancelaRelac 
         Caption         =   "x"
         Enabled         =   0   'False
         Height          =   315
         Index           =   5
         Left            =   7380
         TabIndex        =   37
         ToolTipText     =   "Cancela relacionamento"
         Top             =   2340
         Width           =   435
      End
      Begin VB.CommandButton cmdCancelaRelac 
         Caption         =   "x"
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   7380
         TabIndex        =   36
         ToolTipText     =   "Cancela relacionamento"
         Top             =   1980
         Width           =   435
      End
      Begin VB.CommandButton cmdCancelaRelac 
         Caption         =   "x"
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   7380
         TabIndex        =   35
         ToolTipText     =   "Cancela relacionamento"
         Top             =   1620
         Width           =   435
      End
      Begin VB.CommandButton cmdCancelaRelac 
         Caption         =   "x"
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   7380
         TabIndex        =   34
         ToolTipText     =   "Cancela relacionamento"
         Top             =   1260
         Width           =   435
      End
      Begin VB.CommandButton cmdCancelaRelac 
         Caption         =   "x"
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   7380
         TabIndex        =   33
         ToolTipText     =   "Cancela relacionamento"
         Top             =   900
         Width           =   435
      End
      Begin VB.CommandButton cmdCancelaRelac 
         Caption         =   "x"
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   7380
         TabIndex        =   32
         ToolTipText     =   "Cancela relacionamento"
         Top             =   540
         Width           =   435
      End
      Begin MSDataListLib.DataCombo cboColunaFonte 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   540
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cboColunaRelac 
         Height          =   315
         Index           =   0
         Left            =   3960
         TabIndex        =   6
         Top             =   540
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cboColunaFonte 
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   900
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cboColunaRelac 
         Height          =   315
         Index           =   1
         Left            =   3960
         TabIndex        =   9
         Top             =   900
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cboColunaFonte 
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   1260
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cboColunaRelac 
         Height          =   315
         Index           =   2
         Left            =   3960
         TabIndex        =   12
         Top             =   1260
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cboColunaFonte 
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   1620
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cboColunaRelac 
         Height          =   315
         Index           =   3
         Left            =   3960
         TabIndex        =   15
         Top             =   1620
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cboColunaFonte 
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   1980
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cboColunaRelac 
         Height          =   315
         Index           =   4
         Left            =   3960
         TabIndex        =   18
         Top             =   1980
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cboColunaFonte 
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   2340
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cboColunaRelac 
         Height          =   315
         Index           =   5
         Left            =   3960
         TabIndex        =   21
         Top             =   2340
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cboColunaFonte 
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   22
         Top             =   2700
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cboColunaRelac 
         Height          =   315
         Index           =   6
         Left            =   3960
         TabIndex        =   24
         Top             =   2700
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblRelac 
         Alignment       =   2  'Center
         Caption         =   "Campos tabela relacionada"
         Height          =   255
         Left            =   3960
         TabIndex        =   40
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label lblFonte 
         Alignment       =   2  'Center
         Caption         =   "Campos tabela principal"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "com"
         Height          =   315
         Left            =   3420
         TabIndex        =   23
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "com"
         Height          =   315
         Left            =   3420
         TabIndex        =   20
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "com"
         Height          =   315
         Left            =   3420
         TabIndex        =   17
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "com"
         Height          =   315
         Left            =   3420
         TabIndex        =   14
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "com"
         Height          =   315
         Left            =   3420
         TabIndex        =   11
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "com"
         Height          =   315
         Left            =   3420
         TabIndex        =   8
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblCom1 
         Alignment       =   2  'Center
         Caption         =   "com"
         Height          =   315
         Left            =   3420
         TabIndex        =   5
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame frmDestino 
      Caption         =   "Relacionamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7935
      Begin VB.CommandButton cmdCamposRelac 
         Caption         =   "&Campos relacionados"
         Enabled         =   0   'False
         Height          =   315
         Left            =   6060
         TabIndex        =   41
         Top             =   240
         Width           =   1755
      End
      Begin MSDataListLib.DataCombo cboTabelaRelac 
         Height          =   315
         Left            =   1620
         TabIndex        =   2
         Top             =   240
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblTabelaRelac 
         Caption         =   "Tabela relacionada"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1515
      End
   End
End
Attribute VB_Name = "SP01wxRelacionamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public moRSetCamposTabelaRelac      As ADODB.Recordset
Public msOpcao                      As String
Public mbFormAtivado                As Boolean

Private Sub Form_Load()
    Dim siLoop          As Integer
    
    With cboTabelaRelac
        .BoundColumn = "tiponome"
        .ListField = "name"
        Set .RowSource = SP01wxStoredProced.moRSetTabela
        .Text = ""
    End With
    For siLoop = 0 To 6
        Call RelacionaCombo(cboColunaFonte(siLoop), SP01wxStoredProced.moRSetCamposTabela)
    Next siLoop
    
    Load SP01wxCamposRelac
End Sub

Private Sub Form_Activate()
    Dim siLoop          As Integer
    
    mbFormAtivado = False
    
    '------ habilitando frames
    frmDestino.Enabled = True
    frmColunas.Enabled = True
    frmTipoRelac.Enabled = True
    
    If msOpcao <> "Inc" Then
        '------ mostrando nome da tabela
        cboTabelaRelac.Text = paNomeTabelas(SP01wxStoredProced.lvwTabelaRelac.SelectedItem.Index - 1)
        '------ mostrando tipo de relacionamento
        optInnerJoin.Value = (paTipoRelac(SP01wxStoredProced.lvwTabelaRelac.SelectedItem.Index - 1) = enumInnerJoin)
        optLeftJoin.Value = (paTipoRelac(SP01wxStoredProced.lvwTabelaRelac.SelectedItem.Index - 1) = enumLeftJoin)
        optRightJoin.Value = (paTipoRelac(SP01wxStoredProced.lvwTabelaRelac.SelectedItem.Index - 1) = enumRightJoin)
        optFullJoin.Value = (paTipoRelac(SP01wxStoredProced.lvwTabelaRelac.SelectedItem.Index - 1) = enumFullJoin)
        '------ mostrando campos relacionados
        For siLoop = 0 To piQtdCpoRelac
            If paCamposFonte(SP01wxStoredProced.lvwTabelaRelac.SelectedItem.Index - 1, siLoop) <> "" Then
                cboColunaFonte(siLoop).Text = paCamposFonte(SP01wxStoredProced.lvwTabelaRelac.SelectedItem.Index - 1, siLoop)
                cboColunaRelac(siLoop).Text = paCamposRelac(SP01wxStoredProced.lvwTabelaRelac.SelectedItem.Index - 1, siLoop)
            End If
        Next siLoop
        '------ mostrando campos relacionados da tabela
        Call SP01wxCamposRelac.MontaListViewColunaRelac(SP01wxCamposRelac.lvwCamposRelac, moRSetCamposTabelaRelac)
        
        If msOpcao = "Exc" Then
            '------ desabilitando frames
            frmDestino.Enabled = False
            frmColunas.Enabled = False
            frmTipoRelac.Enabled = False
        End If
    End If
    mbFormAtivado = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload SP01wxCamposRelac
End Sub

Private Sub cboTabelaRelac_Change()
    Dim siLoop          As Integer
    
    Set moRSetCamposTabelaRelac = New ADODB.Recordset
    Set moRSetCamposTabelaRelac = SP01wxStoredProced.moConexao.Execute("sp_columns '" & Mid(cboTabelaRelac.BoundText, InStr(1, cboTabelaRelac.BoundText, "/") + 1, Len(cboTabelaRelac.BoundText)) & "'")

    For siLoop = 0 To 6
        Call RelacionaCombo(cboColunaRelac(siLoop), moRSetCamposTabelaRelac)
        cboColunaFonte(siLoop).Text = ""
        cboColunaFonte(siLoop).Enabled = (cboTabelaRelac.Text <> "")
        cboColunaRelac(siLoop).Text = ""
        cboColunaRelac(siLoop).Enabled = (cboTabelaRelac.Text <> "")
        cmdCancelaRelac(siLoop).Enabled = (cboTabelaRelac.Text <> "")
    Next siLoop
    cmdCamposRelac.Enabled = (cboTabelaRelac.Text <> "")
    optInnerJoin.Enabled = (cboTabelaRelac.Text <> "")
    optLeftJoin.Enabled = (cboTabelaRelac.Text <> "")
    optRightJoin.Enabled = (cboTabelaRelac.Text <> "")
    optFullJoin.Enabled = (cboTabelaRelac.Text <> "")
    optInnerJoin.Value = True
    cmdOK.Enabled = (cboTabelaRelac.Text <> "")
    If msOpcao <> "Inc" And mbFormAtivado Then Call AtualizaArray
    Call SP01wxCamposRelac.MontaListViewColunaRelac(SP01wxCamposRelac.lvwCamposRelac, moRSetCamposTabelaRelac)
End Sub

Private Sub cmdCamposRelac_Click()
    SP01wxCamposRelac.Show vbModal
End Sub

Private Sub cmdCancelaRelac_Click(Index As Integer)
    Dim siLoop          As Integer
    
    For siLoop = Index To 6
        cboColunaFonte(siLoop).Text = ""
        cboColunaRelac(siLoop).Text = ""
    Next siLoop
End Sub

Private Sub cmdOK_Click()
    Dim siLoop          As Integer
    Dim siLoop1         As Integer
    
    If msOpcao <> "Exc" Then
        If cboColunaFonte(0).Text = "" Or cboColunaRelac(0).Text = "" Then
            MsgBox "Definir corretamente os campos de relacionamento", vbOKOnly + vbInformation, "Mensagem"
            Exit Sub
        End If
        Call AtualizaArray
    Else
        If MsgBox("Confirma exclusão deste relacionamento?", vbYesNo + vbQuestion, "Deleção") = vbYes Then
            '------ movimenta próximos elementos para cima
            For siLoop = SP01wxStoredProced.lvwTabelaRelac.SelectedItem.Index To piQtdTabelas
                '------ movimentando nome da tabela
                paNomeTabelas(siLoop - 1) = paNomeTabelas(siLoop)
                '------ movimentando tipo de relacionamento
                paTipoRelac(siLoop - 1) = paTipoRelac(siLoop)
                '------ armazenando campos relacionados
                For siLoop1 = 0 To 6
                    paCamposFonte(siLoop - 1, siLoop1) = paCamposFonte(siLoop, siLoop1)
                    paCamposRelac(siLoop - 1, siLoop1) = paCamposRelac(siLoop, siLoop1)
                Next siLoop1
                '------ armazenando campos relacionados da tabela e atributos
                For siLoop1 = 0 To piQtdCpoTabRelac
                    '------ armazenando campos relacionados da tabela
                    paCamposTabRelac(siLoop - 1, siLoop1) = paCamposTabRelac(siLoop, siLoop1)
                    '------ armazenando campos relacionados da tabela (atributos)
                    paCamposTabRelacChk(siLoop - 1, siLoop1) = paCamposTabRelacChk(siLoop, siLoop1)
                    paCamposTabRelacAli(siLoop - 1, siLoop1) = paCamposTabRelacAli(siLoop, siLoop1)
                Next siLoop1
            Next siLoop
            '------ deletando último elemento
            paNomeTabelas(piQtdTabelas) = ""
            '------ movimentando tipo de relacionamento
            paTipoRelac(piQtdTabelas) = enumInnerJoin
            '------ armazenando campos relacionados
            For siLoop1 = 0 To 6
                paCamposFonte(piQtdTabelas, siLoop1) = ""
                paCamposRelac(piQtdTabelas, siLoop1) = ""
            Next siLoop1
            '------ armazenando campos relacionados da tabela e atributos
            For siLoop1 = 0 To piQtdCpoTabRelac
                '------ armazenando campos relacionados da tabela
                paCamposTabRelac(piQtdTabelas, siLoop1) = ""
                '------ armazenando campos relacionados da tabela (atributos)
                paCamposTabRelacChk(piQtdTabelas, siLoop1) = False
                paCamposTabRelacAli(piQtdTabelas, siLoop1) = ""
            Next siLoop1
            '------ corrige apontadores
            piProxTabela = piProxTabela - 1
            piProxCpoRelac(piProxTabela) = 0
            piProxCpoTabRelac(piProxTabela) = 0
        End If
    End If
    '-------- atualizando listview com relacionamentos
    SP01wxStoredProced.lvwTabelaRelac.ListItems.Clear
    For siLoop = 0 To piQtdTabelas
        If SP01mxMain.paNomeTabelas(siLoop) <> "" Then
            SP01wxStoredProced.lvwTabelaRelac.ListItems.Add , "T" & Format(siLoop, "00000"), siLoop + 1
            SP01wxStoredProced.lvwTabelaRelac.ListItems.Item("T" & Format(siLoop, "00000")).SubItems(1) = SP01mxMain.paNomeTabelas(siLoop)
        End If
    Next siLoop
    Unload Me
    SP01wxStoredProced.lvwTabelaRelac.Refresh
    SP01wxStoredProced.lvwTabelaRelac.SetFocus
End Sub

Private Sub RelacionaCombo(ByRef poCombo As MSDataListLib.DataCombo, _
                           ByRef poRSet As ADODB.Recordset)
    With poCombo
        .BoundColumn = "COLUMN_NAME"
        .ListField = "COLUMN_NAME"
        Set .RowSource = poRSet
        .Text = ""
    End With
End Sub

Private Sub AtualizaArray()
    Dim siLoop          As Integer
    Dim siOldProxTabela As Integer
    
    '------ limpando campo não utilizados
    For siLoop = 0 To 5
        If cboColunaFonte(siLoop).Text = "" Then
            cboColunaFonte(siLoop + 1).Text = ""
            cboColunaRelac(siLoop).Text = ""
            cboColunaRelac(siLoop + 1).Text = ""
        End If
    Next siLoop
    For siLoop = 0 To 5
        If cboColunaRelac(siLoop).Text = "" Then
            cboColunaRelac(siLoop + 1).Text = ""
            cboColunaFonte(siLoop).Text = ""
            cboColunaFonte(siLoop + 1).Text = ""
        End If
    Next siLoop
    
    If msOpcao = "Alt" Then
        siOldProxTabela = piProxTabela
        piProxTabela = SP01wxStoredProced.lvwTabelaRelac.SelectedItem.Index - 1
    End If
    '------ armazenando nome da tabela
    paNomeTabelas(piProxTabela) = Mid(cboTabelaRelac.BoundText, InStr(1, cboTabelaRelac.BoundText, "/") + 1, Len(cboTabelaRelac.BoundText))
    '------ armazenando tipo de relacionamento
    If optInnerJoin.Value Then paTipoRelac(piProxTabela) = enumInnerJoin
    If optLeftJoin.Value Then paTipoRelac(piProxTabela) = enumLeftJoin
    If optRightJoin.Value Then paTipoRelac(piProxTabela) = enumRightJoin
    If optFullJoin.Value Then paTipoRelac(piProxTabela) = enumFullJoin
    '------ armazenando campos relacionados
    piProxCpoRelac(piProxTabela) = 0
    For siLoop = 0 To 6
        paCamposFonte(piProxTabela, piProxCpoRelac(piProxTabela)) = cboColunaFonte(siLoop).BoundText
        paCamposRelac(piProxTabela, piProxCpoRelac(piProxTabela)) = cboColunaRelac(siLoop).BoundText
        If cboColunaFonte(siLoop).BoundText <> "" Then
            piProxCpoRelac(piProxTabela) = piProxCpoRelac(piProxTabela) + 1
        End If
    Next siLoop
    '------ armazenando campos relacionados da tabela e atributos
    piProxCpoTabRelac(piProxTabela) = 0
    For siLoop = 0 To piQtdCpoTabRelac
        If siLoop <= SP01wxCamposRelac.lvwCamposRelac.ListItems.Count - 1 Then
            '------ armazenando campos relacionados da tabela
            paCamposTabRelac(piProxTabela, piProxCpoTabRelac(piProxTabela)) = SP01wxCamposRelac.lvwCamposRelac.ListItems.Item(siLoop + 1).SubItems(1)
            '------ armazenando campos relacionados da tabela (atributos)
            paCamposTabRelacChk(piProxTabela, piProxCpoTabRelac(piProxTabela)) = SP01wxCamposRelac.lvwCamposRelac.ListItems.Item(siLoop + 1).Checked
            paCamposTabRelacAli(piProxTabela, piProxCpoTabRelac(piProxTabela)) = SP01wxCamposRelac.lvwCamposRelac.ListItems.Item(siLoop + 1).SubItems(2)
            
            piProxCpoTabRelac(piProxTabela) = piProxCpoTabRelac(piProxTabela) + 1
        Else
            '------ armazenando campos relacionados da tabela (limpando)
            paCamposTabRelac(piProxTabela, siLoop) = ""
            '------ armazenando campos relacionados da tabela (atributos) (limpando)
            paCamposTabRelacChk(piProxTabela, siLoop) = False
            paCamposTabRelacAli(piProxTabela, siLoop) = ""
        End If
    Next siLoop
    If msOpcao = "Inc" Then
        piProxTabela = piProxTabela + 1
    Else
        piProxTabela = siOldProxTabela
    End If
End Sub

