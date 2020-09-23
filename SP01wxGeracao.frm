VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form SP01wxGeracao 
   Caption         =   "Gerador de Stored Procedure "
   ClientHeight    =   7605
   ClientLeft      =   135
   ClientTop       =   750
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   11655
   Begin TabDlg.SSTab tabStoredProcedGeradas 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "SP01wxGeracao.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "rtbTextoGeradoConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "dlgSalvarComo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Inclusão"
      TabPicture(1)   =   "SP01wxGeracao.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "rtbTextoGeradoInclusao"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Alteração"
      TabPicture(2)   =   "SP01wxGeracao.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "rtbTextoGeradoAlteracao"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Exclusão"
      TabPicture(3)   =   "SP01wxGeracao.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "rtbTextoGeradoExclusao"
      Tab(3).ControlCount=   1
      Begin MSComDlg.CommonDialog dlgSalvarComo 
         Left            =   660
         Top             =   780
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DialogTitle     =   "Salvar como"
         InitDir         =   "C:\MSSQL7\Binn"
      End
      Begin RichTextLib.RichTextBox rtbTextoGeradoConsulta 
         Height          =   7155
         Left            =   60
         TabIndex        =   1
         Top             =   360
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   12621
         _Version        =   393217
         ScrollBars      =   3
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"SP01wxGeracao.frx":0070
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtbTextoGeradoInclusao 
         Height          =   7155
         Left            =   -74940
         TabIndex        =   2
         Top             =   360
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   12621
         _Version        =   393217
         ScrollBars      =   3
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"SP01wxGeracao.frx":015E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtbTextoGeradoAlteracao 
         Height          =   7155
         Left            =   -74940
         TabIndex        =   3
         Top             =   360
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   12621
         _Version        =   393217
         ScrollBars      =   3
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"SP01wxGeracao.frx":024C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtbTextoGeradoExclusao 
         Height          =   7155
         Left            =   -74940
         TabIndex        =   4
         Top             =   360
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   12621
         _Version        =   393217
         ScrollBars      =   3
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"SP01wxGeracao.frx":033A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuSalvarComo 
         Caption         =   "S&alvar como..."
      End
      Begin VB.Menu d1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSair 
         Caption         =   "Sai&r"
      End
   End
   Begin VB.Menu mnuEditar 
      Caption         =   "&Editar"
      Begin VB.Menu mnuSelecionarTudo 
         Caption         =   "&Selecionar Tudo"
      End
      Begin VB.Menu mnuCopiar 
         Caption         =   "&Copiar"
      End
   End
   Begin VB.Menu mnupopup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu mnupopSelecionarTudo 
         Caption         =   "&Selecionar Tudo"
      End
      Begin VB.Menu mnupopCopiar 
         Caption         =   "&Copiar"
      End
      Begin VB.Menu mnupopSalvarComo 
         Caption         =   "S&alvar como..."
      End
   End
End
Attribute VB_Name = "SP01wxGeracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public psTextoGeradoConsulta    As String
Public psTextoGeradoInclusao    As String
Public psTextoGeradoAlteracao   As String
Public psTextoGeradoExclusao    As String

Private Sub Form_Activate()
    rtbTextoGeradoConsulta.Text = IIf(Len(psTextoGeradoConsulta) > 0, psTextoGeradoConsulta, "Nenhum texto gerado")
    rtbTextoGeradoInclusao.Text = IIf(Len(psTextoGeradoInclusao) > 0, psTextoGeradoInclusao, "Nenhum texto gerado")
    rtbTextoGeradoAlteracao.Text = IIf(Len(psTextoGeradoAlteracao) > 0, psTextoGeradoAlteracao, "Nenhum texto gerado")
    rtbTextoGeradoExclusao.Text = IIf(Len(psTextoGeradoExclusao) > 0, psTextoGeradoExclusao, "Nenhum texto gerado")
End Sub

Private Sub mnuSair_Click()
    Unload Me
End Sub

Private Sub MostraMenuPopUP(ByVal Button As Integer)
    If Button = 2 Then
        PopupMenu mnupopup
    End If
End Sub

Private Sub mnupopCopiar_Click()
    Call mnuCopiar_Click
End Sub

Private Sub mnupopSelecionarTudo_Click()
    Call mnuSelecionarTudo_Click
End Sub

Private Sub mnupopSalvarComo_Click()
    Call mnuSalvarComo_Click
End Sub

Private Sub mnuSalvarComo_Click()
    On Error GoTo ErrHandler
    Select Case tabStoredProcedGeradas.Tab
        Case 0
            dlgSalvarComo.FileName = SP01wxStoredProced.txtNomeConsulta
        Case 1
            dlgSalvarComo.FileName = SP01wxStoredProced.txtNomeInclusao
        Case 2
            dlgSalvarComo.FileName = SP01wxStoredProced.txtNomeAlteracao
        Case 3
            dlgSalvarComo.FileName = SP01wxStoredProced.txtNomeExclusao
    End Select
    dlgSalvarComo.FileName = UCase(dlgSalvarComo.FileName) & ".SQL"
    dlgSalvarComo.Filter = "Comandos SQL (*.SQL) |*.SQL"
    dlgSalvarComo.ShowSave
    If dlgSalvarComo.FileName <> "" Then
        Select Case tabStoredProcedGeradas.Tab
            Case 0
                rtbTextoGeradoConsulta.SaveFile dlgSalvarComo.FileName, rtfText
            Case 1
                rtbTextoGeradoInclusao.SaveFile dlgSalvarComo.FileName, rtfText
            Case 2
                rtbTextoGeradoAlteracao.SaveFile dlgSalvarComo.FileName, rtfText
            Case 3
                rtbTextoGeradoExclusao.SaveFile dlgSalvarComo.FileName, rtfText
        End Select
    End If
    Exit Sub
ErrHandler:

End Sub

Private Sub mnuSelecionarTudo_Click()
    Select Case tabStoredProcedGeradas.Tab
        Case 0
            Call SelecionaTexto(rtbTextoGeradoConsulta)
        Case 1
            Call SelecionaTexto(rtbTextoGeradoInclusao)
        Case 2
            Call SelecionaTexto(rtbTextoGeradoAlteracao)
        Case 3
            Call SelecionaTexto(rtbTextoGeradoExclusao)
    End Select
End Sub

Private Sub mnuCopiar_Click()
    Select Case tabStoredProcedGeradas.Tab
        Case 0
            Call SelecionaTextoClipboard(rtbTextoGeradoConsulta)
        Case 1
            Call SelecionaTextoClipboard(rtbTextoGeradoInclusao)
        Case 2
            Call SelecionaTextoClipboard(rtbTextoGeradoAlteracao)
        Case 3
            Call SelecionaTextoClipboard(rtbTextoGeradoExclusao)
    End Select
End Sub

Private Sub SelecionaTexto(ByRef poObjTexto As RichTextBox)
    poObjTexto.SelStart = 0
    poObjTexto.SelLength = Len(rtbTextoGeradoConsulta.Text)
    poObjTexto.SetFocus
End Sub

Private Sub SelecionaTextoClipboard(ByRef poObjTexto As RichTextBox)
    Clipboard.Clear
    Clipboard.SetText poObjTexto.SelText
End Sub

Private Sub rtbTextoGeradoConsulta_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MostraMenuPopUP(Button)
End Sub

Private Sub rtbTextoGeradoInclusao_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MostraMenuPopUP(Button)
End Sub

Private Sub rtbTextoGeradoAlteracao_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MostraMenuPopUP(Button)
End Sub

Private Sub rtbTextoGeradoExclusao_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MostraMenuPopUP(Button)
End Sub


