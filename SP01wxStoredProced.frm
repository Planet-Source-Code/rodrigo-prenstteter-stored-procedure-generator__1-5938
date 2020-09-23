VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SP01wxStoredProced 
   Caption         =   "Gerador de Stored Procedure "
   ClientHeight    =   7245
   ClientLeft      =   1320
   ClientTop       =   735
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tabStoredProced 
      Height          =   7095
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Configuração"
      TabPicture(0)   =   "SP01wxStoredProced.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmConexao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmOrigemDados"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmProcedures"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frmChavePrimaria"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frmOpcoes"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdGeracao"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Consulta"
      TabPicture(1)   =   "SP01wxStoredProced.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmRelacionamento"
      Tab(1).Control(1)=   "frmAtributoConsulta"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Inclusão"
      TabPicture(2)   =   "SP01wxStoredProced.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frmAtributoInclusao"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Alteração"
      TabPicture(3)   =   "SP01wxStoredProced.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "frmAtributoAlteracao"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Usuários"
      TabPicture(4)   =   "SP01wxStoredProced.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "frmUsuarios"
      Tab(4).ControlCount=   1
      Begin VB.Frame frmRelacionamento 
         Caption         =   "Relacionamentos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2955
         Left            =   -74880
         TabIndex        =   62
         Top             =   4020
         Width           =   8715
         Begin VB.CommandButton cmdConsultaRelacInc 
            Caption         =   "+"
            Height          =   675
            Left            =   8160
            TabIndex        =   28
            Top             =   240
            Width           =   460
         End
         Begin VB.CommandButton cmdConsultaRelacAlt 
            Caption         =   "->"
            Height          =   675
            Left            =   8160
            TabIndex        =   29
            Top             =   960
            Width           =   460
         End
         Begin VB.CommandButton cmdConsultaRelacExc 
            Caption         =   "-"
            Height          =   675
            Left            =   8160
            TabIndex        =   30
            Top             =   1680
            Width           =   460
         End
         Begin MSComctlLib.ListView lvwTabelaRelac 
            Height          =   2595
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   7995
            _ExtentX        =   14102
            _ExtentY        =   4577
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
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
      End
      Begin VB.Frame frmUsuarios 
         Caption         =   "Usuários que receberão GRANT EXEC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -74880
         TabIndex        =   61
         Top             =   360
         Width           =   8715
         Begin VB.CommandButton cmdUsuariosMarcarTodos 
            Caption         =   "xxx"
            Height          =   675
            Left            =   8160
            TabIndex        =   42
            Top             =   240
            Width           =   460
         End
         Begin VB.CommandButton cmdUsuariosDesmarcarTodos 
            Caption         =   "---"
            Height          =   675
            Left            =   8160
            TabIndex        =   43
            Top             =   960
            Width           =   460
         End
         Begin MSComctlLib.ListView lvwUsuarios 
            Height          =   6255
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   7995
            _ExtentX        =   14102
            _ExtentY        =   11033
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
            NumItems        =   0
         End
      End
      Begin VB.Frame frmAtributoAlteracao 
         Caption         =   "Atributos a serem alterados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -74880
         TabIndex        =   59
         Top             =   360
         Width           =   8715
         Begin VB.CommandButton cmdAlteracaoDesmarcarTodos 
            Caption         =   "---"
            Height          =   675
            Left            =   8160
            TabIndex        =   40
            Top             =   2400
            Width           =   460
         End
         Begin VB.CommandButton cmdAlteracaoMarcarTodos 
            Caption         =   "xxx"
            Height          =   675
            Left            =   8160
            TabIndex        =   39
            Top             =   1680
            Width           =   460
         End
         Begin VB.CommandButton cmdAlteracaoAbaixo 
            Caption         =   "v"
            Height          =   675
            Left            =   8160
            TabIndex        =   38
            Top             =   960
            Width           =   460
         End
         Begin VB.CommandButton cmdAlteracaoAcima 
            Caption         =   "^"
            Height          =   675
            Left            =   8160
            TabIndex        =   37
            Top             =   240
            Width           =   460
         End
         Begin MSComctlLib.ListView lvwAlteracao 
            Height          =   6255
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   7995
            _ExtentX        =   14102
            _ExtentY        =   11033
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
      End
      Begin VB.CommandButton cmdGeracao 
         Caption         =   "&Geração Stored Procedure"
         Enabled         =   0   'False
         Height          =   675
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   6180
         Width           =   2715
      End
      Begin VB.Frame frmAtributoInclusao 
         Caption         =   "Atributos a serem gravados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -74880
         TabIndex        =   58
         Top             =   360
         Width           =   8715
         Begin VB.CommandButton cmdInclusaoAcima 
            Caption         =   "^"
            Height          =   675
            Left            =   8160
            TabIndex        =   32
            Top             =   240
            Width           =   460
         End
         Begin VB.CommandButton cmdInclusaoAbaixo 
            Caption         =   "v"
            Height          =   675
            Left            =   8160
            TabIndex        =   33
            Top             =   960
            Width           =   460
         End
         Begin VB.CommandButton cmdInclusaoMarcarTodos 
            Caption         =   "xxx"
            Height          =   675
            Left            =   8160
            TabIndex        =   34
            Top             =   1680
            Width           =   460
         End
         Begin VB.CommandButton cmdInclusaoDesmarcarTodos 
            Caption         =   "---"
            Height          =   675
            Left            =   8160
            TabIndex        =   35
            Top             =   2400
            Width           =   460
         End
         Begin MSComctlLib.ListView lvwInclusao 
            Height          =   6255
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   7995
            _ExtentX        =   14102
            _ExtentY        =   11033
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
      End
      Begin VB.Frame frmOpcoes 
         Caption         =   "Opções"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   6120
         TabIndex        =   57
         Top             =   5100
         Width           =   2715
         Begin VB.CheckBox chkOpcaoComentario 
            Caption         =   "Insere comentário inicial"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   2445
         End
         Begin VB.CheckBox chkOpcaoGrant 
            Caption         =   "Insere comando GRANT EXEC"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   720
            Width           =   2565
         End
         Begin VB.CheckBox chkOpcaoDrop 
            Caption         =   "Insere comando DROP PROC"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   480
            Width           =   2445
         End
      End
      Begin VB.Frame frmAtributoConsulta 
         Caption         =   "Atributos a serem consultados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   -74880
         TabIndex        =   56
         Top             =   360
         Width           =   8715
         Begin VB.CommandButton cmdConsultaDesmarcarTodos 
            Caption         =   "---"
            Height          =   675
            Left            =   8160
            TabIndex        =   26
            Top             =   2400
            Width           =   460
         End
         Begin VB.CommandButton cmdConsultaMarcarTodos 
            Caption         =   "xxx"
            Height          =   675
            Left            =   8160
            TabIndex        =   25
            Top             =   1680
            Width           =   460
         End
         Begin VB.CommandButton cmdConsultaAbaixo 
            Caption         =   "v"
            Height          =   675
            Left            =   8160
            TabIndex        =   24
            Top             =   960
            Width           =   460
         End
         Begin VB.CommandButton cmdConsultaAcima 
            Caption         =   "^"
            Height          =   675
            Left            =   8160
            TabIndex        =   23
            Top             =   240
            Width           =   460
         End
         Begin MSComctlLib.ListView lvwConsulta 
            Height          =   3255
            Left            =   120
            TabIndex        =   22
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
      End
      Begin VB.Frame frmChavePrimaria 
         Caption         =   "Chave Primária (PK)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   120
         TabIndex        =   55
         Top             =   2880
         Width           =   8715
         Begin MSComctlLib.ImageList imgImagens 
            Left            =   7860
            Top             =   1440
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   43
            ImageHeight     =   43
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   3
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "SP01wxStoredProced.frx":008C
                  Key             =   "Login"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "SP01wxStoredProced.frx":170C
                  Key             =   "Logout"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "SP01wxStoredProced.frx":2D8C
                  Key             =   "Geracao"
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid grdChavePrimaria 
            Height          =   1815
            Left            =   180
            TabIndex        =   9
            Top             =   240
            Width           =   8355
            _ExtentX        =   14737
            _ExtentY        =   3201
            _Version        =   393216
            AllowUpdate     =   0   'False
            Enabled         =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   "KEY_SEQ"
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "COLUMN_NAME"
               Caption         =   "Nome Coluna"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               AllowRowSizing  =   0   'False
               AllowSizing     =   0   'False
               BeginProperty Column00 
                  ColumnAllowSizing=   0   'False
                  ColumnWidth     =   480,189
               EndProperty
               BeginProperty Column01 
                  ColumnAllowSizing=   0   'False
                  ColumnWidth     =   7529,953
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame frmProcedures 
         Caption         =   "Procedures a gerar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   120
         TabIndex        =   50
         Top             =   5100
         Width           =   5895
         Begin VB.TextBox txtNomeExclusao 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1725
            TabIndex        =   17
            Top             =   1320
            Width           =   4000
         End
         Begin VB.TextBox txtNomeAlteracao 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1725
            TabIndex        =   15
            Top             =   960
            Width           =   4000
         End
         Begin VB.TextBox txtNomeInclusao 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1725
            TabIndex        =   13
            Top             =   600
            Width           =   4000
         End
         Begin VB.TextBox txtNomeConsulta 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1725
            TabIndex        =   11
            Top             =   240
            Width           =   4000
         End
         Begin VB.CheckBox chkExclusao 
            Caption         =   "Exclusão"
            Enabled         =   0   'False
            Height          =   195
            Left            =   180
            TabIndex        =   16
            Top             =   1380
            Width           =   1000
         End
         Begin VB.CheckBox chkAlteracao 
            Caption         =   "Alteração"
            Enabled         =   0   'False
            Height          =   195
            Left            =   180
            TabIndex        =   14
            Top             =   1020
            Width           =   1000
         End
         Begin VB.CheckBox chkInclusao 
            Caption         =   "Inclusão"
            Enabled         =   0   'False
            Height          =   195
            Left            =   180
            TabIndex        =   12
            Top             =   660
            Width           =   1000
         End
         Begin VB.CheckBox chkConsulta 
            Caption         =   "Consulta"
            Enabled         =   0   'False
            Height          =   195
            Left            =   180
            TabIndex        =   10
            Top             =   300
            Width           =   1000
         End
         Begin VB.Label lblNomeExclusao 
            Caption         =   "Nome "
            Enabled         =   0   'False
            Height          =   255
            Left            =   1200
            TabIndex        =   54
            Top             =   1380
            Width           =   1275
         End
         Begin VB.Label lblNomeAlteracao 
            Caption         =   "Nome"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1200
            TabIndex        =   53
            Top             =   1020
            Width           =   1275
         End
         Begin VB.Label lblNomeInclusao 
            Caption         =   "Nome"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1200
            TabIndex        =   52
            Top             =   660
            Width           =   1275
         End
         Begin VB.Label lblNomeConsulta 
            Caption         =   "Nome"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1200
            TabIndex        =   51
            Top             =   300
            Width           =   1275
         End
      End
      Begin VB.Frame frmOrigemDados 
         Caption         =   "Origem dos dados"
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
         Left            =   120
         TabIndex        =   48
         Top             =   2160
         Width           =   8715
         Begin VB.CheckBox chkIncluirView 
            Caption         =   "Incluir Views"
            Height          =   255
            Left            =   7320
            TabIndex        =   8
            Top             =   300
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo cboTabela 
            Height          =   315
            Left            =   1620
            TabIndex        =   7
            Top             =   240
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label lblTabela 
            Caption         =   "Tabela"
            Enabled         =   0   'False
            Height          =   255
            Left            =   180
            TabIndex        =   49
            Top             =   300
            Width           =   1395
         End
      End
      Begin VB.Frame frmConexao 
         Caption         =   "Conexão"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   8715
         Begin VB.CheckBox chkTrusted 
            Caption         =   "Trusted Connection"
            Height          =   255
            Left            =   5460
            TabIndex        =   2
            Top             =   660
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.TextBox txtDatabase 
            Height          =   315
            Left            =   1620
            TabIndex        =   5
            Text            =   "administracao"
            Top             =   1320
            Width           =   5535
         End
         Begin VB.TextBox txtSenha 
            Enabled         =   0   'False
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1620
            PasswordChar    =   "*"
            TabIndex        =   4
            Top             =   960
            Width           =   5535
         End
         Begin VB.TextBox txtUsuario 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1620
            TabIndex        =   3
            Top             =   600
            Width           =   3735
         End
         Begin VB.CommandButton cmdConectar 
            Caption         =   "&Conectar"
            Default         =   -1  'True
            Height          =   1395
            Left            =   7320
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtServidor 
            Height          =   315
            Left            =   1620
            TabIndex        =   1
            Text            =   "dcptlab72"
            Top             =   240
            Width           =   5535
         End
         Begin VB.Label lblDatabase 
            Caption         =   "Database"
            Height          =   255
            Left            =   180
            TabIndex        =   60
            Top             =   1380
            Width           =   1395
         End
         Begin VB.Label lblSenha 
            Caption         =   "Senha"
            Height          =   255
            Left            =   180
            TabIndex        =   47
            Top             =   1020
            Width           =   1395
         End
         Begin VB.Label lblUsuario 
            Caption         =   "Nome do Usuário"
            Height          =   255
            Left            =   180
            TabIndex        =   46
            Top             =   660
            Width           =   1395
         End
         Begin VB.Label lblServidor 
            Caption         =   "Nome do Servidor"
            Height          =   255
            Left            =   180
            TabIndex        =   45
            Top             =   300
            Width           =   1395
         End
      End
   End
End
Attribute VB_Name = "SP01wxStoredProced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum enumMarcarColunas
    enumMarcar = True
    enumDesmarcar = False
End Enum

Public Enum enumMovimentoLinha
    enumAcima = -1
    enumAbaixo = 1
End Enum

Private Enum enumOpcao
    enumConsulta = 1
    enumInclusao
    enumAlteracao
    enumExclusao
End Enum

Public moConexao                As ADODB.Connection
Public moRSetDatabase           As ADODB.Recordset
Public moRSetTabela             As ADODB.Recordset
Public moRSetIndicePK           As ADODB.Recordset
Public moRSetCamposTabela       As ADODB.Recordset
Public moRSetUsuarios           As ADODB.Recordset

Private Sub Form_Load()
    Call InicializaAplicativo
End Sub

Private Sub chkTrusted_Click()
    If chkTrusted.Value = vbChecked Then
        txtUsuario.Text = ""
        txtUsuario.Enabled = False
        txtSenha.Text = ""
        txtSenha.Enabled = False
    Else
        txtUsuario.Enabled = True
        txtSenha.Enabled = True
        txtUsuario.SetFocus
    End If
End Sub

Private Sub cmdConectar_Click()
    Dim siContador          As Integer
    
    On Error GoTo ErroConexao
    If moConexao Is Nothing Then Set moConexao = New ADODB.Connection
    If moConexao.State = adStateOpen Then
        Call InicializaAplicativo
    Else
        '-------- checagens
        If txtServidor.Text = "" Then
            MsgBox "Nome do servidor deve ser digitado", vbOKOnly + vbInformation, "Erro"
            Exit Sub
        End If
        If chkTrusted.Value = vbUnchecked Then
            If txtUsuario.Text = "" Then
                MsgBox "Nome do usuário deve ser digitado", vbOKOnly + vbInformation, "Erro"
                Exit Sub
            End If
        End If
        If txtDatabase.Text = "" Then
            MsgBox "Nome do database deve ser digitado", vbOKOnly + vbInformation, "Erro"
            Exit Sub
        End If
        Me.MousePointer = vbHourglass
        '-------- executa conexão
        moConexao.Errors.Clear
        With moConexao
            .CommandTimeout = 10
            If chkTrusted.Value = vbChecked Then
                .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & txtDatabase & ";Data Source=" & txtServidor.Text
            Else
                .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;Initial Catalog=" & txtDatabase & ";User ID=" & txtUsuario.Text & ";PWD=" & txtSenha.Text & ";Data Source=" & txtServidor.Text
            End If
            .ConnectionTimeout = 10
            .CursorLocation = adUseClient
            .Open
        End With
        If moConexao.Errors.Count > 0 Then GoTo ErroConexao
        '-------- libera e carrega tabelas
        If Not CarregaTabelas() Then GoTo ErroConexao
        '-------- carrega usuários do database
        moConexao.Errors.Clear
        Set moRSetUsuarios = New ADODB.Recordset
        Set moRSetUsuarios = moConexao.Execute("sp_helpuser")
        If moConexao.Errors.Count > 0 Then GoTo ErroConexao
        lvwUsuarios.ColumnHeaders.Clear
        lvwUsuarios.ColumnHeaders.Add 1, , "Nome Usuário", 2000
        lvwUsuarios.ColumnHeaders.Add 2, , "Nome Login"
        lvwUsuarios.ColumnHeaders.Item(2).Width = lvwUsuarios.Width - _
                                                  lvwUsuarios.ColumnHeaders.Item(1).Width
        lvwUsuarios.ListItems.Clear
        siContador = 1
        While Not moRSetUsuarios.EOF
            If Not IsNull(moRSetUsuarios!LoginName) And RTrim(moRSetUsuarios!LoginName) <> "sa" Then
                lvwUsuarios.ListItems.Add , "U" & Format(siContador, "00000"), moRSetUsuarios!UserName
                lvwUsuarios.ListItems.Item("U" & Format(siContador, "00000")).SubItems(1) = moRSetUsuarios!LoginName
                
                lvwUsuarios.ListItems.Item("U" & Format(siContador, "00000")).Checked = True
                siContador = siContador + 1
            End If
            moRSetUsuarios.MoveNext
        Wend
        Call MarcaColunas(enumMarcar, lvwUsuarios)
        '-------- atualiza botão conectar/desconectar
        cmdConectar.Caption = "&Desconectar"
        cmdConectar.Picture = imgImagens.ListImages.Item("Logout").Picture
        '-------- protege parâmetros de login
        chkTrusted.Enabled = False
        txtServidor.Enabled = False
        txtUsuario.Enabled = False
        txtSenha.Enabled = False
        txtDatabase.Enabled = False
        
        Me.MousePointer = vbNormal
    End If
    Exit Sub
ErroConexao:
    Me.MousePointer = vbNormal
    If moConexao.Errors.Count > 0 Then
        MsgBox "Erro ao executar conexão com servidor de dados" & vbCrLf & vbCrLf & "(" & moConexao.Errors.Item(0).NativeError & ") " & moConexao.Errors.Item(0).Description, vbOKOnly + vbCritical, "Erro"
    Else
        MsgBox "Erro ao executar conexão com servidor de dados" & vbCrLf & vbCrLf & Err.Description, vbOKOnly + vbCritical, "Erro"
    End If
    
    Call InicializaAplicativo
End Sub

Private Sub chkIncluirView_Click()
    If txtDatabase.Text <> "" Then
        If Not CarregaTabelas() Then
            Call InicializaAplicativo
            Exit Sub
        End If
        If cboTabela.Text <> "" Then
            If chkIncluirView.Value = vbChecked Then
                chkInclusao.Enabled = False
                chkAlteracao.Enabled = False
                chkExclusao.Enabled = False
                
                chkConsulta.Value = vbUnchecked
                chkInclusao.Value = vbUnchecked
                chkAlteracao.Value = vbUnchecked
                chkExclusao.Value = vbUnchecked
            Else
                chkInclusao.Enabled = True
                chkAlteracao.Enabled = True
                chkExclusao.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub cboTabela_Change()
    On Error GoTo ErroTabela
    If cboTabela.Text = "" Then
        Call InicializaPK
        Call InicializaCheck
        Call InicializaOpcoes
    Else
        '-------- gera novo nome de procedure
        If chkConsulta.Value = vbChecked Then
            Call MontaNomeProcedure(txtNomeConsulta, _
                                    tabStoredProced, _
                                    enumConsulta, _
                                    vbChecked, _
                                    Mid(cboTabela.BoundText, InStr(1, cboTabela.BoundText, "/") + 1, Len(cboTabela.BoundText)), _
                                    False)
        End If
        If chkInclusao.Value = vbChecked Then
            Call MontaNomeProcedure(txtNomeInclusao, _
                                    tabStoredProced, _
                                    enumInclusao, _
                                    vbChecked, _
                                    Mid(cboTabela.BoundText, InStr(1, cboTabela.BoundText, "/") + 1, Len(cboTabela.BoundText)), _
                                    False)
        End If
        If chkAlteracao.Value = vbChecked Then
            Call MontaNomeProcedure(txtNomeAlteracao, _
                                    tabStoredProced, _
                                    enumAlteracao, _
                                    vbChecked, _
                                    Mid(cboTabela.BoundText, InStr(1, cboTabela.BoundText, "/") + 1, Len(cboTabela.BoundText)), _
                                    False)
        End If
        If chkExclusao.Value = vbChecked Then
            Call MontaNomeProcedure(txtNomeExclusao, _
                                    tabStoredProced, _
                                    enumExclusao, _
                                    vbChecked, _
                                    Mid(cboTabela.BoundText, InStr(1, cboTabela.BoundText, "/") + 1, Len(cboTabela.BoundText)), _
                                    False)
        End If
        '-------- recupera campos da PK
        grdChavePrimaria.Enabled = True
        
        moConexao.Errors.Clear
        Set moRSetIndicePK = New ADODB.Recordset
        Set moRSetIndicePK = moConexao.Execute("sp_pkeys '" & Mid(cboTabela.BoundText, InStr(1, cboTabela.BoundText, "/") + 1, Len(cboTabela.BoundText)) & "'")
        If moConexao.Errors.Count > 0 Then GoTo ErroTabela
        Set grdChavePrimaria.DataSource = moRSetIndicePK
        grdChavePrimaria.Refresh
        
        chkConsulta.Enabled = True
        chkConsulta.Value = vbChecked
        If chkIncluirView.Value = vbUnchecked Or _
           Mid(cboTabela.BoundText, 1, InStr(1, cboTabela.BoundText, "/") - 1) = "U" Then
            chkInclusao.Enabled = True
            chkAlteracao.Enabled = True
            chkExclusao.Enabled = True
            
            chkInclusao.Value = vbChecked
            chkAlteracao.Value = vbChecked
            chkExclusao.Value = vbChecked
        End If
        
        Set moRSetCamposTabela = New ADODB.Recordset
        Set moRSetCamposTabela = moConexao.Execute("sp_columns '" & Mid(cboTabela.BoundText, InStr(1, cboTabela.BoundText, "/") + 1, Len(cboTabela.BoundText)) & "'")
        '-------- monta listview de consulta/relacionamentos/inclusão/alteração/exclusão
        If Not MontaListView(lvwConsulta, moRSetCamposTabela) Then GoTo ErroTabela
        lvwTabelaRelac.ColumnHeaders.Clear
        lvwTabelaRelac.ColumnHeaders.Add 1, , "Relacionamento", 1400
        lvwTabelaRelac.ColumnHeaders.Add 2, , "Nome Tabela"
        lvwTabelaRelac.ColumnHeaders.Item(2).Width = lvwTabelaRelac.Width - _
                                                     lvwTabelaRelac.ColumnHeaders.Item(1).Width
        lvwTabelaRelac.View = lvwReport
        lvwTabelaRelac.ListItems.Clear
        
        Erase paNomeTabelas
        Erase paTipoRelac
        Erase paCamposFonte
        Erase paCamposRelac
        Erase paCamposTabRelac
        Erase paCamposTabRelacChk
        Erase paCamposTabRelacAli
        
        piProxTabela = 0
        Erase piProxCpoRelac
        Erase piProxCpoTabRelac
        
        If Not MontaListView(lvwInclusao, moRSetCamposTabela) Then GoTo ErroTabela
        If Not MontaListView(lvwAlteracao, moRSetCamposTabela) Then GoTo ErroTabela
        '-------- assinala opções de geração
        chkOpcaoDrop.Value = vbChecked
        chkOpcaoGrant.Value = vbChecked
        chkOpcaoComentario.Value = vbChecked
        
        chkOpcaoDrop.Enabled = True
        chkOpcaoGrant.Enabled = True
        chkOpcaoComentario.Enabled = True
    End If
    Exit Sub
ErroTabela:
    If moConexao.Errors.Count > 0 Then
        MsgBox "Erro ao carregar campos ou PK da tabela escolhida" & vbCrLf & vbCrLf & "(" & moConexao.Errors.Item(0).NativeError & ") " & moConexao.Errors.Item(0).Description, vbOKOnly + vbCritical, "Erro"
    Else
        MsgBox "Erro ao carregar campos ou PK da tabela escolhida" & vbCrLf & vbCrLf & Err.Description, vbOKOnly + vbCritical, "Erro"
    End If
    Call InicializaAplicativo
End Sub

Private Sub chkConsulta_Click()
    lblNomeConsulta.Enabled = (chkConsulta.Value = vbChecked)
    txtNomeConsulta.Enabled = (chkConsulta.Value = vbChecked)
    Call MontaNomeProcedure(txtNomeConsulta, _
                            tabStoredProced, _
                            enumConsulta, _
                            chkConsulta.Value, _
                            Mid(cboTabela.BoundText, InStr(1, cboTabela.BoundText, "/") + 1, Len(cboTabela.BoundText)), _
                            True)
    Call AtualizaBotaoGeracao
End Sub

Private Sub cmdConsultaMarcarTodos_Click()
    Call MarcaColunas(enumMarcar, lvwConsulta)
End Sub

Private Sub cmdConsultaDesmarcarTodos_Click()
    Call MarcaColunas(enumDesmarcar, lvwConsulta)
End Sub

Private Sub cmdConsultaAcima_Click()
    Call MovimentoLinha(enumAcima, lvwConsulta)
End Sub

Private Sub cmdConsultaAbaixo_Click()
    Call MovimentoLinha(enumAbaixo, lvwConsulta)
End Sub

Private Sub cmdConsultaRelacAlt_Click()
    If lvwTabelaRelac.ListItems.Count > 0 Then
        Load SP01wxRelacionamentos
        SP01wxRelacionamentos.msOpcao = "Alt"
        SP01wxRelacionamentos.cmdOK.Caption = "Alteração"
        SP01wxRelacionamentos.Show vbModal
    Else
        MsgBox "Não existe ítem a ser alterado", vbOKOnly + vbInformation, "Mensagem"
    End If
End Sub

Private Sub cmdConsultaRelacExc_Click()
    If lvwTabelaRelac.ListItems.Count > 0 Then
        Load SP01wxRelacionamentos
        SP01wxRelacionamentos.msOpcao = "Exc"
        SP01wxRelacionamentos.cmdOK.Caption = "Exclusão"
        SP01wxRelacionamentos.Show vbModal
    Else
        MsgBox "Não existe ítem a ser excluído", vbOKOnly + vbInformation, "Mensagem"
    End If
End Sub

Private Sub cmdConsultaRelacInc_Click()
    If lvwTabelaRelac.ListItems.Count < piQtdTabelas Then
        Load SP01wxRelacionamentos
        SP01wxRelacionamentos.msOpcao = "Inc"
        SP01wxRelacionamentos.cmdOK.Caption = "Inclusão"
        SP01wxRelacionamentos.Show vbModal
    Else
        MsgBox "Limite de relacionamentos excedido", vbOKOnly + vbInformation, "Mensagem"
    End If
End Sub

Private Sub chkInclusao_Click()
    lblNomeInclusao.Enabled = (chkInclusao.Value = vbChecked)
    txtNomeInclusao.Enabled = (chkInclusao.Value = vbChecked)
    Call MontaNomeProcedure(txtNomeInclusao, _
                            tabStoredProced, _
                            enumInclusao, _
                            chkInclusao.Value, _
                            Mid(cboTabela.BoundText, InStr(1, cboTabela.BoundText, "/") + 1, Len(cboTabela.BoundText)), _
                            True)
    Call AtualizaBotaoGeracao
End Sub

Private Sub cmdInclusaoMarcarTodos_Click()
    Call MarcaColunas(enumMarcar, lvwInclusao)
End Sub

Private Sub cmdInclusaoDesmarcarTodos_Click()
    Call MarcaColunas(enumDesmarcar, lvwInclusao)
End Sub

Private Sub cmdInclusaoAcima_Click()
    Call MovimentoLinha(enumAcima, lvwInclusao)
End Sub

Private Sub cmdInclusaoAbaixo_Click()
    Call MovimentoLinha(enumAbaixo, lvwInclusao)
End Sub

Private Sub chkAlteracao_Click()
    lblNomeAlteracao.Enabled = (chkAlteracao.Value = vbChecked)
    txtNomeAlteracao.Enabled = (chkAlteracao.Value = vbChecked)
    Call MontaNomeProcedure(txtNomeAlteracao, _
                            tabStoredProced, _
                            enumAlteracao, _
                            chkAlteracao.Value, _
                            Mid(cboTabela.BoundText, InStr(1, cboTabela.BoundText, "/") + 1, Len(cboTabela.BoundText)), _
                            True)
    Call AtualizaBotaoGeracao
End Sub

Private Sub cmdAlteracaoMarcarTodos_Click()
    Call MarcaColunas(enumMarcar, lvwAlteracao)
End Sub

Private Sub cmdAlteracaoDesmarcarTodos_Click()
    Call MarcaColunas(enumDesmarcar, lvwAlteracao)
End Sub

Private Sub cmdAlteracaoAcima_Click()
    Call MovimentoLinha(enumAcima, lvwAlteracao)
End Sub

Private Sub cmdAlteracaoAbaixo_Click()
    Call MovimentoLinha(enumAbaixo, lvwAlteracao)
End Sub

Private Sub chkExclusao_Click()
    lblNomeExclusao.Enabled = (chkExclusao.Value = vbChecked)
    txtNomeExclusao.Enabled = (chkExclusao.Value = vbChecked)
    Call MontaNomeProcedure(txtNomeExclusao, _
                            tabStoredProced, _
                            enumExclusao, _
                            chkExclusao.Value, _
                            Mid(cboTabela.BoundText, InStr(1, cboTabela.BoundText, "/") + 1, Len(cboTabela.BoundText)), _
                            True)
    Call AtualizaBotaoGeracao
End Sub

Private Sub cmdUsuariosMarcarTodos_Click()
    Call MarcaColunas(enumMarcar, lvwUsuarios)
End Sub

Private Sub cmdUsuariosDesmarcarTodos_Click()
    Call MarcaColunas(enumDesmarcar, lvwUsuarios)
End Sub

Private Sub tabStoredProced_Click(PreviousTab As Integer)
    Dim siLoop          As Integer
    Dim sbMsg           As Boolean
    
    If PreviousTab = 2 Then
        sbMsg = False
        For siLoop = 1 To lvwInclusao.ListItems.Count
            If lvwInclusao.ListItems.Item(siLoop).Checked = False And RTrim(lvwInclusao.ListItems.Item(siLoop).SubItems(4)) = "" Then
                If Not sbMsg Then
                    sbMsg = True
                    MsgBox "As colunas NÃO NULAS que estavam desmarcadas foram remarcadas para inclusão", vbOKOnly + vbInformation, "Mensagem"
                End If
                lvwInclusao.ListItems.Item(siLoop).Checked = True
            End If
        Next siLoop
        'If sbMsg Then tabStoredProced.Tab = PreviousTab
    End If
    
    If tabStoredProced.Tab = 1 Then
        lvwConsulta.SetFocus
    End If
    If tabStoredProced.Tab = 2 Then
        lvwInclusao.SetFocus
    End If
    If tabStoredProced.Tab = 3 Then
        lvwAlteracao.SetFocus
    End If
    If tabStoredProced.Tab = 4 Then
        lvwUsuarios.SetFocus
    End If
End Sub

Private Sub cmdGeracao_Click()
    Dim ssTextoGerado       As String
    
    On Error GoTo ErroGeracao
    
    If chkConsulta.Value = vbChecked Then
        If Not GeradorTexto(ssTextoGerado, enumConsulta) Then GoTo ErroGeracao
        SP01wxGeracao.psTextoGeradoConsulta = ssTextoGerado
    Else
        SP01wxGeracao.psTextoGeradoConsulta = "Opção desabilitada"
    End If
    If chkInclusao.Value = vbChecked Then
        If Not GeradorTexto(ssTextoGerado, enumInclusao) Then GoTo ErroGeracao
        SP01wxGeracao.psTextoGeradoInclusao = ssTextoGerado
    Else
        SP01wxGeracao.psTextoGeradoInclusao = "Opção desabilitada"
    End If
    If chkAlteracao.Value = vbChecked Then
        If Not GeradorTexto(ssTextoGerado, enumAlteracao) Then GoTo ErroGeracao
        SP01wxGeracao.psTextoGeradoAlteracao = ssTextoGerado
    Else
        SP01wxGeracao.psTextoGeradoAlteracao = "Opção desabilitada"
    End If
    If chkExclusao.Value = vbChecked Then
        If Not GeradorTexto(ssTextoGerado, enumExclusao) Then GoTo ErroGeracao
        SP01wxGeracao.psTextoGeradoExclusao = ssTextoGerado
    Else
        SP01wxGeracao.psTextoGeradoExclusao = "Opção desabilitada"
    End If
    
    SP01wxGeracao.Show vbModal
    Exit Sub
ErroGeracao:
    MsgBox "Stored Procedure não gerada por erros", vbOKOnly + vbCritical, "Erro"
End Sub

Private Sub chkOpcaoGrant_Click()
    tabStoredProced.TabEnabled(4) = (chkOpcaoGrant.Value = vbChecked)
End Sub

'----------------------------------------- Rotinas auxiliares
Private Sub InicializaAplicativo()
    cmdConectar.Caption = "&Conectar"
    cmdConectar.Picture = imgImagens.ListImages.Item("Login").Picture
    '-------- desprotege parâmetros de login
    txtServidor.Enabled = True
    chkTrusted.Enabled = True
    txtUsuario.Enabled = False
    txtSenha.Enabled = False
    txtDatabase.Enabled = True
    chkTrusted.Value = vbChecked
    txtUsuario.Text = ""
    txtSenha.Text = ""
    
    cmdGeracao.Picture = imgImagens.ListImages.Item("Geracao").Picture
    
    If Not moConexao Is Nothing Then
        If moConexao.State = adStateOpen Then moConexao.Close
    End If
    If Not moConexao Is Nothing Then Set moConexao = Nothing
    
    chkIncluirView.Value = vbUnchecked
    chkOpcaoGrant.Value = vbUnchecked

    lblTabela.Enabled = False
    cboTabela.Enabled = False
    cboTabela.Text = ""
    
    Call InicializaPK
    Call InicializaCheck
    
    tabStoredProced.TabEnabled(1) = False
    tabStoredProced.TabEnabled(2) = False
    tabStoredProced.TabEnabled(3) = False
    tabStoredProced.TabEnabled(4) = False
End Sub

Private Sub InicializaPK()
    grdChavePrimaria.Enabled = False
    
    Set grdChavePrimaria.DataSource = Nothing
    grdChavePrimaria.Refresh
End Sub

Private Sub InicializaCheck()
    chkConsulta.Enabled = False
    chkInclusao.Enabled = False
    chkAlteracao.Enabled = False
    chkExclusao.Enabled = False
    
    chkConsulta.Value = vbUnchecked
    chkInclusao.Value = vbUnchecked
    chkAlteracao.Value = vbUnchecked
    chkExclusao.Value = vbUnchecked
    
    lblNomeConsulta.Enabled = False
    lblNomeInclusao.Enabled = False
    lblNomeAlteracao.Enabled = False
    lblNomeExclusao.Enabled = False
    
    txtNomeConsulta.Text = ""
    txtNomeInclusao.Text = ""
    txtNomeAlteracao.Text = ""
    txtNomeExclusao.Text = ""
    
    txtNomeConsulta.Enabled = False
    txtNomeInclusao.Enabled = False
    txtNomeAlteracao.Enabled = False
    txtNomeExclusao.Enabled = False
End Sub

Private Sub InicializaOpcoes()
    chkOpcaoDrop.Value = vbUnchecked
    chkOpcaoGrant.Value = vbUnchecked
    chkOpcaoComentario.Value = vbUnchecked
    
    chkOpcaoDrop.Enabled = False
    chkOpcaoGrant.Enabled = False
    chkOpcaoComentario.Enabled = False
End Sub
Private Function CarregaTabelas() As Boolean
    On Error GoTo ErroCarregaTabelas
    
    If Not moConexao Is Nothing Then
        If moConexao.State = adStateOpen Then
            lblTabela.Enabled = True
            cboTabela.Enabled = True
            cboTabela.Text = ""
            
            moConexao.Errors.Clear
            Set moRSetTabela = New ADODB.Recordset
            Set moRSetTabela = moConexao.Execute("select name, rtrim(xtype)+'/'+name as tiponome from sysobjects where xtype = 'U' " & IIf(chkIncluirView.Value = vbChecked, "or xtype = 'V'", "") & " order by name")
            If moConexao.Errors.Count > 0 Then GoTo ErroCarregaTabelas
            With cboTabela
                .BoundColumn = "tiponome"
                .ListField = "name"
                Set .RowSource = moRSetTabela
            End With
        End If
    End If
    CarregaTabelas = True
    Exit Function
ErroCarregaTabelas:
    CarregaTabelas = False
End Function

Public Function MontaListView(ByRef polvwFonte As ListView, _
                              ByRef poRSetCamposTabela As ADODB.Recordset) As Boolean
    Dim siContador          As Integer
        
    On Error GoTo ErroMontaListView
    polvwFonte.ColumnHeaders.Clear
    polvwFonte.ColumnHeaders.Add 1, , "Posição", 800
    polvwFonte.ColumnHeaders.Add 2, , "Nome Coluna"
    polvwFonte.ColumnHeaders.Add 3, , "Tipo Dado", 1000
    polvwFonte.ColumnHeaders.Add 4, , "Tamanho", 850
    polvwFonte.ColumnHeaders.Add 5, , "Nulo", 600
    polvwFonte.ColumnHeaders.Item(2).Width = polvwFonte.Width - _
                                             polvwFonte.ColumnHeaders.Item(1).Width - _
                                             polvwFonte.ColumnHeaders.Item(3).Width - _
                                             polvwFonte.ColumnHeaders.Item(4).Width - _
                                             polvwFonte.ColumnHeaders.Item(5).Width
    polvwFonte.View = lvwReport
    
    polvwFonte.ListItems.Clear
    If poRSetCamposTabela.RecordCount > 0 Then poRSetCamposTabela.MoveFirst
    siContador = 1
    While Not poRSetCamposTabela.EOF
        If Not EhChavePrimaria(poRSetCamposTabela!COLUMN_NAME) Then
            'polvwFonte.ListItems.Add , "K" & Format(poRSetCamposTabela!ORDINAL_POSITION, "00000"), siContador
            
            'polvwFonte.ListItems.Item("K" & Format(poRSetCamposTabela!ORDINAL_POSITION, "00000")).SubItems(1) = poRSetCamposTabela!COLUMN_NAME
            'polvwFonte.ListItems.Item("K" & Format(poRSetCamposTabela!ORDINAL_POSITION, "00000")).SubItems(2) = poRSetCamposTabela!TYPE_NAME
            'polvwFonte.ListItems.Item("K" & Format(poRSetCamposTabela!ORDINAL_POSITION, "00000")).SubItems(3) = IIf(InStr(1, poRSetCamposTabela!TYPE_NAME, "char") > 0, poRSetCamposTabela!LENGTH, "")
            'polvwFonte.ListItems.Item("K" & Format(poRSetCamposTabela!ORDINAL_POSITION, "00000")).SubItems(4) = IIf(poRSetCamposTabela!NULLABLE, "SIM", "")
            
            polvwFonte.ListItems.Add , "K" & poRSetCamposTabela!COLUMN_NAME, siContador
            
            polvwFonte.ListItems.Item("K" & poRSetCamposTabela!COLUMN_NAME).SubItems(1) = poRSetCamposTabela!COLUMN_NAME
            polvwFonte.ListItems.Item("K" & poRSetCamposTabela!COLUMN_NAME).SubItems(2) = poRSetCamposTabela!TYPE_NAME
            polvwFonte.ListItems.Item("K" & poRSetCamposTabela!COLUMN_NAME).SubItems(3) = IIf(InStr(1, poRSetCamposTabela!TYPE_NAME, "char") > 0, poRSetCamposTabela!LENGTH, "")
            polvwFonte.ListItems.Item("K" & poRSetCamposTabela!COLUMN_NAME).SubItems(4) = IIf(poRSetCamposTabela!NULLABLE, "SIM", "")
            siContador = siContador + 1
        End If
        
        poRSetCamposTabela.MoveNext
    Wend
    Call MarcaColunas(enumMarcar, polvwFonte)
    MontaListView = True
    Exit Function
ErroMontaListView:
    MontaListView = False
End Function

Public Sub MarcaColunas(ByVal pbMarcar As enumMarcarColunas, _
                        ByRef polvwFonte As ListView)
    Dim siLoop          As Integer
    
    For siLoop = 1 To polvwFonte.ListItems.Count
        polvwFonte.ListItems.Item(siLoop).Checked = pbMarcar
    Next siLoop
    'polvwFonte.SetFocus
End Sub

Public Sub MovimentoLinha(ByVal pbMovimento As enumMovimentoLinha, _
                          ByRef polvwFonte As ListView)
    Dim sbMarcado          As Boolean
    
    If (pbMovimento = enumAcima And polvwFonte.SelectedItem.Index > 1) Or _
       (pbMovimento = enumAbaixo And polvwFonte.SelectedItem.Index < polvwFonte.ListItems.Count) Then
        polvwFonte.ListItems.Add , , 0
        Call CopiaDadosListItem(polvwFonte.SelectedItem, polvwFonte.ListItems.Item(polvwFonte.ListItems.Count))
        sbMarcado = polvwFonte.SelectedItem.Checked
        
        Call CopiaDadosListItem(polvwFonte.ListItems.Item(polvwFonte.SelectedItem.Index + pbMovimento), polvwFonte.SelectedItem)
        polvwFonte.SelectedItem.Checked = polvwFonte.ListItems.Item(polvwFonte.SelectedItem.Index + pbMovimento).Checked
        
        Call CopiaDadosListItem(polvwFonte.ListItems.Item(polvwFonte.ListItems.Count), polvwFonte.ListItems.Item(polvwFonte.SelectedItem.Index + pbMovimento))
        polvwFonte.ListItems.Item(polvwFonte.SelectedItem.Index + pbMovimento).Checked = sbMarcado
        
        polvwFonte.ListItems.Item(polvwFonte.SelectedItem.Index + pbMovimento).Selected = True
        polvwFonte.ListItems.Remove polvwFonte.ListItems.Count
        polvwFonte.Refresh
    End If
    polvwFonte.SetFocus
End Sub

Private Sub CopiaDadosListItem(ByRef ListItemFonte As ListItem, _
                               ByRef ListItemDestino As ListItem)
    Dim siLoop          As Integer
    
    For siLoop = 1 To ListItemFonte.ListSubItems.Count
        ListItemDestino.SubItems(siLoop) = ListItemFonte.SubItems(siLoop)
    Next siLoop
End Sub

Private Sub AtualizaBotaoGeracao()
    cmdGeracao.Enabled = (chkConsulta.Value = vbChecked) Or _
                         (chkInclusao.Value = vbChecked) Or _
                         (chkAlteracao.Value = vbChecked) Or _
                         (chkExclusao.Value = vbChecked)
End Sub

Private Function GeradorTexto(ByRef psTextoGerado As String, _
                              ByVal peOpcao As enumOpcao) As Boolean

    On Error GoTo ErroGeradorTexto
    
    psTextoGerado = ""
    If peOpcao = enumConsulta Then
        If Not GeraDrop(psTextoGerado, enumConsulta) Then GoTo ErroGeradorTexto
        If Not GeraConsulta(psTextoGerado) Then GoTo ErroGeradorTexto
        If Not GeraGrantExec(psTextoGerado, enumConsulta) Then GoTo ErroGeradorTexto
    End If
    If peOpcao = enumInclusao Then
        If Not GeraDrop(psTextoGerado, enumInclusao) Then GoTo ErroGeradorTexto
        If Not GeraInclusao(psTextoGerado) Then GoTo ErroGeradorTexto
        If Not GeraGrantExec(psTextoGerado, enumInclusao) Then GoTo ErroGeradorTexto
    End If
    If peOpcao = enumAlteracao Then
        If Not GeraDrop(psTextoGerado, enumAlteracao) Then GoTo ErroGeradorTexto
        'If Not GeraConsulta(psTextoGerado) Then GoTo ErroGeradorTexto
        If Not GeraGrantExec(psTextoGerado, enumAlteracao) Then GoTo ErroGeradorTexto
    End If
    If peOpcao = enumExclusao Then
        If Not GeraDrop(psTextoGerado, enumExclusao) Then GoTo ErroGeradorTexto
        'If Not GeraConsulta(psTextoGerado) Then GoTo ErroGeradorTexto
        If Not GeraGrantExec(psTextoGerado, enumExclusao) Then GoTo ErroGeradorTexto
    End If
    
    GeradorTexto = True
    Exit Function
ErroGeradorTexto:
    GeradorTexto = False
End Function

Private Function GeraComentario(ByRef psTextoGerado As String, _
                                ByVal peOpcao As enumOpcao) As Boolean
    Dim siLoop          As Integer
    Dim ssLinhaAux      As String
    
    On Error GoTo ErroGeraComentario
    If chkOpcaoComentario.Value = vbChecked Then
        psTextoGerado = psTextoGerado & "       " & psTraco & vbCrLf
        psTextoGerado = psTextoGerado & "       --                    " & psGerador & vbCrLf
        psTextoGerado = psTextoGerado & "       " & psTraco & vbCrLf
        psTextoGerado = psTextoGerado & "       -- Empresa..........: " & psEmpresa & vbCrLf
        psTextoGerado = psTextoGerado & "       -- Gerado em .......: " & Format(Now(), "dd/mm/yyyy hh:mm:ss") & vbCrLf
        psTextoGerado = psTextoGerado & "       -- Usuário..........: " & NomeUsuario & vbCrLf
        psTextoGerado = psTextoGerado & "       -- Máquina..........: " & NomeMaquina & vbCrLf
        psTextoGerado = psTextoGerado & "       -- Função...........: "
        If peOpcao = enumConsulta Then psTextoGerado = psTextoGerado & "Consulta da tabela "
        If peOpcao = enumInclusao Then psTextoGerado = psTextoGerado & "Inclusão na tabela "
        If peOpcao = enumAlteracao Then psTextoGerado = psTextoGerado & "Alteração na tabela "
        If peOpcao = enumExclusao Then psTextoGerado = psTextoGerado & "Exclusão da tabela "
        psTextoGerado = psTextoGerado & UCase(RTrim(txtDatabase.Text)) & "." & UCase(RTrim(Mid(cboTabela.BoundText, InStr(1, cboTabela.BoundText, "/") + 1, Len(cboTabela.BoundText)))) & vbCrLf
        If peOpcao = enumConsulta And paNomeTabelas(0) <> "" Then
            For siLoop = 0 To piQtdTabelas
                If paNomeTabelas(siLoop) <> "" Then
                    If siLoop = 0 Then
                        ssLinhaAux = "       -- Tab.Relacionadas.: "
                    Else
                        ssLinhaAux = "       --                    "
                    End If
                    ssLinhaAux = ssLinhaAux & UCase(RTrim(txtDatabase.Text)) & "." & UCase(paNomeTabelas(siLoop)) & vbCrLf
                    psTextoGerado = psTextoGerado & ssLinhaAux
                End If
            Next siLoop
        End If
        psTextoGerado = psTextoGerado & "       --                    "
        psTextoGerado = psTextoGerado & "       " & psTraco & vbCrLf
    End If
    GeraComentario = True
    Exit Function
ErroGeraComentario:
    GeraComentario = False
End Function

Private Function GeraDrop(ByRef psTextoGerado As String, _
                          ByVal peOpcao As enumOpcao) As Boolean
    Dim ssNomeProc          As String
    
    On Error GoTo ErroGeraDrop
    
    If peOpcao = enumConsulta Then ssNomeProc = txtNomeConsulta
    If peOpcao = enumInclusao Then ssNomeProc = txtNomeInclusao
    If peOpcao = enumAlteracao Then ssNomeProc = txtNomeAlteracao
    If peOpcao = enumExclusao Then ssNomeProc = txtNomeExclusao
    If chkOpcaoDrop.Value = vbChecked Then
        psTextoGerado = psTextoGerado & vbCrLf
        psTextoGerado = psTextoGerado & "IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id(N'" & ssNomeProc & "') " & vbCrLf
        psTextoGerado = psTextoGerado & "   AND OBJECTPROPERTY(id, N'IsProcedure') = 1)" & vbCrLf
        psTextoGerado = psTextoGerado & "   DROP PROCEDURE   " & ssNomeProc & vbCrLf
        psTextoGerado = psTextoGerado & "   GO      " & vbCrLf
    End If
    
    GeraDrop = True
    Exit Function
ErroGeraDrop:
    GeraDrop = False
End Function

Private Function GeraConsulta(ByRef psTextoGerado As String) As Boolean
    Dim siLoop          As Integer
    Dim siLoop1         As Integer
    Dim ssLinhaAux      As String
    Dim siQtdColunas    As Integer
    
    On Error GoTo ErroGeraConsulta
    psTextoGerado = psTextoGerado & vbCrLf
    psTextoGerado = psTextoGerado & "CREATE PROCEDURE " & txtNomeConsulta & vbCrLf
    '-------------- gera comentário
    If Not GeraComentario(psTextoGerado, enumConsulta) Then GoTo ErroGeraConsulta
    '-------------- conta número de colunas
    siQtdColunas = 0
    For siLoop = 1 To lvwConsulta.ListItems.Count
        If lvwConsulta.ListItems.Item(siLoop).Checked Or _
           EhChavePrimaria(lvwConsulta.ListItems.Item(siLoop).SubItems(1)) Then
            siQtdColunas = siQtdColunas + 1
        End If
    Next siLoop
    '-------------- declara parâmetros
    If moRSetIndicePK.RecordCount > 0 Then
        moRSetCamposTabela.MoveFirst
        siLoop = 1
        While Not moRSetCamposTabela.EOF And siLoop <= moRSetIndicePK.RecordCount
            If EhChavePrimaria(moRSetCamposTabela!COLUMN_NAME) Then
                ssLinhaAux = "       @" & RTrim(moRSetCamposTabela!COLUMN_NAME)
                ssLinhaAux = CompletaComBrancos(ssLinhaAux, 30)
                If Right(RTrim(moRSetCamposTabela!TYPE_NAME), 8) = "identity" Then
                    ssLinhaAux = ssLinhaAux & Mid(moRSetCamposTabela!TYPE_NAME, 1, Len(RTrim(moRSetCamposTabela!TYPE_NAME)) - 8)
                Else
                    ssLinhaAux = ssLinhaAux & moRSetCamposTabela!TYPE_NAME
                End If
                ssLinhaAux = RTrim(ssLinhaAux)
                If InStr(1, moRSetCamposTabela!TYPE_NAME, "char") > 0 Then ssLinhaAux = ssLinhaAux & "(" & RTrim(moRSetCamposTabela!LENGTH) & ")"
                ssLinhaAux = ssLinhaAux & IIf(siLoop < moRSetIndicePK.RecordCount, ",", "")
                psTextoGerado = psTextoGerado & ssLinhaAux & vbCrLf
                siLoop = siLoop + 1
            End If
            moRSetCamposTabela.MoveNext
        Wend
    End If
    '-------------- declarativas iniciais
    psTextoGerado = psTextoGerado & "AS BEGIN" & vbCrLf
    If moRSetIndicePK.RecordCount > 0 Or _
       siQtdColunas > 0 Then
        '-------------- monta SELECT
        psTextoGerado = psTextoGerado & "       SELECT " & vbCrLf
        If moRSetIndicePK.RecordCount > 0 Then
            moRSetCamposTabela.MoveFirst
            siLoop = 1
            While Not moRSetCamposTabela.EOF And siLoop <= moRSetIndicePK.RecordCount
                If EhChavePrimaria(moRSetCamposTabela!COLUMN_NAME) Then
                    psTextoGerado = psTextoGerado & "              T1." & RTrim(moRSetCamposTabela!COLUMN_NAME) & IIf(siLoop < moRSetIndicePK.RecordCount Or siQtdColunas > 0 Or paNomeTabelas(0) <> "", ",", "") & vbCrLf
                    siLoop = siLoop + 1
                End If
                moRSetCamposTabela.MoveNext
            Wend
        End If
        For siLoop = 1 To lvwConsulta.ListItems.Count
            If lvwConsulta.ListItems.Item(siLoop).Checked Then
                psTextoGerado = psTextoGerado & "              T1." & RTrim(lvwConsulta.ListItems.Item(siLoop).SubItems(1)) & IIf(siLoop < siQtdColunas Or paNomeTabelas(0) <> "", ",", "") & vbCrLf
            End If
        Next siLoop
        '-------------- se existe relacionamento
        If paNomeTabelas(0) <> "" Then
            For siLoop = 0 To piQtdTabelas
                If paNomeTabelas(siLoop) <> "" Then
                    For siLoop1 = 0 To piQtdCpoTabRelac
                        If paCamposTabRelac(siLoop, siLoop1) <> "" And _
                           paCamposTabRelacChk(siLoop, siLoop1) Then
                            ssLinhaAux = "              T" & siLoop + 2 & "." & paCamposTabRelac(siLoop, siLoop1)
                            ssLinhaAux = CompletaComBrancos(ssLinhaAux, 50)
                            ssLinhaAux = ssLinhaAux & " AS " & paCamposTabRelacAli(siLoop, siLoop1)
                            ssLinhaAux = ssLinhaAux & "," & vbCrLf
                            psTextoGerado = psTextoGerado & ssLinhaAux
                        End If
                    Next siLoop1
                End If
            Next siLoop
            '-------------- retira último ","
            psTextoGerado = Mid(psTextoGerado, 1, Len(psTextoGerado) - 3) & vbCrLf
        End If
        '-------------- monta FROM
        If paNomeTabelas(0) <> "" Then
            '---------- nome da tabela principal
            psTextoGerado = psTextoGerado & "       FROM   " & String(piProxTabela, "(") & vbCrLf
            ssLinhaAux = "              " & RTrim(Mid(cboTabela.BoundText, InStr(1, cboTabela.BoundText, "/") + 1, Len(cboTabela.BoundText))) & " T1"
            psTextoGerado = psTextoGerado & ssLinhaAux & vbCrLf
            '---------- nomes/campos das tabelas relacionadas
            For siLoop = 0 To piQtdTabelas
                If paNomeTabelas(siLoop) <> "" Then
                    '---------- nome da tabela relacionada
                    ssLinhaAux = ""
                    ssLinhaAux = CompletaComBrancos(ssLinhaAux, 17)
                    Select Case paTipoRelac(siLoop)
                        Case enumInnerJoin
                            ssLinhaAux = ssLinhaAux & "INNER JOIN       "
                        Case enumLeftJoin
                            ssLinhaAux = ssLinhaAux & "LEFT OUTER JOIN  "
                        Case enumRightJoin
                            ssLinhaAux = ssLinhaAux & "RIGHT OUTER JOIN "
                        Case enumFullJoin
                            ssLinhaAux = ssLinhaAux & "FULL OUTER JOIN  "
                    End Select
                    ssLinhaAux = ssLinhaAux & paNomeTabelas(siLoop) & " T" & siLoop + 2 & " ON "
                    psTextoGerado = psTextoGerado & ssLinhaAux & vbCrLf
                    '---------- campos da tabela relacionada
                    For siLoop1 = 0 To piQtdCpoRelac
                        ssLinhaAux = ""
                        ssLinhaAux = CompletaComBrancos(ssLinhaAux, 20)
                        If paCamposFonte(siLoop, siLoop1) <> "" Then
                            ssLinhaAux = ssLinhaAux & "T1." & paCamposFonte(siLoop, siLoop1)
                            ssLinhaAux = CompletaComBrancos(ssLinhaAux, 50)
                            ssLinhaAux = ssLinhaAux & " =  " & "T" & siLoop + 2 & "." & paCamposRelac(siLoop, siLoop1) & vbCrLf
                            psTextoGerado = psTextoGerado & ssLinhaAux
                        End If
                    Next siLoop1
                    '---------- coloca ")" na última posição
                    psTextoGerado = Mid(psTextoGerado, 1, Len(psTextoGerado) - 2) & ")" & vbCrLf
                End If
            Next siLoop
        Else
            psTextoGerado = psTextoGerado & "       FROM   " & vbCrLf
            psTextoGerado = psTextoGerado & "              " & Mid(cboTabela.BoundText, InStr(1, cboTabela.BoundText, "/") + 1, Len(cboTabela.BoundText)) & vbCrLf
        End If
        '-------------- monta WHERE
        If moRSetIndicePK.RecordCount > 0 Then
            psTextoGerado = psTextoGerado & "       WHERE  " & vbCrLf
            moRSetIndicePK.MoveFirst
            siLoop = 1
            While Not moRSetIndicePK.EOF
                ssLinhaAux = "              T1." & RTrim(moRSetIndicePK!COLUMN_NAME)
                ssLinhaAux = CompletaComBrancos(ssLinhaAux, 50)
                ssLinhaAux = ssLinhaAux & " =  @" & RTrim(moRSetIndicePK!COLUMN_NAME)
                ssLinhaAux = CompletaComBrancos(ssLinhaAux, 75)
                ssLinhaAux = ssLinhaAux & IIf(siLoop < moRSetIndicePK.RecordCount, " AND ", "")
                psTextoGerado = psTextoGerado & ssLinhaAux & vbCrLf
                
                moRSetIndicePK.MoveNext
                siLoop = siLoop + 1
            Wend
        End If
    End If
    '-------------- declarativas finais
    psTextoGerado = psTextoGerado & "END     " & vbCrLf
    psTextoGerado = psTextoGerado & "GO      " & vbCrLf
    
    GeraConsulta = True
    Exit Function
ErroGeraConsulta:
    GeraConsulta = False
End Function

Private Function GeraInclusao(ByRef psTextoGerado As String) As Boolean
    Dim siLoop          As Integer
    Dim ssLinhaAux      As String
    Dim ssLinhaAux1     As String
    Dim ssLinhaAux2     As String
    
    On Error GoTo ErroGeraInclusao
    psTextoGerado = psTextoGerado & vbCrLf
    psTextoGerado = psTextoGerado & "CREATE PROCEDURE " & txtNomeInclusao & vbCrLf
    '-------------- gera comentário
    If Not GeraComentario(psTextoGerado, enumInclusao) Then GoTo ErroGeraInclusao
    '-------------- declara parâmetros
    If moRSetIndicePK.RecordCount > 0 Then
        ssLinhaAux1 = ""
        ssLinhaAux2 = ""
        psTextoGerado = psTextoGerado & vbCrLf & "       -- Chave Primária (PK)" & vbCrLf
        moRSetCamposTabela.MoveFirst
        siLoop = 1
        While Not moRSetCamposTabela.EOF And siLoop <= moRSetIndicePK.RecordCount
            If EhChavePrimaria(moRSetCamposTabela!COLUMN_NAME) Then
                ssLinhaAux = "       @" & RTrim(moRSetCamposTabela!COLUMN_NAME)
                ssLinhaAux = CompletaComBrancos(ssLinhaAux, 30)
                If Right(RTrim(moRSetCamposTabela!TYPE_NAME), 8) = "identity" Then
                    ssLinhaAux = ssLinhaAux & Mid(moRSetCamposTabela!TYPE_NAME, 1, Len(RTrim(moRSetCamposTabela!TYPE_NAME)) - 8)
                Else
                    ssLinhaAux = ssLinhaAux & moRSetCamposTabela!TYPE_NAME
                End If
                ssLinhaAux = RTrim(ssLinhaAux)
                If InStr(1, moRSetCamposTabela!TYPE_NAME, "char") > 0 Then ssLinhaAux = ssLinhaAux & "(" & RTrim(moRSetCamposTabela!LENGTH) & ")"
                ssLinhaAux = ssLinhaAux & ","
                ssLinhaAux1 = ssLinhaAux1 & IIf(siLoop = 1, "             (", "              ") & RTrim(moRSetCamposTabela!COLUMN_NAME) & "," & vbCrLf
                ssLinhaAux2 = ssLinhaAux2 & IIf(siLoop = 1, "             (", "              ") & "@" & RTrim(moRSetCamposTabela!COLUMN_NAME) & "," & vbCrLf
                psTextoGerado = psTextoGerado & ssLinhaAux & vbCrLf
                siLoop = siLoop + 1
            End If
            moRSetCamposTabela.MoveNext
        Wend
    End If
    psTextoGerado = psTextoGerado & vbCrLf & "       -- Atributos" & vbCrLf
    With lvwInclusao
        While Not moRSetCamposTabela.EOF
            If .ListItems.Item("K" & moRSetCamposTabela!COLUMN_NAME).Checked Then
                If Not EhChavePrimaria(moRSetCamposTabela!COLUMN_NAME) Then
                    ssLinhaAux = "       @" & RTrim(moRSetCamposTabela!COLUMN_NAME)
                    ssLinhaAux = CompletaComBrancos(ssLinhaAux, 30)
                    If Right(RTrim(moRSetCamposTabela!TYPE_NAME), 8) = "identity" Then
                        ssLinhaAux = ssLinhaAux & Mid(moRSetCamposTabela!TYPE_NAME, 1, Len(RTrim(moRSetCamposTabela!TYPE_NAME)) - 8)
                    Else
                        ssLinhaAux = ssLinhaAux & moRSetCamposTabela!TYPE_NAME
                    End If
                    ssLinhaAux = RTrim(ssLinhaAux)
                    If InStr(1, moRSetCamposTabela!TYPE_NAME, "char") > 0 Then ssLinhaAux = ssLinhaAux & "(" & RTrim(moRSetCamposTabela!LENGTH) & ")"
                    ssLinhaAux = ssLinhaAux & ","
                    ssLinhaAux1 = ssLinhaAux1 & "              " & RTrim(moRSetCamposTabela!COLUMN_NAME) & "," & vbCrLf
                    ssLinhaAux2 = ssLinhaAux2 & "              " & "@" & RTrim(moRSetCamposTabela!COLUMN_NAME) & "," & vbCrLf
                    psTextoGerado = psTextoGerado & ssLinhaAux & vbCrLf
                End If
            End If
            moRSetCamposTabela.MoveNext
        Wend
    End With
    '-------------- retira último ","
    psTextoGerado = Mid(psTextoGerado, 1, Len(psTextoGerado) - 3) & vbCrLf
    
    '-------------- declarativas iniciais
    psTextoGerado = psTextoGerado & "AS BEGIN" & vbCrLf
    'psTextoGerado = psTextoGerado & "       BEGIN TRANSACTION" & vbCrLf & vbCrLf

    '-------------- monta INSERT
    psTextoGerado = psTextoGerado & "       INSERT " & vbCrLf
    psTextoGerado = psTextoGerado & "       INTO   " & vbCrLf
    psTextoGerado = psTextoGerado & "              " & Mid(cboTabela.BoundText, InStr(1, cboTabela.BoundText, "/") + 1, Len(cboTabela.BoundText)) & vbCrLf
    psTextoGerado = psTextoGerado & ssLinhaAux1
    '-------------- retira último "," e coloca ")"
    psTextoGerado = Mid(psTextoGerado, 1, Len(psTextoGerado) - 3) & ")" & vbCrLf
    
    psTextoGerado = psTextoGerado & "       VALUES " & vbCrLf
    psTextoGerado = psTextoGerado & ssLinhaAux2
    '-------------- retira último "," e coloca ")"
    psTextoGerado = Mid(psTextoGerado, 1, Len(psTextoGerado) - 3) & ")" & vbCrLf

    '-------------- declarativas finais
    'psTextoGerado = psTextoGerado & vbCrLf & "       COMMIT TRANSACTION" & vbCrLf
    psTextoGerado = psTextoGerado & "END     " & vbCrLf
    psTextoGerado = psTextoGerado & "GO      " & vbCrLf
    
    GeraInclusao = True
    Exit Function
ErroGeraInclusao:
    GeraInclusao = False
End Function
Private Function GeraGrantExec(ByRef psTextoGerado As String, _
                               ByVal peOpcao As enumOpcao) As Boolean
    Dim siLoop              As Integer
    Dim ssLinhaAux          As String
    
    On Error GoTo ErroGeraGrantExec
    If chkOpcaoGrant.Value = vbChecked Then
        psTextoGerado = psTextoGerado & vbCrLf
        For siLoop = 1 To lvwUsuarios.ListItems.Count
            If lvwUsuarios.ListItems.Item(siLoop).Checked = True Then
                ssLinhaAux = "GRANT EXECUTE ON "
                If peOpcao = enumConsulta Then ssLinhaAux = ssLinhaAux & txtNomeConsulta
                If peOpcao = enumInclusao Then ssLinhaAux = ssLinhaAux & txtNomeInclusao
                If peOpcao = enumAlteracao Then ssLinhaAux = ssLinhaAux & txtNomeAlteracao
                If peOpcao = enumExclusao Then ssLinhaAux = ssLinhaAux & txtNomeExclusao
                ssLinhaAux = CompletaComBrancos(ssLinhaAux, 50)
                ssLinhaAux = ssLinhaAux & " TO " & lvwUsuarios.ListItems.Item(siLoop).Text & vbCrLf
                psTextoGerado = psTextoGerado & ssLinhaAux
            End If
        Next siLoop
        psTextoGerado = psTextoGerado & "GO      " & vbCrLf
    End If
    GeraGrantExec = True
    Exit Function
ErroGeraGrantExec:
    GeraGrantExec = False
End Function

Private Function CompletaComBrancos(ByVal psLinha As String, ByVal piTamanho As Integer) As String
    If Len(psLinha) < piTamanho Then
        CompletaComBrancos = psLinha & Space(piTamanho - Len(psLinha))
    Else
        CompletaComBrancos = psLinha & " "
    End If
End Function
                
Private Function EhChavePrimaria(ByVal psNomeColuna As String) As Boolean
    If moRSetIndicePK.RecordCount > 0 Then
        moRSetIndicePK.MoveFirst
        moRSetIndicePK.Find "COLUMN_NAME='" & psNomeColuna & "'"
        EhChavePrimaria = Not moRSetIndicePK.EOF
    Else
        EhChavePrimaria = False
    End If
End Function

Private Sub MontaNomeProcedure(ByRef poCtrlTxt As TextBox, _
                               ByRef poCtrlTab As tabdlg.SSTab, _
                               ByVal peOpcao As enumOpcao, _
                               ByVal peStatus As CheckBoxConstants, _
                               ByVal psNomeTabela As String, _
                               ByVal pbSetarFoco As Boolean)
    If peStatus = vbUnchecked Then
        poCtrlTxt.Text = ""
        If peOpcao <> enumExclusao Then
            poCtrlTab.TabEnabled(peOpcao) = False
        End If
    Else
        If poCtrlTxt.Enabled Then
            poCtrlTxt.Text = "[dbo].["
            Select Case peOpcao
                Case enumConsulta
                    poCtrlTxt.Text = poCtrlTxt.Text & "PR_SEL_TAB_" & UCase(psNomeTabela)
                Case enumInclusao
                    poCtrlTxt.Text = poCtrlTxt.Text & "PR_INS_TAB_" & UCase(psNomeTabela)
                Case enumAlteracao
                    poCtrlTxt.Text = poCtrlTxt.Text & "PR_UPD_TAB_" & UCase(psNomeTabela)
                Case enumExclusao
                    poCtrlTxt.Text = poCtrlTxt.Text & "PR_DEL_TAB_" & UCase(psNomeTabela)
            End Select
            poCtrlTxt.Text = poCtrlTxt.Text & "]"
            If peOpcao <> enumExclusao Then
                poCtrlTab.TabEnabled(peOpcao) = True
            End If
            If pbSetarFoco Then poCtrlTxt.SetFocus
        Else
            poCtrlTxt.Text = ""
        End If
    End If
End Sub
