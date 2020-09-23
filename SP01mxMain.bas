Attribute VB_Name = "SP01mxMain"
Option Explicit
Public Const psEmpresa = "Nome Empresa"
Public Const psGerador = "Gerador de Stored Procedures V.0.8 (14/01/2000)"
Public Const psTraco = "-------------------------------------------------------------------------------------------"
Public Const piQtdTabelas = 10
Public Const piQtdCpoRelac = 7
Public Const piQtdCpoTabRelac = 50

Public Enum enumTipoRelac
    enumInnerJoin = 1
    enumLeftJoin
    enumRightJoin
    enumFullJoin
End Enum

Public paNomeTabelas(piQtdTabelas)                          As String
Public paTipoRelac(piQtdTabelas)                            As enumTipoRelac
Public paCamposFonte(piQtdTabelas, piQtdCpoRelac)           As String
Public paCamposRelac(piQtdTabelas, piQtdCpoRelac)           As String
Public paCamposTabRelac(piQtdTabelas, piQtdCpoTabRelac)     As String
Public paCamposTabRelacChk(piQtdTabelas, piQtdCpoTabRelac)  As Boolean
Public paCamposTabRelacAli(piQtdTabelas, piQtdCpoTabRelac)  As String

Public piProxTabela                                         As Integer
Public piProxCpoRelac(piQtdTabelas)                         As Integer
Public piProxCpoTabRelac(piQtdTabelas)                      As Integer
    
Public Sub Main()
    SP01wxStoredProced.Show vbModal
End Sub
