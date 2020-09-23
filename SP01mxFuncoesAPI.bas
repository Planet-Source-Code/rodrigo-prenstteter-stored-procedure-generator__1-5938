Attribute VB_Name = "SP01mxFuncoesAPI"
Option Explicit

Private Declare Function GetComputerName Lib "KERNEL32" Alias "GetComputerNameA" _
                (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
                (ByVal rsBuffer As String, rlSize As Long) As Long

Public Function NomeMaquina() As String
    Dim lSize As Long, sBuffer As String
    
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    
    If GetComputerName(sBuffer, lSize) Then
        NomeMaquina = Left$(sBuffer, lSize)
    Else
        NomeMaquina = ""
    End If
End Function

Public Function NomeUsuario() As String
    Dim lSize As Long, sBuffer As String
    
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    
    If GetUserName(sBuffer, lSize) Then
        NomeUsuario = Left$(sBuffer, lSize - 1)
    Else
        NomeUsuario = ""
    End If
End Function


