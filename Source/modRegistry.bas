Attribute VB_Name = "modRegistry"
Option Explicit

Public Const REG_ROOTKEY = HKEY_CURRENT_USER
Public Const REG_SECTIONKEY = "SOFTWARE\SQLsdu"
Public Const REG_SCRIPT_DIR = "Script_folder"
Public Const REG_USER = "User"

Public Function GetRegString(ByVal KeyValue As String) As String
   On Error Resume Next

   Dim cReg As New cRegistry
   With cReg
      .ClassKey = REG_ROOTKEY
      .SectionKey = REG_SECTIONKEY
      .ValueKey = KeyValue
      GetRegString = .Value
   End With
   Set cReg = Nothing

End Function

Public Sub SetRegString(ByVal KeyValue As String, ByVal vValue)
   On Error Resume Next

   Dim cReg As New cRegistry
   With cReg
      .ClassKey = REG_ROOTKEY
      .SectionKey = REG_SECTIONKEY
      .ValueType = REG_SZ
      .ValueKey = KeyValue
      .Value = vValue
   End With
   Set cReg = Nothing

End Sub

