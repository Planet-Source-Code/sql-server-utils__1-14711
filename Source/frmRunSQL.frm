VERSION 5.00
Begin VB.Form frmRunSQL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SQL Command"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   Icon            =   "frmRunSQL.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1400
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   2775
      Begin VB.ComboBox cboDatabase 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   920
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "User:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Server:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblUser 
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblServer 
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "D&atabase:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   3020
      TabIndex        =   4
      Top             =   1320
      Width           =   1245
   End
   Begin VB.CommandButton cmdRunScript 
      Caption         =   "&Run Command"
      Height          =   375
      Left            =   3020
      TabIndex        =   3
      Top             =   840
      Width           =   1245
   End
   Begin VB.CommandButton cmdSQL 
      Caption         =   "..."
      Height          =   300
      Left            =   3960
      TabIndex        =   1
      Top             =   360
      Width           =   330
   End
   Begin VB.TextBox txtSQLPath 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label4 
      Caption         =   "&SQL Script Path"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmRunSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim frmExec As frmExecSQL
Dim strCurDir$

Private Sub Form_Load()

   lblUser.Caption = frmMain.Username
   lblServer.Caption = frmMain.Server.Name
   FillDatabases
   strCurDir = GetRegString(REG_SCRIPT_DIR)
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next

   If Len(strCurDir) > 0 Then SetRegString REG_SCRIPT_DIR, strCurDir
   Unload frmExec
   frmMain.ZOrder 0

End Sub

Private Sub cmdSQL_Click()
   Dim strPath$

   strPath = OpenDialog(Me.hwnd, txtSQLPath.Text, SQL_FILTER, , strCurDir, OPEN_FLAG)
   If Len(strPath) > 0 Then
      strCurDir = GetFolder(strPath)
      txtSQLPath.Text = strPath
   End If
   
End Sub

Private Sub cmdRunScript_Click()
   Dim colCommands As Collection

   Set colCommands = New Collection
   If GetCommands(txtSQLPath.Text, colCommands) Then
      Set frmExec = New frmExecSQL
      With frmExec
         .DatabaseName = cboDatabase
         .Show
         .RunScript colCommands, txtSQLPath.Text
      End With
   End If
   Set colCommands = Nothing

End Sub

Private Sub txtSQLPath_Change()
   cmdRunScript.Caption = IIf(Len(txtSQLPath) > 0, "&Run Script", "&Run Command")
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub FillDatabases()
   Dim i&
   Dim oCbo As ComboBox

   cboDatabase.Clear
   Set oCbo = frmMain.cboDatabase
   For i = 0 To oCbo.ListCount - 1
      cboDatabase.AddItem oCbo.List(i)
   Next i
   cboDatabase = cboDatabase.List(oCbo.ListIndex)
   Set oCbo = Nothing

End Sub

Private Function GetCommands(ByVal ScriptPath As String, colCommands As Collection) As Boolean
   On Error GoTo GetCommands_E
   Dim iFile%
   Dim strCmd$, strLine$

   If Len(ScriptPath) > 0 Then
      Me.MousePointer = vbHourglass
      iFile = FreeFile
      Open ScriptPath For Input As #iFile
      Do While Not (EOF(iFile))
         Line Input #iFile, strLine
         strLine = Trim$(strLine)
         If Len(strLine) > 0 Then
            If StrComp(Left(strLine, 2), "GO", vbTextCompare) <> 0 Then
               strCmd = strCmd + strLine + vbCrLf
            Else
               colCommands.Add Trim$(strCmd)
               strCmd = vbNullString
            End If
         End If
      Loop
   End If
   GetCommands = True

GetCommands_X:
   On Error Resume Next
   Close #iFile
   Me.MousePointer = vbDefault
   Exit Function

GetCommands_E:
   MsgBox Err.Description, vbCritical

End Function

