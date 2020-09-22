VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExecSQL 
   Caption         =   "Executing Script"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4920
   Icon            =   "frmSQLexec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   4920
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkErrMsg 
      Alignment       =   1  'Right Justify
      Caption         =   "&Error message:"
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   3840
      Value           =   1  'Checked
      Width           =   1395
   End
   Begin VB.CheckBox chkError 
      Alignment       =   1  'Right Justify
      Caption         =   "Pause on &error:"
      Height          =   375
      Left            =   80
      TabIndex        =   9
      Top             =   3840
      Value           =   1  'Checked
      Width           =   1440
   End
   Begin MSComctlLib.StatusBar stbStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   4215
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2963
            MinWidth        =   2470
            Text            =   "Current command: 1/1"
            TextSave        =   "Current command: 1/1"
            Key             =   "CURRENT"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3916
            Key             =   "EXECUTED"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1235
            MinWidth        =   1235
            Text            =   "Paused"
            TextSave        =   "Paused"
            Key             =   "PAUSE"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraLine 
      Height          =   15
      Left            =   0
      TabIndex        =   7
      Top             =   3360
      Width           =   4935
   End
   Begin VB.TextBox txtSrch 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   3520
      Width           =   3615
   End
   Begin VB.CommandButton cmdStep 
      Height          =   300
      Left            =   3480
      Picture         =   "frmSQLexec.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Go"
      Top             =   3000
      Width           =   330
   End
   Begin VB.CommandButton cmdStop 
      Height          =   300
      Left            =   4560
      Picture         =   "frmSQLexec.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Stop"
      Top             =   3000
      Width           =   330
   End
   Begin VB.CommandButton cmdGo 
      Height          =   300
      Left            =   3840
      Picture         =   "frmSQLexec.frx":0F56
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Go"
      Top             =   3000
      Width           =   330
   End
   Begin VB.CommandButton cmdPause 
      Height          =   300
      Left            =   4200
      Picture         =   "frmSQLexec.frx":14E0
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Pause"
      Top             =   3000
      Width           =   330
   End
   Begin VB.TextBox txtText 
      Height          =   2900
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label lblDbName 
      AutoSize        =   -1  'True
      Caption         =   "Database:"
      Height          =   195
      Left            =   135
      TabIndex        =   11
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label lblSrch 
      Caption         =   "Pause on &string:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3520
      Width           =   1215
   End
End
Attribute VB_Name = "frmExecSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" ( _
      lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Type POINTAPI
   x As Long
   y As Long
End Type
Private Type MINMAXINFO
   ptReserved As POINTAPI
   ptMaxSize As POINTAPI
   ptMaxPosition As POINTAPI
   ptMinTrackSize As POINTAPI
   ptMaxTrackSize As POINTAPI
End Type
Private Const WM_GETMINMAXINFO = &H24

Implements ISubclass
Private m_emr As EMsgResponse

Dim lngGoIndentLeft&, lngStepIndentLeft&, lngPauseIndentLeft&
Dim lngStopIndentLeft&, lngLblPauseIndentLeft&
Dim lngIndentTop&, lngTextIndentHeight&
Dim lngLineIndentTop&, lngSrchIndentTop&, lngChkIndentTop&
Dim Server As SQLDMO.SQLServer
Dim bTextChange As Boolean
Dim bExecuting As Boolean
Dim bStopExecuting As Boolean
Dim bStep As Boolean
Dim bPauseString As Boolean
Dim bError As Boolean

Public m_sDbName As String

Public Property Let DatabaseName(ByVal strValue As String)
   m_sDbName = strValue
   lblDbName.Caption = "Database: " & m_sDbName
End Property

Private Sub Form_Load()

   AttachMessage Me, Me.hwnd, WM_GETMINMAXINFO
   lngGoIndentLeft = Me.ScaleWidth - Me.cmdGo.Left
   lngStepIndentLeft = Me.ScaleWidth - Me.cmdStep.Left
   lngPauseIndentLeft = Me.ScaleWidth - Me.cmdPause.Left
   lngStopIndentLeft = Me.ScaleWidth - Me.cmdStop.Left
   lngIndentTop = Me.ScaleHeight - Me.cmdGo.Top
   lngLineIndentTop = Me.ScaleHeight - Me.fraLine.Top
   lngSrchIndentTop = Me.ScaleHeight - Me.txtSrch.Top
   lngChkIndentTop = Me.ScaleHeight - Me.chkError.Top
   lngTextIndentHeight = Me.ScaleHeight - Me.txtText.Height
   Set Server = frmMain.Server
   
End Sub

Private Sub Form_Resize()
   Dim lngTop&, lngLineTop&, lngSrchTop&, lngChkTop&

   lngTop = Me.ScaleHeight - lngIndentTop
   lngLineTop = Me.ScaleHeight - lngLineIndentTop
   lngSrchTop = Me.ScaleHeight - lngSrchIndentTop
   lngChkTop = Me.ScaleHeight - lngChkIndentTop
   Me.txtText.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - lngTextIndentHeight
   Me.cmdGo.Move Me.ScaleWidth - lngGoIndentLeft, lngTop
   Me.cmdStep.Move Me.ScaleWidth - lngStepIndentLeft, lngTop
   Me.cmdPause.Move Me.ScaleWidth - lngPauseIndentLeft, lngTop
   Me.cmdStop.Move Me.ScaleWidth - lngStopIndentLeft, lngTop
   lblDbName.Move lblDbName.Left, lngTop
   fraLine.Move 0, lngLineTop, Me.ScaleWidth
   lblSrch.Move lblSrch.Left, lngSrchTop
   txtSrch.Move txtSrch.Left, lngSrchTop, Me.ScaleWidth - txtSrch.Left - 40
   chkError.Move chkError.Left, lngChkTop
   chkErrMsg.Move chkErrMsg.Left, lngChkTop
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   DetachMessage Me, Me.hwnd, WM_GETMINMAXINFO
End Sub

Private Sub cmdGo_Click()
   Me.txtText.SetFocus
   bExecuting = True
   stbStatus.Panels("PAUSE").Visible = False
   bError = False
End Sub

Private Sub cmdPause_Click()
   Me.txtText.SetFocus
   bExecuting = False
   stbStatus.Panels("PAUSE").Visible = True
End Sub

Private Sub cmdStep_Click()
   Me.txtText.SetFocus
   bExecuting = False
   bStep = True
   stbStatus.Panels("PAUSE").Visible = False
   bError = False
End Sub

Private Sub cmdStop_Click()
   Me.txtText.SetFocus
   bStopExecuting = True
End Sub

Private Sub txtText_Change()
   If Not bTextChange Then
      bExecuting = False
      stbStatus.Panels("EXECUTED").Text = vbNullString
      stbStatus.Panels("PAUSE").Visible = True
   End If
End Sub

Private Sub txtSrch_Change()
   bPauseString = Len(txtSrch) > 0
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer6.EMsgResponse)
   m_emr = RHS
End Property

Private Property Get ISubclass_MsgResponse() As SSubTimer6.EMsgResponse
   Debug.Print CurrentMessage
   m_emr = emrConsume
   ISubclass_MsgResponse = m_emr
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Dim mmiT As MINMAXINFO

   ' Copy parameter to local variable for processing
   CopyMemory mmiT, ByVal lParam, LenB(mmiT)

   ' Minimium width and height for sizing
   mmiT.ptMinTrackSize.x = 330
   mmiT.ptMinTrackSize.y = 175

   ' Copy modified results back to parameter
   CopyMemory ByVal lParam, mmiT, LenB(mmiT)

End Function

Public Function RunScript(ByVal colCommands As Collection, ByVal ScriptPath As String) As Boolean
   On Error GoTo RunScript_E
   Dim i&, lngCnt&
   Dim bExecuted As Boolean
   
   bExecuted = True
   i = InStrRev(ScriptPath, "\", , vbTextCompare)
   If i > 0 Then Me.Caption = "Executing script " & Mid$(ScriptPath, i + 1)
   i = 0
   lngCnt = colCommands.Count
   If lngCnt = 0 Then
      i = 1
      Me.Caption = "Executing SQL Command"
   End If
   
   Do While Not bStopExecuting
      DoEvents
      If bExecuted And Not bError And lngCnt >= i + 1 Then
         i = i + 1
         bTextChange = True
         txtText.Text = colCommands(i)
         bTextChange = False
         stbStatus.Panels("CURRENT").Text = "Current command " & i & "/" & lngCnt
         If bPauseString Then
            If InStr(1, txtText.Text, txtSrch.Text, vbTextCompare) > 0 Then
               cmdPause_Click
            End If
         End If
         bExecuted = False
      End If
      If bExecuting Or bStep Then
         If Len(txtText.Text) > 0 Then
            DoEvents
            Server.Databases(m_sDbName).ExecuteImmediate txtText.Text, SQLDMOExec_ContinueOnError
            stbStatus.Panels("EXECUTED").Text = "Command " & i & " executed"
            bError = False
RunScript_Executed:
            If lngCnt = 0 Then bExecuting = False
            If bStep Then
               stbStatus.Panels("PAUSE").Visible = True ' .Text = "Paused"
            End If
            bStep = False
         End If
         bExecuted = True
         If (i >= lngCnt) And (lngCnt > 0) Then bStopExecuting = True
      End If
   Loop
   RunScript = True

RunScript_X:
   Unload Me
   Exit Function

RunScript_E:
   If Err.Description = "[SQL-DMO]" Then Resume Next
   If chkErrMsg Then MsgBox Err.Description
   stbStatus.Panels("EXECUTED").Text = "Command " & i & " failed"
   If chkError Or lngCnt = 0 Then
      bError = True
      cmdPause_Click
      Resume RunScript_Executed
   End If
   Select Case Err
      Case -2147199229
         Resume RunScript_Executed
      Case Else
         MsgBox Err.Description, vbCritical
         Resume RunScript_X
   End Select

End Function
