VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCreateDB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Database"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   Icon            =   "frmCreateDB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCreate 
      Caption         =   "C&reate"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   21
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   4680
      TabIndex        =   20
      Top             =   4080
      Width           =   975
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   5295
      Begin VB.Frame Frame1 
         Height          =   15
         Left            =   1200
         TabIndex        =   23
         Top             =   1800
         Width           =   3975
      End
      Begin VB.CommandButton cmdGetLog 
         Caption         =   "..."
         Height          =   300
         Left            =   4800
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Frame fraMaxSize 
         Caption         =   "Maximum file size"
         Height          =   1215
         Left            =   2400
         TabIndex        =   19
         Top             =   2040
         Width           =   2775
         Begin VB.TextBox txtRestr 
            BackColor       =   &H80000004&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1060
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            Height          =   285
            Left            =   2160
            TabIndex        =   17
            Top             =   720
            Width           =   495
         End
         Begin VB.OptionButton OptRest 
            Caption         =   "&Restrict filegrowth (MB):"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   720
            Width           =   2055
         End
         Begin VB.OptionButton OptUnRest 
            Caption         =   "&Unrestricted filegrowth"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Value           =   -1  'True
            Width           =   1935
         End
      End
      Begin VB.Frame fraGrowFile 
         Height          =   1215
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   2175
         Begin VB.TextBox txtPerc 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1060
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1560
            TabIndex        =   14
            Text            =   "10"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtMB 
            BackColor       =   &H80000004&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1060
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            TabIndex        =   12
            Top             =   360
            Width           =   495
         End
         Begin VB.OptionButton OptPercent 
            Caption         =   "By &percent:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton OptMB 
            Caption         =   "In &megabytes:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   1815
         End
         Begin VB.CheckBox cheGrow 
            Caption         =   "Automatically &grow file"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   0
            Value           =   1  'Checked
            Width           =   1935
         End
      End
      Begin VB.CommandButton CmdGetFile 
         Caption         =   "..."
         Height          =   300
         Left            =   4800
         TabIndex        =   9
         Top             =   1080
         Width           =   330
      End
      Begin VB.TextBox txtFilePath 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   4575
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtLogPath 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.Label Label1 
         Caption         =   "File properties"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblDbFile 
         Caption         =   "Database &File"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblName 
         Caption         =   "&Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblLogFile 
         Caption         =   "Transaction &Log File"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin MSComctlLib.TabStrip TabStripCre 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   6800
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&General"
            Key             =   "GENERAL"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Transaction Log"
            Key             =   "LOG"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCreateDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const COLOR_WINDOW = -2147483643
Const COLOR_GRAY = -2147483644
Const DATA_EXT = "_data.mdf"
Const LOG_EXT = "_log.ldf"

Dim oServer As SQLDMO.SQLServer
Dim strDataPath As String
Dim strFileGrowthMb$, strLogGrowthMb$
Dim strFileGrowthPerc$, strLogGrowthPerc$
Dim strFileRest$, strLogRest$
Dim bAutoFile As Boolean, bAutoLog As Boolean
Dim bGrowPercFile As Boolean, bGrowPercLog As Boolean
Dim bGrowRestFile As Boolean, bGrowRestLog As Boolean
Dim bGeneral As Boolean

Private Sub Form_Load()
   bGeneral = True
   bGrowPercFile = True
   bGrowPercLog = True
   strFileGrowthPerc = "10"
   strLogGrowthPerc = "10"
   bAutoFile = True
   bAutoLog = True
   ShowControls
   Set oServer = frmMain.Server
   strDataPath = oServer.Databases(oServer.Databases.Count).PrimaryFilePath

End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   Set oServer = Nothing

End Sub

Private Sub cheGrow_Click()

   If bGeneral Then
      bAutoFile = cheGrow
   Else
      bAutoLog = cheGrow
   End If
   EnableGrow

End Sub

Private Sub CmdGetFile_Click()
   Dim strPath$

   strPath = SaveDialog(Me.hwnd, txtName.Text & DATA_EXT, DATA_FILTER, , strDataPath, SAVE_FLAG)
   If Len(strPath) > 0 Then txtFilePath.Text = strPath

End Sub

Private Sub cmdGetLog_Click()
   Dim strPath$

   strPath = SaveDialog(Me.hwnd, txtName.Text & LOG_EXT, DATA_FILTER, , strDataPath, SAVE_FLAG)
   If Len(strPath) > 0 Then txtLogPath.Text = strPath
End Sub

Private Sub OptMB_Click()

   If bGeneral Then
      bGrowPercFile = False
   Else
      bGrowPercLog = False
   End If
   EnableGrow
   If txtMB.Enabled And Len(txtMB.Text) = 0 Then
      txtMB.Text = "1"
      If bGeneral Then
         strFileGrowthMb = txtMB.Text
      Else
         strLogGrowthMb = txtMB.Text
      End If
   End If


End Sub

Private Sub OptPercent_Click()

   If bGeneral Then
      bGrowPercFile = True
   Else
      bGrowPercLog = True
   End If
   EnableGrow
   If txtPerc.Enabled And Len(txtPerc.Text) = 0 Then
      txtPerc.Text = "1"
      If bGeneral Then
         strFileGrowthPerc = txtPerc.Text
      Else
         strLogGrowthPerc = txtPerc.Text
      End If
   End If

End Sub

Private Sub OptRest_Click()

   If bGeneral Then
      bGrowRestFile = True
   Else
      bGrowRestLog = True
   End If
   txtRestr.Enabled = OptRest
   If txtRestr.Enabled And Len(txtRestr.Text) = 0 Then
      txtRestr.Text = "1"
      If bGeneral Then
         strFileRest = txtRestr.Text
      Else
         strLogRest = txtRestr.Text
      End If
   End If
   txtRestr.BackColor = IIf(OptRest, COLOR_WINDOW, COLOR_GRAY)

End Sub

Private Sub OptUnRest_Click()

   If bGeneral Then
      bGrowRestFile = False
   Else
      bGrowRestLog = False
   End If
   txtRestr.Enabled = OptRest
   txtRestr.BackColor = IIf(OptRest, COLOR_WINDOW, COLOR_GRAY)

End Sub

Private Sub txtMB_Change()

   If bGeneral Then
      strFileGrowthMb = txtMB.Text
   Else
      strLogGrowthMb = txtMB.Text
   End If

End Sub

Private Sub txtPerc_Change()

   If bGeneral Then
      strFileGrowthPerc = txtPerc.Text
   Else
      strLogGrowthPerc = txtPerc.Text
   End If

End Sub

Private Sub txtRestr_Change()

   If bGeneral Then
      strFileRest = txtRestr.Text
   Else
      strLogRest = txtRestr.Text
   End If

End Sub

Private Sub cmdCreate_Click()

   If CreateDataBase() Then
      MsgBox "Database '" & txtName.Text & "' successfuly created.", vbInformation
      Me.Visible = False
   End If
   Unload Me

End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub txtFilePath_Change()
   EnableCreate
End Sub

Private Sub txtLogPath_Change()
   EnableCreate
End Sub

Private Sub txtName_Change()
   On Error Resume Next

   txtFilePath.Text = strDataPath & txtName.Text & DATA_EXT
   txtLogPath.Text = strDataPath & txtName.Text & LOG_EXT
   EnableCreate
End Sub

Private Function CreateDataBase() As Boolean
   On Error GoTo CreateDataBase_E

   Dim oDatabase As SQLDMO.Database
   Dim oDBFileData As SQLDMO.DBFile
   Dim oLogFile As SQLDMO.LogFile

   Set oDatabase = New SQLDMO.Database
   Set oDBFileData = New SQLDMO.DBFile
   Set oLogFile = New SQLDMO.LogFile

   frmMain.stbStatus.Panels("SB_TEXT").Text = "Creating Database..."
   Me.MousePointer = vbHourglass
   DoEvents
   oDatabase.Name = txtName.Text
   With oDBFileData
      .Name = txtName.Text
      .PhysicalName = txtFilePath.Text
      .PrimaryFile = True
      If bAutoFile Then
         .FileGrowthType = IIf(bGrowPercFile, SQLDMOGrowth_Percent, SQLDMOGrowth_MB)
         .FileGrowth = Val(IIf(bGrowPercFile, strFileGrowthPerc, strFileGrowthMb))
      Else
         .FileGrowthType = SQLDMOGrowth_Invalid
      End If
      If bGrowRestFile Then
         .MaximumSize = Val(strFileRest)
      End If
   End With
   oDatabase.FileGroups("PRIMARY").DBFiles.Add oDBFileData
   With oLogFile
      .Name = txtLogPath.Text
      .PhysicalName = txtLogPath.Text
      If bAutoLog Then
         .FileGrowthType = IIf(bGrowPercLog, SQLDMOGrowth_Percent, SQLDMOGrowth_MB)
         .FileGrowth = Val(IIf(bGrowPercLog, strFileGrowthPerc, strFileGrowthMb))
      Else
         .FileGrowthType = SQLDMOGrowth_Invalid
      End If
      If bGrowRestLog Then
         .MaximumSize = Val(strLogRest)
      End If
   End With

   oDatabase.TransactionLog.LogFiles.Add oLogFile
   DoEvents
   oServer.Databases.Add oDatabase
   '   oServer.Databases.Refresh
   frmMain.cboDatabase.AddItem txtName.Text
   frmMain.cboDatabase = txtName.Text
   CreateDataBase = True
   '   MsgBox "Database created.", vbInformation

CreateDataBase_X:
   frmMain.stbStatus.Panels("SB_TEXT").Text = "Connected"
   Me.MousePointer = vbDefault
   Set oDatabase = Nothing
   Set oDBFileData = Nothing
   Set oLogFile = Nothing

   Exit Function

CreateDataBase_E:
   MsgBox Err.Description, vbCritical
   Resume CreateDataBase_X

End Function

Private Sub EnableCreate()

   cmdCreate.Enabled = _
         (Len(txtName.Text) > 0) And _
         (Len(txtFilePath.Text) > 0) And _
         (Len(txtLogPath.Text) > 0)

End Sub

Private Sub TabStripCre_Click()

   bGeneral = TabStripCre.SelectedItem.Index = 1
   ShowControls

End Sub

Private Sub ShowControls()

   lblName.Visible = bGeneral
   txtName.Visible = bGeneral
   lblDbFile.Visible = bGeneral
   txtFilePath.Visible = bGeneral
   CmdGetFile.Visible = bGeneral
   lblLogFile.Visible = Not bGeneral
   txtLogPath.Visible = Not bGeneral
   lblLogFile.Visible = Not bGeneral
   txtLogPath.Visible = Not bGeneral
   cmdGetLog.Visible = Not bGeneral
   If bGeneral Then
      cheGrow.Value = Abs(bAutoFile)
      OptMB.Value = Not bGrowPercFile
      OptPercent.Value = bGrowPercFile
      OptUnRest.Value = Not bGrowRestFile
      OptRest.Value = bGrowRestFile
      txtMB.Text = strFileGrowthMb
      txtPerc.Text = strFileGrowthPerc
      txtRestr.Text = strFileRest
   Else
      cheGrow.Value = Abs(bAutoLog)
      OptMB.Value = Not bGrowPercLog
      OptPercent.Value = bGrowPercLog
      OptUnRest.Value = Not bGrowRestLog
      OptRest.Value = bGrowRestLog
      txtMB.Text = strLogGrowthMb
      txtPerc.Text = strLogGrowthPerc
      txtRestr.Text = strLogRest
   End If
   EnableGrow
   txtRestr.Enabled = OptRest
   txtRestr.BackColor = IIf(OptRest, COLOR_WINDOW, COLOR_GRAY)

End Sub

Private Sub EnableGrow()

   OptMB.Enabled = cheGrow
   OptPercent.Enabled = cheGrow
   txtMB.Enabled = OptMB And cheGrow
   txtPerc.Enabled = Not txtMB.Enabled And cheGrow
   txtMB.BackColor = IIf(OptMB And cheGrow, COLOR_WINDOW, COLOR_GRAY)
   txtPerc.BackColor = IIf(OptPercent And cheGrow, COLOR_WINDOW, COLOR_GRAY)

End Sub
