VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "SQL Server Database Utility"
   ClientHeight    =   2040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1440
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2775
      Begin VB.ComboBox cboDatabase 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   920
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "D&atabase:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblServer 
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblUser 
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Server:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "User:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdBackup 
      Caption         =   "&Backup Database"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect to Server"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdRemoveDB 
      Caption         =   "D&elete Database"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin MSComctlLib.StatusBar stbStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1785
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8731
            Key             =   "SB_TEXT"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            Object.Width           =   1589
            MinWidth        =   1589
            TextSave        =   "25.1.01"
            Key             =   "SB_DATE"
            Object.ToolTipText     =   "Current Date"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdRunScript 
      Caption         =   "&SQL Commands"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdCreDB 
      Caption         =   "Create &Database"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents oSvr As SQLDMO.SQLServer
Attribute oSvr.VB_VarHelpID = -1
Dim WithEvents conn As ADODB.Connection
Attribute conn.VB_VarHelpID = -1

Dim strUser As String
Dim strPassword As String
Dim strServer As String
Dim strDataPath As String
Dim bConnected As Boolean

Public Property Get Server() As SQLDMO.SQLServer
   Set Server = oSvr
End Property

Public Property Get Username() As String
   Username = strUser
End Property

Public Property Get Password() As String
   Password = strUser
End Property

Private Sub conn_ConnectComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
   stbStatus.Panels("SB_TEXT").Text = "Connected"
End Sub

Private Sub conn_WillConnect(ConnectionString As String, UserID As String, Password As String, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
   DoEvents
   stbStatus.Panels("SB_TEXT").Text = "Connecting..."
End Sub

Private Sub Form_Load()

   RemoveMenuItem Me.hwnd, SC_SIZE
   Set oSvr = New SQLDMO.SQLServer
   stbStatus.Panels("SB_TEXT").Text = "Not Connected"
   strUser = GetRegString(REG_USER)
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Dim frm As Form

   If Len(strUser) > 0 Then SetRegString REG_USER, strUser
   If bConnected Then oSvr.DisConnect
   Set oSvr = Nothing
   If conn.State = adStateOpen Then conn.Close
   Set conn = Nothing
   For Each frm In Forms
      If frm.Name <> Me.Name Then Unload frm
      Set frm = Nothing
   Next frm

End Sub

Private Sub cmdConnect_Click()
   ConnectMe
End Sub

Private Sub cmdCreDB_Click()
   frmCreateDB.Show , Me
End Sub

Private Sub cmdRemoveDB_Click()
   Dim strMsg$

   strMsg = "Are you sure you want to delete database '" & cboDatabase & "'?"
   If MsgBox(strMsg, vbYesNo + vbDefaultButton2 + vbQuestion) <> vbYes Then Exit Sub
   If RemoveDB(cboDatabase, True) Then
      cboDatabase.RemoveItem cboDatabase.ListIndex
      cboDatabase.ListIndex = 0
   End If

End Sub

Private Sub cmdRunScript_Click()
   Dim frmSQL As New frmRunSQL

   frmSQL.Show , Me

End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cboDatabase_Click()
   cmdRemoveDB.Enabled = cboDatabase <> "master"
End Sub

Private Sub oSvr_CommandSent(ByVal SQLCommand As String)
   'CommandSent event occurs when SQL-DMO submits a Transact-SQL command batch to the connected instance
   '   Dim sMsg As String
   '   sMsg = "CommandSent: " & vbCrLf & SQLCommand
   '   MsgBox sMsg, vbOKOnly, vbInformation

End Sub

Private Function oSvr_ConnectionBroken(ByVal Message As String) As Boolean
   'ConnectionBroken event occurs when a connected SQLServer object loses its connection
   Dim sMsg As String
   sMsg = "ConnectionBroken: " & vbCrLf & Message
   MsgBox sMsg, vbOKOnly, vbCritical
   stbStatus.Panels("SB_TEXT").Text = "Not Connected"

End Function

Private Function oSvr_QueryTimeout(ByVal Message As String) As Boolean
   'QueryTimeout event occurs when Microsoft® SQL Server™ cannot complete execution of a Transact-SQL command batch within a user-defined period of time
   Dim sMsg As String
   sMsg = "QueryTimeout: " & vbCrLf & Message
   MsgBox sMsg, vbOKOnly, vbCritical

End Function

Private Sub oSvr_RemoteLoginFailed(ByVal Severity As Long, ByVal MessageNumber As Long, ByVal MessageState As Long, ByVal Message As String)
   'RemoteLoginFailed event occurs when an instance of Microsoft® SQL Server™ attempts to connect to a remote server fails
   Dim sMsg As String
   sMsg = "RemoteLoginFailed: " & vbCrLf & _
         "Severity: " & Severity & vbCrLf & _
         "MessageNumber: " & MessageNumber & vbCrLf & _
         "MessageState: " & MessageState & vbCrLf & _
         "Message: " & Message

   MsgBox sMsg, vbOKOnly, vbCritical
   stbStatus.Panels("SB_TEXT").Text = "Not Connected"

End Sub

Private Sub oSvr_ServerMessage(ByVal Severity As Long, ByVal MessageNumber As Long, ByVal MessageState As Long, ByVal Message As String)
   'ServerMessage event occurs when a Microsoft® SQL Server™ success-with-information message is returned to the SQL-DMO application
   '   Dim sMsg As String
   '   sMsg = "ServerMessage: " & vbCrLf & _
       '         "Severity: " & Severity & vbCrLf & _
       '         "MessageNumber: " & MessageNumber & vbCrLf & _
       '         "MessageState: " & MessageState & vbCrLf & _
       '         "Message: " & Message
   '
   '   MsgBox sMsg, vbOKOnly, "SQLServer Object Event"
End Sub

Private Function ConnectMe()
   On Error GoTo ConnectMe_E

   Me.MousePointer = vbHourglass
   DoEvents
   If GetServerData() Then
      If ConnectToServer(strServer, strUser, strPassword) Then
         DoEvents
         Me.MousePointer = vbHourglass
         FillDatabases
         bConnected = True
         stbStatus.Panels("SB_TEXT").Text = "Retrieving Server Data..."
         strDataPath = oSvr.Databases(oSvr.Databases.Count).PrimaryFilePath
         Me.MousePointer = vbDefault
         cmdCreDB.Enabled = True
         cmdRemoveDB.Enabled = Len(cboDatabase) > 0
         '         cmdBackup.Enabled = cmdRemoveDB.Enabled
         cmdRunScript.Enabled = cmdRemoveDB.Enabled
         stbStatus.Panels("SB_TEXT").Text = "Connected"
         cboDatabase.SetFocus
      End If
   End If

ConnectMe_X:
   Me.MousePointer = vbDefault
   Exit Function

ConnectMe_E:
   stbStatus.Panels("SB_TEXT").Text = "Not Connected"
   MsgBox Err.Description, vbCritical
   Resume ConnectMe_X

End Function

Private Sub FillDatabases()
   Dim i&
   Dim db As SQLDMO.Database

   cboDatabase.Clear
   For Each db In oSvr.Databases
      cboDatabase.AddItem db.Name
   Next db
   Set db = Nothing
   cboDatabase.Enabled = True
   cboDatabase = cboDatabase.List(0)

End Sub

Private Function ConnectToServer( _
         ByVal SrvName As String, _
         ByVal User As String, _
         ByVal Password As String) As Boolean

   On Error GoTo ConnectToServer_E

   DoEvents
   With oSvr
      .DisConnect
      .LoginTimeout = 30
      stbStatus.Panels("SB_TEXT").Text = "Connecting..."
      DoEvents
      .Start True, SrvName, User, Password
   End With
   ConnectToServer = True

ConnectToServer_X:
   Exit Function

ConnectToServer_E:

   Select Case Err
         ' This error happens if service is already running
      Case -2147023840, 1056
         ' Use this to logon
         oSvr.Connect SrvName, User, Password
         Resume Next
         'Process errors in attempt to connect to target server
      Case -2147165949
         MsgBox "Your version of SQL-DMO is out of date for the " & _
               "server named " & SrvName & ". Update your SQL-DMO " & _
               "and Enterprise Manager to a more recent version in " & _
               "order to connect to the server.", vbCritical
      Case Else
         MsgBox Err.Description, vbCritical
   End Select
   ConnectToServer = False

End Function

Private Function RemoveDB(ByVal Name As String, Optional ByVal Msg As Boolean = False) As Boolean
   On Error GoTo RemoveDB_E

   stbStatus.Panels("SB_TEXT").Text = "Removing Database..."
   Me.MousePointer = vbHourglass
   DoEvents
   oSvr.Databases.Remove Name
   RemoveDB = True
   If Msg Then
      MsgBox "Database '" & Name & "' deleted.", vbInformation
   End If

RemoveDB_X:
   Me.MousePointer = vbDefault
   stbStatus.Panels("SB_TEXT").Text = vbNullString
   Exit Function

RemoveDB_E:
   MsgBox Err.Description, vbCritical
   Resume RemoveDB_X

End Function

Private Function GetServerData() As Boolean
   On Error GoTo GetServerData_E

   Dim bRet As Boolean

   DoEvents
   Set conn = New ADODB.Connection
   With conn
      .Provider = PROPVAL_DATAPROVIDER
      If Len(strUser) > 0 Then .Properties(PROPNAME_UID) = strUser
      .Properties(PROPNAME_SOURCE) = PROPVAL_SOURCE
      .Properties(PROPNAME_PERSEC) = True
      .Properties(PROPNAME_PROMPT) = adPromptAlways
      .Properties(PROPNAME_PROMPT) = adPromptComplete
      .Open
      bRet = .State = adStateOpen
      If bRet Then
         strUser = .Properties(PROPNAME_UID)
         If Len(strUser) = 0 Then strUser = GetLogin(conn)
         strPassword = .Properties(PROPNAME_PWD)
         strServer = .Properties(PROPNAME_SOURCE)
         .Close
         GetServerData = True
         lblUser.Caption = strUser
         lblServer.Caption = strServer
      End If
   End With

GetServerData_X:
   Set conn = Nothing
   Exit Function

GetServerData_E:
   Select Case Err
      Case -2147217842
         ' Canceled
      Case Else
         MsgBox Err.Description, vbCritical
   End Select
   Resume GetServerData_X

End Function


Private Function GetLogin(conn As ADODB.Connection) As String
   On Error GoTo GetLogin_E
   Dim rst As New ADODB.Recordset

   With rst
      .ActiveConnection = conn
      .CursorLocation = adUseClient
      .CursorType = adOpenStatic
      .Open ("select SYSTEM_USER")
      GetLogin = .Fields(0).Value
      .Close
   End With
   Set rst = Nothing

GetLogin_X:
   If rst.State = adStateOpen Then rst.Close
   Set rst = Nothing
   Exit Function

GetLogin_E:
   MsgBox Err.Description, vbCritical
   Resume GetLogin_X

End Function

Private Function BackupDatabase(ByVal Server As SQLDMO.SQLServer, ByVal DbName As String) As Boolean
   ' Create a Backup object and set action and source database properties.

   'Dim oBackup As New SQLDMO.Backup
   '
   '
   ''   Dim db As SQLDMO.Database
   ''   Set db = oSvr.Databases(DbName)
   '   Set mBackupEvents = oBackup
   '   DoEvents
   '   With oBackup
   '      .Action = SQLDMOBackup_Database
   '      .Database = DbName
   '      .Files = "e:\programi\MSSQL7\BACKUP\" & DbName & Format(Date, "yyyy-mm-dd") & ".bak"
   '      .SQLBackup Server
   '   End With
   '   Set mBackupEvents = Nothing

End Function
Private Function DBExists(ByVal Name As String) As Boolean
   On Error Resume Next

   DBExists = False
   DBExists = TypeName(oSvr.Databases(Name)) = "Database"

End Function
