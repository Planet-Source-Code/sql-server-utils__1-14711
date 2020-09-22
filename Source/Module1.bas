Attribute VB_Name = "mGlobals"
Option Explicit

Public Const SAVE_FLAG = OFN_LONGNAMES Or OFN_NOCHANGEDIR Or _
      OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT _
      Or OFN_EXTENSIONDIFFERENT
Public Const OPEN_FLAG = OFN_LONGNAMES Or OFN_NOCHANGEDIR Or _
      OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST

Public Const DATA_FILTER = "SQL Server Data Files (*.mdf)|*.mdf|All Files (*.*)|*.*"
Public Const LOG_FILTER = "SQL Server LOg Files (*.ldf)|*.ldf|All Files (*.*)|*.*"
Public Const SQL_FILTER = "SQL Scripts (*.sql)|*.sql|All Files (*.*)|*.*"

Public Const PROPVAL_DATAPROVIDER = "SQLOLEDB"
Public Const PROPNAME_SOURCE = "Data Source"
Public Const PROPVAL_SOURCE = "(local)"
Public Const PROPNAME_PERSEC = "Persist Security Info"
Public Const PROPNAME_PROMPT = "Prompt"
Public Const PROPNAME_UID = "User ID"
Public Const PROPNAME_PWD = "Password"


