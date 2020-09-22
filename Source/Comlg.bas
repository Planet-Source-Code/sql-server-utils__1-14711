Attribute VB_Name = "OpenFile321"
Option Explicit

Private Type OPENFILENAME
   lStructSize As Long
   hwndOwner As Long
   hInstance As Long
   lpstrFilter As String
   lpstrCustomFilter As String
   nMaxCustFilter As Long
   nFilterIndex As Long
   lpstrFile As String
   nMaxFile As Long
   lpstrFileTitle As String
   nMaxFileTitle As Long
   lpstrInitialDir As String
   lpstrTitle As String
   Flags As Long
   nFileOffset As Integer
   nFileExtension As Integer
   lpstrDefExt As String
   lCustData As Long
   lpfnHook As Long
   lpTemplateName As String
End Type

Public Enum eFlags
   OFN_READONLY = &H1
   OFN_OVERWRITEPROMPT = &H2
   OFN_HIDEREADONLY = &H4
   OFN_NOCHANGEDIR = &H8
   OFN_SHOWHELP = &H10
   OFN_ENABLEHOOK = &H20
   OFN_ENABLETEMPLATE = &H40
   OFN_ENABLETEMPLATEHANDLE = &H80
   OFN_NOVALIDATE = &H100
   OFN_ALLOWMULTISELECT = &H200
   OFN_EXTENSIONDIFFERENT = &H400
   OFN_PATHMUSTEXIST = &H800
   OFN_FILEMUSTEXIST = &H1000
   OFN_CREATEPROMPT = &H2000
   OFN_SHAREAWARE = &H4000
   OFN_NOREADONLYRETURN = &H8000
   OFN_NOTESTFILECREATE = &H10000
   OFN_NONETWORKBUTTON = &H20000
   OFN_NOLONGNAMES = &H40000                      '  force no long names for 4.x modules
   OFN_EXPLORER = &H80000                         '  new look commdlg
   OFN_NODEREFERENCELINKS = &H100000
   OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules
   OFN_SHAREFALLTHROUGH = 2
   OFN_SHARENOWARN = 1
   OFN_SHAREWARN = 0
End Enum

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Function SaveDialog( _
         Optional hwnd As Long, _
         Optional FileName As String, _
         Optional Filter As String, _
         Optional Title As String, _
         Optional InitDir As String, _
         Optional ByVal Flags As eFlags) As String

   Dim ofn As OPENFILENAME
   Dim lngRet As Long
   ofn.lStructSize = Len(ofn)
   ofn.hwndOwner = hwnd
   ofn.hInstance = App.hInstance
   If Len(Title) = 0 Then Title = "Open"
   If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
   For lngRet = 1 To Len(Filter)
      If Mid$(Filter, lngRet, 1) = "|" Then Mid$(Filter, lngRet, 1) = Chr$(0)
   Next
   ofn.lpstrFilter = Filter
   ofn.nMaxFile = 255
   ofn.nMaxFileTitle = 255
   ofn.lpstrFile = FileName & Space$(254 - Len(FileName))
   ofn.lpstrFileTitle = Space$(255)
   ofn.lpstrInitialDir = InitDir
   ofn.lpstrTitle = Title
   ofn.lpstrDefExt = 1
   If Flags = 0 Then Flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
   ofn.Flags = Flags
   lngRet = GetSaveFileName(ofn)

   If lngRet Then
      SaveDialog = Trim$(ofn.lpstrFile)
   Else
      SaveDialog = vbNullString
   End If

End Function


Function OpenDialog( _
         Optional hwnd As Long, _
         Optional FileName As String, _
         Optional Filter As String, _
         Optional Title As String, _
         Optional InitDir As String, _
         Optional ByVal Flags As eFlags) As String


   Dim ofn As OPENFILENAME
   Dim lngRet As Long
   ofn.lStructSize = Len(ofn)
   ofn.hwndOwner = hwnd
   ofn.hInstance = App.hInstance
   If Len(Filter) = 0 Then Filter = "All Files (*.*)|*.*"
   If Len(Title) = 0 Then Title = "Open"
   If Len(InitDir) = 0 Then InitDir = vbNullString
   If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
   For lngRet = 1 To Len(Filter)
      If Mid$(Filter, lngRet, 1) = "|" Then Mid$(Filter, lngRet, 1) = Chr$(0)
   Next
   ofn.lpstrFilter = Filter
   ofn.nMaxFile = 255
   ofn.nMaxFileTitle = 255
   ofn.lpstrFile = FileName & Space$(254 - Len(FileName))
   ofn.lpstrFileTitle = Space$(255)
   ofn.lpstrInitialDir = InitDir
   ofn.lpstrTitle = Title
   ofn.lpstrDefExt = 1
   If Flags = 0 Then Flags = OFN_HIDEREADONLY + OFN_FILEMUSTEXIST + OFN_NOCHANGEDIR
   ofn.Flags = Flags
   lngRet = GetOpenFileName(ofn)

   If lngRet Then
      OpenDialog = Trim$(ofn.lpstrFile)
      If Len(FileName) > 0 Then FileName = Trim$(ofn.lpstrFileTitle)
   Else
      OpenDialog = vbNullString
      FileName = vbNullString
   End If

End Function

Public Function GetFolder(ByVal Path As String) As String
   Dim i&
   
   GetFolder = Path
   i = InStrRev(Path, "\", , vbTextCompare)
   If i > 1 Then GetFolder = Left$(Path, i - 1)

End Function

