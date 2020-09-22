Attribute VB_Name = "mSysmenu"
Option Explicit

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
   

Private Type MENUITEMINFO
   cbSize As Long
   fMask As Long
   fType As Long
   fState As Long
   wID As Long
   hSubMenu As Long
   hbmpChecked As Long
   hbmpUnchecked As Long
   dwItemData As Long
   dwTypeData As String
   cch As Long
End Type



Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
         (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
         (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" _
         (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" _
         (ByVal hwnd&, ByVal bRevert&) As Long
Private Declare Function DeleteMenu Lib "user32" _
         (ByVal hMenu&, ByVal nPosition&, ByVal wFlags&) As Long


Public Enum eSysMenuItems
   SC_SIZE = &HF000&
   SC_MOVE = &HF010&
   SC_CLOSE = &HF060&
   SC_MINIMIZE = &HF020&
   SC_MAXIMIZE = &HF030&
   SC_RESTORE = &HF120&
End Enum

Public Enum eMenuItemMask
   MIIM_STATE = &H1
   MIIM_ID = &H2
   MIIM_SUBMENU = &H4
   MIIM_CHECKMARKS = &H8
   MIIM_TYPE = &H10
   MIIM_DATA = &H20
End Enum

Public Enum MenuFlags
  MF_INSERT = &H0
  MF_ENABLED = &H0
  MF_UNCHECKED = &H0
  MF_BYCOMMAND = &H0
  MF_STRING = &H0
  MF_UNHILITE = &H0
  MF_GRAYED = &H1
  MF_DISABLED = &H2
  MF_BITMAP = &H4
  MF_CHECKED = &H8
  MF_POPUP = &H10
  MF_MENUBARBREAK = &H20
  MF_MENUBREAK = &H40
  MF_HILITE = &H80
  MF_CHANGE = &H80
  MF_END = &H80                    ' Obsolete -- only used by old RES files
  MF_APPEND = &H100
  MF_OWNERDRAW = &H100
  MF_DELETE = &H200
  MF_USECHECKBITMAPS = &H200
  MF_BYPOSITION = &H400
  MF_SEPARATOR = &H800
  MF_REMOVE = &H1000
  MF_DEFAULT = &H1000
  MF_SYSMENU = &H2000
  MF_HELP = &H4000
  MF_RIGHTJUSTIFY = &H4000
  MF_MOUSESELECT = &H8000&
End Enum

Public Enum eMenuItemType
   MFT_RADIOCHECK = &H200&
   MFT_RIGHTORDER = &H2000
   MFT_STRING = MF_STRING
   MFT_BITMAP = MF_BITMAP
   MFT_MENUBARBREAK = MF_MENUBARBREAK
   MFT_MENUBREAK = MF_MENUBREAK
   MFT_OWNERDRAW = MF_OWNERDRAW
   MFT_SEPARATOR = MF_SEPARATOR
   MFT_RIGHTJUSTIFY = MF_RIGHTJUSTIFY
End Enum

Public Enum eMenuItemState
   MFS_DEFAULT = &H1000&
   MFS_CHECKED = &H8&
   MFS_DISABLED = &H2&
   MFS_ENABLED = &H0&
   MFS_GRAYED = &H3&
   MFS_HILITE = &H80&
   MFS_UNCHECKED = &H0&
   MFS_UNHILITE = &H0&
End Enum
   
Private Const WS_THICKFRAME = &H40000
Private Const WS_SIZEBOX = WS_THICKFRAME
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000

Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_FRAMECHANGED = &H20

Private Const GWL_STYLE As Long = (-16&)

Public Const WM_SYSCOMMAND As Long = &H112&

Public Function RemoveMenuItem(ByVal hwnd As Long, ByVal lItem As eSysMenuItems) As Boolean
Dim lRet&, hMenu&
   
   hMenu = GetSystemMenu(hwnd, False)
   If lItem > 0 Then
      lRet = DeleteMenu(hMenu, lItem, MF_BYCOMMAND)
      If lRet = -1 Then Exit Function
      Select Case lItem
         Case SC_MAXIMIZE
            SetWindowLong hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) Xor WS_MAXIMIZEBOX
         Case SC_MINIMIZE
            SetWindowLong hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) Xor WS_MINIMIZEBOX
      End Select
      SetWindowPos hwnd, 0&, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED
   End If
   RemoveMenuItem = lRet
   
End Function


