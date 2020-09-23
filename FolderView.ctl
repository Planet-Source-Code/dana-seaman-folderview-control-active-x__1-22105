VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl FolderView 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   5130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2880
   ScaleHeight     =   5130
   ScaleWidth      =   2880
   ToolboxBitmap   =   "FolderView.ctx":0000
   Begin VB.Timer tmrAutoScroll 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1380
      Top             =   4560
   End
   Begin VB.Timer tmrAutoExpand 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   4560
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   540
      Top             =   4560
   End
   Begin MSComctlLib.TreeView TV 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   8070
      _Version        =   393217
      Indentation     =   220
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FolderView.ctx":0312
            Key             =   "hd"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FolderView.ctx":0666
            Key             =   "dt"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FolderView.ctx":09BA
            Key             =   "ram"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FolderView.ctx":0D8E
            Key             =   "mc"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FolderView.ctx":10EA
            Key             =   "cl"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FolderView.ctx":143E
            Key             =   "cd"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FolderView.ctx":1552
            Key             =   "f35"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FolderView.ctx":18A6
            Key             =   "rte"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FolderView.ctx":1BFA
            Key             =   "op"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FolderView.ctx":1F4E
            Key             =   "new"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FolderView.ctx":24AA
            Key             =   "cab"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FolderView.ctx":27FE
            Key             =   "zip"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FolderView.ctx":2B52
            Key             =   "rem"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FolderView.ctx":2EAA
            Key             =   "rar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FolderView.ctx":31FE
            Key             =   "md"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FolderView.ctx":3552
            Key             =   "ace"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FolderView.ctx":38A6
            Key             =   "cp"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   1920
      Index           =   0
      Left            =   960
      Picture         =   "FolderView.ctx":3BFA
      Top             =   3240
      Visible         =   0   'False
      Width           =   1920
   End
End
Attribute VB_Name = "FolderView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Option Compare Text
'----------------------------------
Const ace_      As String = "ace"
Const cab_      As String = "cab"
Const rar_      As String = "rar"
Const zip_      As String = "zip"
Const stardot   As String = "*."
'----------------------------------

Private Enum SHFolders
'    CSIDL_DESKTOP = &H0
'    CSIDL_INTERNET = &H1
'    CSIDL_PROGRAMS = &H2
'    CSIDL_CONTROLS = &H3
'    CSIDL_PRINTERS = &H4
    CSIDL_PERSONAL = &H5
'    CSIDL_FAVORITES = &H6
'    CSIDL_STARTUP = &H7
'    CSIDL_RECENT = &H8
'    CSIDL_SENDTO = &H9
'    CSIDL_BITBUCKET = &HA
'    CSIDL_STARTMENU = &HB
'    CSIDL_DESKTOPDIRECTORY = &H10
'    CSIDL_DRIVES = &H11
'    CSIDL_NETWORK = &H12
'    CSIDL_NETHOOD = &H13
'    CSIDL_FONTS = &H14
'    CSIDL_TEMPLATES = &H15
'    CSIDL_COMMON_STARTMENU = &H16
'    CSIDL_COMMON_PROGRAMS = &H17
'    CSIDL_COMMON_STARTUP = &H18
'    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
'    CSIDL_APPDATA = &H1A
'    CSIDL_PRINTHOOD = &H1B
'    CSIDL_ALTSTARTUP = &H1D '// DBCS
'    CSIDL_COMMON_ALTSTARTUP = &H1E '// DBCS
'    CSIDL_COMMON_FAVORITES = &H1F
'    CSIDL_INTERNET_CACHE = &H20
'    CSIDL_COOKIES = &H21
'    CSIDL_HISTORY = &H22
End Enum
'Formerly from multi-lingual resource file.
'Using English strings for demo.
Private Const s1630 As String = "A folder cannot be dropped onto "
Private Const s1631 As String = "its parent folder."
Private Const s1632 As String = "itself."
Private Const s1633 As String = "a subfolder of itself."
Private Const s1634  As String = "Successful drop of folder "
Private Const s1635  As String = "onto "
Private Const s1636  As String = "A File cannot be dropped onto another File."
'-----------
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260
Private Const SHGFI_DISPLAYNAME = &H200
Private Const I_IMAGECALLBACK = (-1)
Private Const WM_HSCROLL = &H114
Private Const WM_VSCROLL = &H115
Private Const WS_HSCROLL = &H100000
Private Const WS_VSCROLL = &H200000
' Scroll Bar Commands for WM_H/VSCROLL
Private Const SB_LINEUP = 0
Private Const SB_LINELEFT = 0
Private Const SB_LINEDOWN = 1
Private Const SB_LINERIGHT = 1
' TVM_EXPAND wParam action flags
Private Const TVE_EXPAND = &H2
Private Const GWL_STYLE = (-16)
Private Const SM_CXDRAG = 68
Private Const SM_CYDRAG = 69

' TVM_GETNEXTITEM wParam values
Private Const TVGN_ROOT = &H0
Private Const TVGN_PARENT = &H3
'Private Const TVGN_CHILD = &H4
Private Const TVGN_DROPHILITE = &H8

' TVM_GET/SETITEM lParam
' TVITEM mask
'Private Const TVIF_TEXT = &H1
Private Const TVIF_IMAGE = &H2
Private Const TVIF_STATE = &H8
'Private Const TVIF_SELECTEDIMAGE = &H20
Private Const TVIF_CHILDREN = &H40

' TVITEM state, stateMask
Private Const TVIS_EXPANDED = &H20

Private Const TV_FIRST = &H1100
Private Const TVM_EXPAND = (TV_FIRST + 2)
Private Const TVM_GETITEMRECT = (TV_FIRST + 4)
Private Const TVM_GETNEXTITEM = (TV_FIRST + 10)
Private Const TVM_SELECTITEM = (TV_FIRST + 11)
Private Const TVM_GETITEM = (TV_FIRST + 12)
Private Const TVM_SETITEM = (TV_FIRST + 13)
Private Const TVM_HITTEST = (TV_FIRST + 17)
Private Const TVM_CREATEDRAGIMAGE = (TV_FIRST + 18)
'Private Const TVM_SETBKCOLOR As Long = (TV_FIRST + 29)

Private Const FO_MOVE = &H1
Private Const FO_COPY = &H2
Private Const FO_DELETE = &H3
Private Const FO_RENAME = &H4
Private Const FOF_NOCONFIRMATION = &H10
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_MULTIDESTFILES = &H1

Private Type WIN32_FIND_DATA
   dwFileAttributes  As Long
   ftCreationTime    As Currency   'was FILETIME
   ftLastAccessTime  As Currency   'Currency allows direct
   ftLastWriteTime   As Currency   'storage in Grid
   nFileSizeHigh     As String * 4 'Long
   nFileSizeLow      As String * 4 'Long
   dwReserved0       As Long
   dwReserved1       As Long
   cFileName         As String * 260
   cAlternate        As String * 14
End Type

Private Enum TVHITTESTINFO_flags
'  TVHT_NOWHERE = &H1   ' In the client area, but below the last item
  TVHT_ONITEMICON = &H2
  TVHT_ONITEMLABEL = &H4
  TVHT_ONITEMINDENT = &H8
  TVHT_ONITEMBUTTON = &H10
  TVHT_ONITEMRIGHT = &H20
  TVHT_ONITEMSTATEICON = &H40
  TVHT_ONITEM = (TVHT_ONITEMICON Or TVHT_ONITEMLABEL Or TVHT_ONITEMSTATEICON)

  ' user-defined
  TVHT_ONITEMLINE = (TVHT_ONITEM Or TVHT_ONITEMINDENT Or TVHT_ONITEMBUTTON Or TVHT_ONITEMRIGHT)

'  TVHT_ABOVE = &H100
'  TVHT_BELOW = &H200
'  TVHT_TORIGHT = &H400
'  TVHT_TOLEFT = &H800
End Enum
Private Type SHFILEOPSTRUCT
     hwnd As Long
     wFunc As Long
     pFrom As String
     pTo As String
     fFlags As Integer
     fAnyOperationsAborted As Boolean
     hNameMappings As Long
     lpszProgressTitle As String
End Type
Private Enum RectFlags
  rfLeft = &H1
  rfTop = &H2
  rfRight = &H4
  rfBottom = &H8
End Enum
Private Enum SB_Type
  SB_HORZ = 0
  SB_VERT = 1
  SB_CTL = 2
  SB_BOTH = 3
End Enum
Private Enum ScrollDirectionFlags
  sdLeft = &H1
  sdUp = &H2
  sdRight = &H4
  sdDown = &H8
End Enum

Private Type SHFILEINFO
   hicon          As Long
   iIcon          As Long
   dwAttributes   As Long
   szDisplayName  As String * MAX_PATH
   szTypeName     As String * 80
End Type
Private Type RECT   ' rct
  Left As Long
  TOp As Long
  Right As Long
  Bottom As Long
End Type
Private Type POINTAPI   ' pt
  x As Long
  y As Long
End Type
Private Type Size
  cx As Long
  cy As Long
End Type
Private Type TVHITTESTINFO   ' was TV_HITTESTINFO
  pt As POINTAPI
  flags As TVHITTESTINFO_flags
  hitem As Long
End Type

Private Type TVITEM   ' was TV_ITEM
  mask As Long
  hitem As Long
  state As Long
  stateMask As Long
  pszText As Long   ' pointer, if a string, must be pre-allocated before being filled
  cchTextMax As Long
  iImage As Long
  iSelectedImage As Long
  cChildren As Long
  lParam As Long
End Type
Private Enum CBoolean
  CFalse = 0
  CTrue = 1
End Enum
Private Enum SIF_Mask
  SIF_RANGE = &H1
  SIF_PAGE = &H2
  SIF_POS = &H4
  SIF_DISABLENOSCROLL = &H8
  SIF_TRACKPOS = &H10
  SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
End Enum

Private Type SCROLLINFO
  cbSize As Long
  fMask As SIF_Mask
  nMin As Long
  nMax As Long
  nPage As Long
  nPos As Long
  nTrackPos As Long
End Type

'-------------------
Private DragImage       As Integer
Private Buffer          As String * MAX_PATH
Private FullPath        As String
Private sNodeText       As String
Private FvFilter        As Variant 'to pass array (Split)
Private Nodx            As Node
'If using background (wallpaper)...
'Code embedded in DLL since it uses subclassing
Private m_cTVB          As New cTVBackground
'--------   Drag Drop Treeview stuff  --------
#Const DBG = 1

Private m_hwndTV        As Long     ' TV.hWnd
Private m_himlDrag      As Long     ' handle of imagelist holding drag image
Private m_hitemDrag     As Long     ' treeview handle of dragged item
Private m_cxyAutoScroll As Long     ' distance in which auto-scrolling happens

Private m_iButton       As Integer  ' index of button used for dragging (vbLeftButton, vbRightButton)

Private m_nodDrag       As Node     ' Node reference of dragged item
Private m_szDrag        As Size     ' x and y distance cursor moves before dragging begins, in pixels
Private m_ptBtnDown     As POINTAPI ' screen coods of cursor at button down, in pixels
Private m_ptHotSpot     As POINTAPI ' x and y position of the cursor relative to the drag image origin, in pixels
'---------------------
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetLogicalDrives Lib "kernel32" () As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
           (ByVal pszPath As String, _
            ByVal dwFileAttributes As Long, _
            psfi As SHFILEINFO, _
            ByVal cbSizeFileInfo As Long, _
            ByVal uFlags As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
'Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As Any) As Long  ' lpPoint As POINTAPI) As Long
Private Declare Function ImageList_DragShowNolock Lib "comctl32.dll" (ByVal fShow As Boolean) As CBoolean
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As Any) As Long  ' lpPoint As POINTAPI) As Long
Private Declare Function ImageList_GetIconSize Lib "comctl32.dll" (ByVal hIml As Long, lpcx As Long, lpcy As Long) As Boolean
Private Declare Function ImageList_BeginDrag Lib "comctl32.dll" (ByVal himlTrack As Long, ByVal iTrack As Long, ByVal dxHotspot As Long, ByVal dyHotspot As Long) As CBoolean
Private Declare Function ImageList_DragEnter Lib "comctl32.dll" (ByVal hwndLock As Long, ByVal x As Long, ByVal y As Long) As CBoolean
Private Declare Function ImageList_DragMove Lib "comctl32.dll" (ByVal x As Long, ByVal y As Long) As CBoolean
Private Declare Function ImageList_DragLeave Lib "comctl32.dll" (ByVal hwndLock As Long) As CBoolean
Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal hIml As Long) As CBoolean
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function PtInRect Lib "user32" (lprc As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal fnBar As SB_Type, lpsi As SCROLLINFO) As Boolean
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal uID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Sub ImageList_EndDrag Lib "comctl32.dll" ()


'-----------
'Event Declarations:
Event NodeCheck(ByVal Node As Node) 'MappingInfo=TV,TV,-1,NodeCheck
Attribute NodeCheck.VB_Description = "Occurs when Checkboxes = True and a Node object is checked/unchecked."
Event NodeClick(ByVal Node As Node) 'MappingInfo=TV,TV,-1,NodeClick
Attribute NodeClick.VB_Description = "Occurs when a Node object is clicked."
Event AfterLabelEdit(Cancel As Integer, NewString As String) 'MappingInfo=TV,TV,-1,AfterLabelEdit
Attribute AfterLabelEdit.VB_Description = "Occurs after a user edits the label of the currently selected Node or ListItem object."
Event BeforeLabelEdit(Cancel As Integer) 'MappingInfo=TV,TV,-1,BeforeLabelEdit
Attribute BeforeLabelEdit.VB_Description = "Occurs when a user attempts to edit the label of the currently selected ListItem or Node object."
'Event Collapse(ByVal Node As Node) 'MappingInfo=TV,TV,-1,Collapse
Event DblClick() 'MappingInfo=TV,TV,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when you press and release a mouse button and then press and release it again over an object."
'Event Expand(ByVal Node As Node) 'MappingInfo=TV,TV,-1,Expand
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=TV,TV,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=TV,TV,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=TV,TV,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=TV,TV,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=TV,TV,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=TV,TV,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event HitTest(x As Single, y As Single, HitResult As Integer)
Attribute HitTest.VB_Description = "Occurs in a windowless user control in response to mouse activity."
Public Enum BorderLineStyle
    ccNone
    ccFixedSingle
End Enum
'Default Property Values:

Const m_def_DragDropEnable = 0
Const m_def_FileFilter = "*.ace;*.cab;*.rar;*.zip"
Const m_def_hWnd = 0
Const m_def_Path$ = ""
Const MyComputer$ = "MyComputer"
Const Desktop$ = "Desktop"
'Property Variables:

Dim m_DragDropEnable As Boolean
Dim m_FileFilter As String
Dim m_hWnd As Long
Dim m_Path As String
Dim m_MyDocs As String

Private Sub Timer2_Timer()
   TV.StartLabelEdit
   Timer2.Enabled = False
End Sub
Private Sub tmrAutoExpand_Timer()
   On Error GoTo PROC_ERR
  Dim tvhti As TVHITTESTINFO

#If DBG Then
Debug.Print "tmrAutoExpand_Timer"
#End If

  ' Get the cursor postion in TreeView client coords
  GetCursorPos tvhti.pt
  ScreenToClient m_hwndTV, tvhti.pt

  ' Get any TreeView item under the cursor...
  If TreeView_HitTest(m_hwndTV, tvhti) Then

    ' If the cursor is over the button, label, or icon of a collapsed parent item...
    If (tvhti.flags And (TVHT_ONITEMBUTTON Or TVHT_ONITEM)) And _
           IsTVItemCollapsedParent(m_hwndTV, tvhti.hitem) Then

      ' Hide the drag image
      ImageList_DragShowNolock False

      ' Expand the parent item
      TreeView_Expand m_hwndTV, tvhti.hitem, TVE_EXPAND

      ' Reselect the drop target, the TreeView may have been
      ' scrolled putting a different item under the cursor.
      TreeView_SelectDropTarget m_hwndTV, TreeView_HitTest(m_hwndTV, tvhti)

      ' Repaint the TreeView
      UpdateWindow m_hwndTV

      ' Reshow the drag image
      ImageList_DragShowNolock True

    End If   ' tvhti.flags...
  End If   ' TreeView_HitTest

  ' Turn the timer off
  tmrAutoExpand.Enabled = False

PROC_EXIT:
  Exit Sub
PROC_ERR:
  If ErrMsgBox("tmrAutoExpand_Timer") = vbRetry Then Resume Next

End Sub
Private Sub tmrAutoScroll_Timer()
   On Error GoTo PROC_ERR
  Dim pt As POINTAPI
  Dim rcClient As RECT

  Dim dwRectFlags As RectFlags
  Dim dwScrollFlags As ScrollDirectionFlags
  Dim fDidScroll As Boolean
  Dim tvhti As TVHITTESTINFO

#If DBG Then
Debug.Print "tmrAutoScroll_Timer"
#End If

  ' Get the cursor postion in TreeView client coords and the TreeView's
  ' client rect.
  GetCursorPos pt
  ScreenToClient m_hwndTV, pt
  GetClientRect m_hwndTV, rcClient

  ' If the cursor is within an auto scroll region in the TreeView's client area...
  dwRectFlags = PtInRectRegion(rcClient, m_cxyAutoScroll, pt)
  If dwRectFlags Then

    ' Determine which direction the TreeView can be scrolled...
    dwScrollFlags = IsWindowScrollable(m_hwndTV)

    ' hide the drag image
    ImageList_DragShowNolock False

    ' If the cursor is within the respective drag region specified by the
    ' m_cxyAutoScroll distance, and if the TreeView can be scrolled
    ' in that direction, send the TreeView that respective scroll message.

    If (dwRectFlags And rfLeft) And (dwScrollFlags And sdLeft) Then
      fDidScroll = DoDragScroll(SB_HORZ, WM_HSCROLL, SB_LINELEFT)
    End If

    If (dwRectFlags And rfRight) And (dwScrollFlags And sdRight) Then
      fDidScroll = fDidScroll Or DoDragScroll(SB_HORZ, WM_HSCROLL, SB_LINERIGHT)
    End If

    If (dwRectFlags And rfTop) And (dwScrollFlags And sdUp) Then
      fDidScroll = fDidScroll Or DoDragScroll(SB_VERT, WM_VSCROLL, SB_LINEUP)
    End If

    If (dwRectFlags And rfBottom) And (dwScrollFlags And sdDown) Then
      fDidScroll = fDidScroll Or DoDragScroll(SB_VERT, WM_VSCROLL, SB_LINEDOWN)
    End If

    ' If the TreeView was scrolled, update the drop target item
    If fDidScroll Then
      tvhti.pt = pt
      TreeView_SelectDropTarget m_hwndTV, TreeView_HitTest(m_hwndTV, tvhti)
      UpdateWindow m_hwndTV
    End If

    ' Reshow the drag image
    ImageList_DragShowNolock True

  End If   ' dwRectFlags

  ' Enable the scroll timer accordingly.
  tmrAutoScroll.Enabled = fDidScroll

PROC_EXIT:
  Exit Sub
PROC_ERR:
  If ErrMsgBox("tmrAutoScroll_Timer") = vbRetry Then Resume Next

End Sub
Private Sub TV_AfterLabelEdit(Cancel As Integer, NewString As String)
    'RaiseEvent AfterLabelEdit(Cancel, NewString)
   On Error GoTo PROC_ERR
   Dim i As Integer, Temp As String, Success As Long

   ' Make sure that we have a value in the Label
   If Len(NewString) < 1 Then
Retry:
      ' The Label is empty or illegal
      'MsgBox GetResourceString(1236)
      MsgBox "Error. Enter a valid file name (or Esc)."
      ' enable the Timer to get us back to edit mode
      Timer2.Interval = 100
      Timer2.Enabled = True
   Else  'update entry
      Set Nodx = TV.SelectedItem
      FullPath = BuildFullPath(Nodx)
      For i% = (Len(FullPath) - 1) To 1 Step -1
         If Mid$(FullPath, i, 1) = "\" Then
            Temp = Left$(FullPath, i)
            Exit For
        End If
      Next
      FullPath = Left$(FullPath, Len(FullPath) - 1)
      WindowsDiskOps FullPath, Temp & NewString, 128, Success
      If Success = -1 Then
         GoTo Retry
      End If
   End If

PROC_EXIT:
  Exit Sub
PROC_ERR:
  If ErrMsgBox("TV_AfterLabelEdit") = vbRetry Then Resume Next
    
End Sub
Private Sub TV_BeforeLabelEdit(Cancel As Integer)
   ' RaiseEvent BeforeLabelEdit(Cancel)
   On Error GoTo PROC_ERR
        ' If the label is not empty store the string
        If Len(TV.SelectedItem.Text) > 0 Then
           sNodeText = TV.SelectedItem.Text
        End If

PROC_EXIT:
  Exit Sub
PROC_ERR:
  If ErrMsgBox("TV_BeforeLabelEdit") = vbRetry Then Resume Next

End Sub
Private Sub TV_Collapse(ByVal Node As Node)
   ' RaiseEvent Collapse(Node)
   ' Node.Image = 1
End Sub

Private Sub TV_DblClick()
    RaiseEvent DblClick
End Sub
Private Sub TV_Expand(ByVal Node As Node)
   ' RaiseEvent Expand(Node)
   On Error GoTo PROC_ERR

   Screen.MousePointer = vbHourglass
   If Node.Children = 1 And Node.Child.Children <= 0 Then
       ' Remove the "dummy" item
       TV.Nodes.Remove Node.Child.Index
       ' Enumerate file system items under this node
       Node.Sorted = False
       EnumFilesUnder Node
       Node.Sorted = True
   End If

   Screen.MousePointer = vbDefault

PROC_EXIT:
  Exit Sub
PROC_ERR:
  If ErrMsgBox("TV_Expand") = vbRetry Then Resume Next
  
End Sub
Private Sub TV_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub TV_KeyPress(KeyAscii As Integer)
    '
   On Error GoTo PROC_ERR
  If m_himlDrag And (KeyAscii = vbKeyEscape) Then
    EndDrag1
    EndDrag2
    KeyAscii = 0
  End If

RaiseEvent KeyPress(KeyAscii)

PROC_EXIT:
  Exit Sub
PROC_ERR:
  If ErrMsgBox("TV_KeyPress") = vbRetry Then Resume Next

End Sub
Private Function FileExistsW32FD(sSource As String) As WIN32_FIND_DATA

   Dim hFile As Long
   'Returns True in dwReserved1 if file exists as well as
   'raw data in WIN32_FIND_DATA structure
   hFile = FindFirstFile(sSource, FileExistsW32FD)
   FileExistsW32FD.dwReserved1 = hFile <> INVALID_HANDLE_VALUE
   FindClose hFile

End Function
Private Sub TV_KeyUp(KeyCode As Integer, Shift As Integer)
   '
   On Error GoTo PROC_ERR
        ' If the user hits the Esc key then restore the old label
        If KeyCode = vbKeyEscape Then
           TV.SelectedItem.Text = sNodeText
        End If

   RaiseEvent KeyUp(KeyCode, Shift)

PROC_EXIT:
  Exit Sub
PROC_ERR:
  If ErrMsgBox("TV_KeyUp") = vbRetry Then Resume Next

End Sub
Private Sub TV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
   On Error GoTo PROC_ERR
  Dim hitemDest As Long
  
  If m_himlDrag Then
    EndDrag1

    ' get the drop target
    hitemDest = TreeView_GetDropHilight(m_hwndTV)
    If hitemDest Then

      ' Display an error message if the move is illegal.
      If (hitemDest = m_hitemDrag) Then
         '**itself**
        MsgBox s1630 & s1632
      ElseIf (hitemDest = TreeView_GetParent(m_hwndTV, m_hitemDrag)) Then
         '**parent**
         MsgBox s1630 & s1631
      ElseIf AreTVItemsParentChild(m_hwndTV, m_hitemDrag, hitemDest) Then
         '**one of its children**
         MsgBox s1630 & s1633
      Else
        ' Move the dragged item and its subitems (if any)
        ' to the drop point.

         Dim DropDestin As String, DropSource As String

         Set Nodx = TV.DropHighlight
         'Note: Key is "" if folder else is fullpath
         If Nodx.Key <> "" And m_nodDrag.Key <> "" Then
            '"A File cannot be dropped onto another File."
            MsgBox s1636
            GoTo Ed2
         End If
         DropDestin = Nodx.Key
         DropSource = m_nodDrag.Key
         Set m_nodDrag.Parent = TV.DropHighlight
         m_nodDrag.Selected = True
         m_nodDrag.EnsureVisible
         'move the folders/subfolders to new destination
         WindowsDiskOps DropSource, DropDestin, 4, hitemDest '** Move **
         'notify user of action taken
         MsgBox s1634 & DropSource & Chr$(13) & _
                s1635 & DropDestin
      End If

    End If   ' hitemDest
Ed2:
    EndDrag2
  End If   ' m_himlDrag

 RaiseEvent MouseUp(Button, Shift, x, y)

PROC_EXIT:
  Exit Sub
PROC_ERR:
  If ErrMsgBox("TV_MouseUp") = vbRetry Then Resume Next
       
End Sub
Private Sub TV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo PROC_ERR
' we get 2 mousemoves after esc kepress cancel... !!

If m_DragDropEnable Then
   If (Button = m_iButton) Then
     If (m_himlDrag = 0) And ((m_nodDrag Is Nothing) = False) Then
       BeginDrag
     ElseIf m_himlDrag Then
       DoDrag
     End If
   End If
End If

   RaiseEvent MouseMove(Button, Shift, x, y)
     
PROC_EXIT:
   Exit Sub
PROC_ERR:
  If ErrMsgBox("TV_MouseMove") = vbRetry Then Resume Next
        
End Sub
Private Sub TV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo PROC_ERR
   If Button And (m_himlDrag = 0) Then
      ' Store the currently depressed button index, the position of the cursor
      ' in screen coods, and the Node reference of the item under the cursor.
      m_iButton = Button
      GetCursorPos m_ptBtnDown
      Set m_nodDrag = TV.HitTest(x, y)
      If ((m_nodDrag Is Nothing) = False) Then
         If Right$(m_nodDrag, 4) = ".zip" Then
            DragImage = ImageList1.ListImages("zip").Index - 1
         Else
            DragImage = ImageList1.ListImages("cl").Index - 1
         End If
     End If
  End If

RaiseEvent MouseDown(Button, Shift, x, y)

PROC_EXIT:
  Exit Sub
PROC_ERR:
  If ErrMsgBox("TV_MouseDown") = vbRetry Then Resume Next

End Sub
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=TV,TV,-1,StartLabelEdit
'Public Sub StartLabelEdit()
'    TV.StartLabelEdit
'End Sub
''
''Private Sub TV_NodeClick(ByVal Node As MSComctlLib.Node)
''   RaiseEvent Click
''End Sub

Private Sub UserControl_Initialize()
   Dim nod1 As Node
   Const BITSPIXEL& = 12
   'Original Source for background tiling available from
   'www.vbaccelerator.com

   ' TV Background tiling:
   m_cTVB.Attach TV, UserControl.hwnd
   m_cTVB.Tile.Picture = Image1(0).Picture
   
   TV.ImageList = ImageList1
   'Company info, etc. shown only when control is initialized.
   'This info is replaced when folders are enumerated.
   Set nod1 = TV.Nodes.Add(, , Desktop, UserControl.Name & " by Dana Seaman", "dt")
   Set nod1 = TV.Nodes.Add(Desktop, tvwChild, , "dseaman@ieg.com.br", "new")
   Set nod1 = TV.Nodes.Add(Desktop, tvwChild, , stardot & ace_, ace_)
   Set nod1 = TV.Nodes.Add(Desktop, tvwChild, , stardot & cab_, cab_)
   Set nod1 = TV.Nodes.Add(Desktop, tvwChild, , stardot & rar_, rar_)
   Set nod1 = TV.Nodes.Add(Desktop, tvwChild, , stardot & zip_, zip_)
   nod1.Expanded = True
   nod1.EnsureVisible
   TV.Refresh


End Sub

Private Sub UserControl_InitProperties()
   m_Path = m_def_Path
   m_hWnd = m_def_hWnd
   m_FileFilter = m_def_FileFilter
   m_DragDropEnable = m_def_DragDropEnable

End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   m_hWnd = PropBag.ReadProperty("hWnd", m_def_hWnd)
   TV.Appearance = PropBag.ReadProperty("Appearance", 1)
   TV.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
   TV.Checkboxes = PropBag.ReadProperty("Checkboxes", False)
   TV.Enabled = PropBag.ReadProperty("Enabled", True)
   Set TV.Font = PropBag.ReadProperty("Font", Ambient.Font)
   TV.FullRowSelect = PropBag.ReadProperty("FullRowSelect", False)
   TV.HideSelection = PropBag.ReadProperty("HideSelection", True)
   TV.HotTracking = PropBag.ReadProperty("HotTracking", False)
   TV.LabelEdit = PropBag.ReadProperty("LabelEdit", 1)
   TV.LineStyle = PropBag.ReadProperty("LineStyle", 1)
   Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
   TV.MousePointer = PropBag.ReadProperty("MousePointer", 0)
   TV.PathSeparator = PropBag.ReadProperty("PathSeparator", "\")
   m_FileFilter = PropBag.ReadProperty("FileFilter", m_def_FileFilter)
   Set Picture = PropBag.ReadProperty("Picture", Nothing)
   m_DragDropEnable = PropBag.ReadProperty("DragDropEnable", m_def_DragDropEnable)
   UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
'   TV.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
End Sub
Private Sub UserControl_Resize()
    TV.Move 0, 0, UserControl.Width, UserControl.Height
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   Call PropBag.WriteProperty("hWnd", m_hWnd, m_def_hWnd)
   Call PropBag.WriteProperty("Appearance", TV.Appearance, 1)
   Call PropBag.WriteProperty("BorderStyle", TV.BorderStyle, 0)
   Call PropBag.WriteProperty("Checkboxes", TV.Checkboxes, False)
   Call PropBag.WriteProperty("Enabled", TV.Enabled, True)
   Call PropBag.WriteProperty("Font", TV.Font, Ambient.Font)
   Call PropBag.WriteProperty("FullRowSelect", TV.FullRowSelect, False)
   Call PropBag.WriteProperty("HideSelection", TV.HideSelection, True)
   Call PropBag.WriteProperty("HotTracking", TV.HotTracking, False)
   Call PropBag.WriteProperty("LabelEdit", TV.LabelEdit, 1)
   Call PropBag.WriteProperty("LineStyle", TV.LineStyle, 1)
   Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
   Call PropBag.WriteProperty("MousePointer", TV.MousePointer, 0)
   Call PropBag.WriteProperty("PathSeparator", TV.PathSeparator, "\")
   Call PropBag.WriteProperty("FileFilter", m_FileFilter, m_def_FileFilter)
   Call PropBag.WriteProperty("Picture", Picture, Nothing)
   Call PropBag.WriteProperty("DragDropEnable", m_DragDropEnable, m_def_DragDropEnable)
   Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)

'   Call PropBag.WriteProperty("OLEDropMode", TV.OLEDropMode, 0)
End Sub
Public Sub Enumerate()
   Clear
   FvFilter = Split(m_FileFilter, ";")
   LoadTree6
End Sub
Public Sub SetNodeVisible()
   Dim L4 As Long, Nod As Node
   Dim sFullPath As String
   Dim qPath As String
   qPath = QualifyPath(m_Path)
   For L4 = 1 To TV.Nodes.Count
      Set Nod = TV.Nodes(L4)
      sFullPath = BuildFullPath(Nod)
      If StrComp(sFullPath, qPath, vbTextCompare) = 0 Then
         Nod.EnsureVisible
         Nod.Selected = True
         TV.Refresh
         Exit For
      End If
   Next
End Sub
Public Property Let ImageList(iL As Variant)
    Set TV.ImageList = iL
End Property

Public Sub Clear()
    TV.Visible = False
    TV.Nodes.Clear
    TV.Visible = True
End Sub
Private Function AreTVItemsParentChild(hwndTV As Long, hitemParent As Long, ByVal hitemChild As Long) As Boolean

  Do While hitemChild
    hitemChild = TreeView_GetParent(hwndTV, hitemChild)
    If (hitemChild = hitemParent) Then
      AreTVItemsParentChild = True
      Exit Function
    End If
  Loop
End Function
Private Sub BeginDrag()
   On Error GoTo PROC_ERR
  Dim pt As POINTAPI
  Dim tvhti As TVHITTESTINFO
  Dim rcItem As RECT
  Dim szIcon As Size
  Dim ptImage As POINTAPI

#If DBG Then
Debug.Print "BeginDrag"
#End If

  ' If the cursor position has exceeded the non-drag slop...
  GetCursorPos pt
  If (Abs(pt.x - m_ptBtnDown.x) >= m_szDrag.cx) Or _
     (Abs(pt.y - m_ptBtnDown.y) >= m_szDrag.cy) Then

    ' We're dragging, get the TreeView client coords at MouseDown
    pt = m_ptBtnDown
    ScreenToClient m_hwndTV, pt

    ' Get the handle of the item under the cursor at MouseDown.
    tvhti.pt = pt
    TreeView_HitTest m_hwndTV, tvhti
    If (tvhti.flags And TVHT_ONITEM) Then

      ' Allow dragging of only non-root items.
      If TreeView_GetParent(m_hwndTV, tvhti.hitem) Then

        ' Store the handle of the item being dragged.
        m_hitemDrag = tvhti.hitem

        ' The TreeView_CreateDragImage call fails if the TreeView does not have
        ' an imagelist associated with it, the call succeds but the returned imagelist
        ' contains no drag image if the specified item is not assigned an icon, and
        ' for some reason the call also fails if the TreeView item's iImage member is
        ' using callbacks (I_IMAGECALLBACK, but is not the case for real treeviews...?)
        SetNodeCallbacks m_hwndTV, m_hitemDrag, m_nodDrag, False
        m_himlDrag = TreeView_CreateDragImage(m_hwndTV, m_hitemDrag)
        SetNodeCallbacks m_hwndTV, m_hitemDrag, m_nodDrag, True
        If m_himlDrag Then

'Debug.Print ImageList_GetImageCount(m_himlDrag)
'ImageList_Draw m_himlDrag,   0,  Picture1.hDC, 0, 0, ILD_NORMAL

          ' Get the position of the cursor relative to the image's origin,
          ' for ImageList_BeginDrag, is used to postion the drag image
          ' from the window coords passed to ImageList_DragEnter and
          ' ImageList_DragMove (szIcon.cx is the width of the actual
          ' drag image [the item's icon and label])
          TreeView_GetItemRect m_hwndTV, m_hitemDrag, rcItem, CTrue
          ImageList_GetIconSize m_himlDrag, szIcon.cx, szIcon.cy
          m_ptHotSpot.x = (pt.x - rcItem.Left) + (szIcon.cx - (rcItem.Right - rcItem.Left))
          m_ptHotSpot.y = pt.y - rcItem.TOp

          ' Convert the item label's origin, which is relative to the
          ' Treeview's client area, to coods relative to the TreeView's
          ' window rect origin, for ImageList_DragEnter
          ClientToWindow m_hwndTV, ptImage

          ' Set capture so we get MouseMoves when dragging
          ' outside the TreeView,
          SetCapture m_hwndTV

          ' Select the drop target
          TreeView_SelectDropTarget m_hwndTV, m_hitemDrag

          ' Set the drag imagelist and the image hotspot
          ImageList_BeginDrag m_himlDrag, 0&, m_ptHotSpot.x, m_ptHotSpot.y

          ' Lock the screen and draw the first drag image
          ImageList_DragEnter m_hwndTV, ptImage.x, ptImage.y

        End If   ' m_himlDrag
      End If   ' TreeView_GetParent
    End If   ' TreeView_HitTest

  End If   ' begin drag

PROC_EXIT:
  Exit Sub
PROC_ERR:
  If ErrMsgBox("BeginDrag") = vbRetry Then Resume Next

End Sub
Private Function BuildFullPath(ByVal Nod As Node) As String
   On Error GoTo PROC_ERR
   Dim iPos As Integer
   Dim sExt As String
   Dim MyPath As String
   Dim MyDocs2 As String
   
   MyPath = Nod.FullPath
   
   iPos = InStrRev(MyPath, ":")
   If iPos < 2 Then
     ' Select Case MyPath
   ' If nod.Key = QualifyPath(m_MyDocs) Then
       MyDocs2 = Mid(m_MyDocs, 4)
       BuildFullPath = Replace(MyPath, Desktop & "\" & MyDocs2, m_MyDocs)
       GoTo CheckExt
   ' End If
   End If
   MyPath = Mid$(MyPath, iPos - 1) 'Pick up drive letter

   iPos = InStr(MyPath, "\")
   If iPos > 1 Then
      BuildFullPath = Left$(MyPath, 2) & Mid$(MyPath, iPos)
   Else
      BuildFullPath = Left$(MyPath, 2)
   End If
CheckExt:
   sExt = GetExt(Nod.Text)
   If sExt <> "" Then
      For iPos = 0 To UBound(FvFilter)
         If sExt = GetExt(FvFilter(iPos)) Then 'Match
            Exit Function
         End If
      Next
   End If
   
   BuildFullPath = QualifyPath(BuildFullPath)

PROC_EXIT:
  Exit Function
PROC_ERR:
  If ErrMsgBox("BuildFullPath") = vbRetry Then Resume Next

End Function
Private Function StripNull(ByVal StrIn As String) As String
   On Error GoTo PROC_ERR
   Dim nul As Long
   '
   ' Truncate input string at first null.
   ' If no nulls, perform ordinary Trim.
   '
   nul = InStr(StrIn, vbNullChar)
   Select Case nul
      Case Is > 1
         StripNull = Left$(StrIn, nul - 1)
      Case 1
         StripNull = ""
      Case 0
         StripNull = Trim$(StrIn)
   End Select

PROC_EXIT:
  Exit Function
PROC_ERR:
  If ErrMsgBox("StripNull") = vbRetry Then Resume Next

End Function

Private Function ErrMsgBox(Msg As String) As Integer
    ErrMsgBox = MsgBox("Error: " & Err.Number & ". " & Err.Description, vbRetryCancel + vbCritical, "FolderView." & Msg)
End Function
' ===========================================================================
' Treeview macros defined in Commctrl.h

' Expands or collapses the list of child items, if any, associated with the specified parent item.
' Returns TRUE if successful or FALSE otherwise.
' (docs say TVM_EXPAND does not send the TVN_ITEMEXPANDING and
' TVN_ITEMEXPANDED notification messages to the parent window...?)

Private Function TreeView_Expand(ByVal hwnd As Long, hitem As Long, Flag As Long) As Boolean
  TreeView_Expand = SendMessage(hwnd, TVM_EXPAND, ByVal Flag, ByVal hitem)
End Function
' Retrieves the bounding rectangle for a tree-view item and indicates whether the item is visible.
' If the item is visible and retrieves the bounding rectangle, the return value is TRUE.
' Otherwise, the TVM_GETITEMRECT message returns FALSE and does not retrieve
' the bounding rectangle.
' If fItemRect = TRUE, returns label rect. Otherwise, entire item line rect is  returned.

Private Function TreeView_GetItemRect(ByVal hwnd As Long, hitem As Long, prc As RECT, fItemRect As CBoolean) As Boolean
  prc.Left = hitem
  TreeView_GetItemRect = SendMessage(hwnd, TVM_GETITEMRECT, ByVal fItemRect, prc)
End Function
' TreeView_GetNextItem

' Retrieves the tree-view item that bears the specified relationship to a specified item.
' Returns the handle to the item if successful or 0 otherwise.

Private Function TreeView_GetNextItem(ByVal hwnd As Long, hitem As Long, Flag As Long) As Long
  TreeView_GetNextItem = SendMessage(hwnd, TVM_GETNEXTITEM, ByVal Flag, ByVal hitem)
End Function
'
'' Retrieves the first child item. The hitem parameter must be NULL.
'' Returns the handle to the item if successful or 0 otherwise.
'
'Private Function TreeView_GetChild(byval hwnd As Long, hitem As Long) As Long
'  TreeView_GetChild = TreeView_GetNextItem(hwnd, hitem, TVGN_CHILD)
'End Function
' Retrieves the parent of the specified item.
' Returns the handle to the item if successful or 0 otherwise.
Private Function TreeView_GetParent(ByVal hwnd As Long, hitem As Long) As Long
  TreeView_GetParent = TreeView_GetNextItem(hwnd, hitem, TVGN_PARENT)
End Function
' Retrieves the item that is the target of a drag-and-drop operation.
' Returns the handle to the item if successful or 0 otherwise.

Private Function TreeView_GetDropHilight(hwnd As Long) As Long
  TreeView_GetDropHilight = TreeView_GetNextItem(hwnd, 0, TVGN_DROPHILITE)
End Function
' Retrieves the topmost or very first item of the tree-view control.
' Returns the handle to the item if successful or 0 otherwise.

Private Function TreeView_GetRoot(hwnd As Long) As Long
  TreeView_GetRoot = TreeView_GetNextItem(hwnd, 0, TVGN_ROOT)
End Function
' TreeView_Select
' Selects the specified tree-view item, scrolls the item into view, or redraws the item
' in the style used to indicate the target of a drag-and-drop operation.
' If hitem is NULL, the selection is removed from the currently selected item, if any.
' Returns TRUE if successful or FALSE otherwise.
Private Function TreeView_Select(ByVal hwnd As Long, hitem As Long, code As Long) As Boolean
  TreeView_Select = SendMessage(hwnd, TVM_SELECTITEM, ByVal code, ByVal hitem)
End Function

' Redraws the given item in the style used to indicate the target of a drag and drop operation.
' Returns TRUE if successful or FALSE otherwise.

Private Function TreeView_SelectDropTarget(ByVal hwnd As Long, hitem As Long) As Boolean
  TreeView_SelectDropTarget = TreeView_Select(hwnd, hitem, TVGN_DROPHILITE)
End Function

' Retrieves some or all of a tree-view item's attributes.
' Returns TRUE if successful or FALSE otherwise.

Private Function TreeView_GetItem(ByVal hwnd As Long, pItem As TVITEM) As Boolean
  TreeView_GetItem = SendMessage(hwnd, TVM_GETITEM, 0, pItem)
End Function
' Sets some or all of a tree-view item's attributes.
' Old docs say returns zero if successful or - 1 otherwise.
' New docs say returns TRUE if successful, or FALSE otherwise
Private Function TreeView_SetItem(ByVal hwnd As Long, pItem As TVITEM) As Boolean
  TreeView_SetItem = SendMessage(hwnd, TVM_SETITEM, 0, pItem)
End Function
' Determines the location of the specified point relative to the client area of a tree-view control.
' Returns the handle to the tree-view item that occupies the specified point or NULL if no item
' occupies the point.

Private Function TreeView_HitTest(ByVal hwnd As Long, lpht As TVHITTESTINFO) As Long
  TreeView_HitTest = SendMessage(hwnd, TVM_HITTEST, 0, lpht)
End Function

' Creates a dragging bitmap for the specified item in a tree-view control, creates an image list
' for the bitmap, and adds the bitmap to the image list. An application can display the image
' when dragging the item by using the image list functions.
' Returns the handle of the image list to which the dragging bitmap was added if successful or
' NULL otherwise.
Private Function TreeView_CreateDragImage(ByVal hwnd As Long, hitem As Long) As Long
  TreeView_CreateDragImage = SendMessage(hwnd, TVM_CREATEDRAGIMAGE, 0, ByVal hitem)
End Function
' Returns True if the specified treeview item is a collapsed parent
' (has a buttom and is collapsed). returns False otherwise
Private Function IsTVItemCollapsedParent(hwndTV As Long, hitem As Long) As Boolean
  Dim tvi As TVITEM

  tvi.hitem = hitem
  tvi.mask = TVIF_STATE Or TVIF_CHILDREN

  If TreeView_GetItem(hwndTV, tvi) Then
    IsTVItemCollapsedParent = ((tvi.state And TVIS_EXPANDED) = False) And (tvi.cChildren <> 0)
  End If

End Function
' Returns a set of bit flags indicating whether the specified point resides in
' the specified size region with the perimeter of the specified rect. cxyRegion
' defines the rectangular region within rc, and must be a positive value

Private Function PtInRectRegion(RC As RECT, cxyRegion As Long, pt As POINTAPI) As RectFlags
  Dim dwFlags As RectFlags

  If PtInRect(RC, pt.x, pt.y) Then
    dwFlags = (rfLeft And (pt.x <= (RC.Left + cxyRegion)))
    dwFlags = dwFlags Or (rfRight And (pt.x >= (RC.Right - cxyRegion)))
    dwFlags = dwFlags Or (rfTop And (pt.y <= (RC.TOp + cxyRegion)))
    dwFlags = dwFlags Or (rfBottom And (pt.y >= (RC.Bottom - cxyRegion)))
  End If

  PtInRectRegion = dwFlags

End Function
' Returns a set of bit flags indicating whether the specified
' window can be scrolled in any given direction.
Private Function IsWindowScrollable(hwnd As Long) As ScrollDirectionFlags
  Dim si As SCROLLINFO
  Dim dwScrollFlags As ScrollDirectionFlags

  si.cbSize = Len(si)
  si.fMask = SIF_ALL

  ' Get the horizontal scrollbar's info (GetScrollInfo returns
  ' TRUE after a scrollbar has been added to a window,
  ' even if the respective style bit is not set...)
  If (GetWindowLong(hwnd, GWL_STYLE) And WS_HSCROLL) Then
    If GetScrollInfo(hwnd, SB_HORZ, si) Then
      dwScrollFlags = (sdLeft And (si.nPos > 0))
      dwScrollFlags = dwScrollFlags Or (sdRight And (si.nPos < (((si.nMax - si.nMin) + 1) - si.nPage)))
    End If
  End If

  ' Get the vertical scrollbar's info.
  If (GetWindowLong(hwnd, GWL_STYLE) And WS_VSCROLL) Then
    If GetScrollInfo(hwnd, SB_VERT, si) Then
      dwScrollFlags = dwScrollFlags Or (sdUp And (si.nPos > 0))
      dwScrollFlags = dwScrollFlags Or (sdDown And (si.nPos < (((si.nMax - si.nMin) + 1) - si.nPage)))
    End If
  End If

  IsWindowScrollable = dwScrollFlags

End Function
' Toggles the callback attributes (text, images, button) for the specific
' item (Node), as specified by the fSet flag. The TreeView normally
' does callback for these items, but can be explicitly set and overridden
' (and the VB TreeView has no idea what's happening).
Private Function SetNodeCallbacks(hwndTV As Long, _
                                 hitem As Long, _
                                 Nod As Node, _
                                 fSet As Boolean) As Boolean

  Dim tvi As TVITEM

  tvi.mask = TVIF_IMAGE  ' Or TVIF_SELECTEDIMAGE Or TVIF_TEXT  ' Or TVIF_CHILDREN
  tvi.hitem = hitem

  If fSet Then
    ' Set the callbacks
    tvi.iImage = I_IMAGECALLBACK
'    tvi.iSelectedImage = I_IMAGECALLBACK
'    tvi.pszText = LPSTR_TEXTCALLBACK
'' this causes problems, something's going on with the
'' VB TreeView here that has yet to be understood...
'    tvi.cChildren = I_CHILDRENCALLBACK

  Else
    ' Get the Node from the hItem, and remove the item's
    ' callback attributes by explicitly setting them.
    If ((Nod Is Nothing) = False) Then
      ' real imagelist indices are zero-based

      tvi.iImage = DragImage
     'bogus type missmatch tvi.iImage = nod.Image - 1


'      tvi.iSelectedImage = CLng(nod.SelectedImage) - 1
'      ' Store the Node's Text in an allocated pointer'
'      tvi.pszText = StrPtr(String$(MAX_ITEM, 0))
'      lstrcpyA ByVal tvi.pszText, ByVal nod.Text
'' see above...
'      tvi.cChildren = Abs(CBool(TreeView_GetChild(m_hwndTV, hitem)))
    End If

  End If   ' fSet

  SetNodeCallbacks = TreeView_SetItem(hwndTV, tvi)

End Function
' Converts the specified window's client coords to window
' coords (relative to the window's rect origin)
Private Function ClientToWindow(ByVal hwnd As Long, pt As POINTAPI) As Boolean
  Dim fRtn As Boolean
  Dim rcClient As RECT
  Dim rcWindow As RECT

  If IsWindow(hwnd) Then
    fRtn = CBool(GetClientRect(hwnd, rcClient))
    fRtn = fRtn And CBool(ClientToScreen(hwnd, rcClient))
    fRtn = fRtn And CBool(GetWindowRect(hwnd, rcWindow))
    If fRtn Then
      pt.x = pt.x + (rcClient.Left - rcWindow.Left)
      pt.y = pt.y + (rcClient.TOp - rcWindow.TOp)
      ClientToWindow = True
    End If
  End If

End Function
Private Sub WindowsDiskOps(Source As String, dest As String, Flavor As Long, Success As Long)

Dim Result As Long
Dim fileop As SHFILEOPSTRUCT

With fileop
.hwnd = 0
   Select Case Flavor
      Case 1                           ' SmartCopy
         .wFunc = FO_COPY
      Case 2                           ' Copy
         .wFunc = FO_COPY
      Case 4                           ' Move
         .wFunc = FO_MOVE
      Case 8                           ' Delete
         .wFunc = FO_DELETE
      Case 16                          ' Delete (Recycle bin)
         .wFunc = FO_DELETE
         .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION
      Case 64                          ' Vault
         .wFunc = FO_COPY
         .fFlags = FOF_MULTIDESTFILES
      Case 128                         ' Rename
         .wFunc = FO_RENAME
   End Select
  '.lpszProgressTitle = ""
   .pFrom = Source & vbNullChar & vbNullChar    ' The files to copy separated by Nulls and terminated by 2 nulls
   .pTo = dest & vbNullChar & vbNullChar                ' The directory or filename(s) to copy into terminated in 2 nulls

End With
   Result = SHFileOperation(fileop)
   If Result <> 0 Then 'Operation failed
      'Msgbox the error that occurred in the API.
      MsgBox Err.LastDllError, vbCritical Or vbOKOnly
   Else
      If fileop.fAnyOperationsAborted <> 0 Then
         MsgBox "Operation Failed", vbCritical Or vbOKOnly
         Success = -1
      End If
   End If
End Sub
Private Function DoDragScroll(nBar As Long, uMsg As Long, dwCmd As Long) As Boolean
   On Error GoTo PROC_ERR
  Dim dwPos  As Long

  dwPos = GetScrollPos(m_hwndTV, nBar)
  SendMessage m_hwndTV, uMsg, dwCmd, ByVal 0&
  If (dwPos <> GetScrollPos(m_hwndTV, nBar)) Then
    UpdateWindow m_hwndTV
    DoDragScroll = True
  End If
  
PROC_EXIT:
  Exit Function
PROC_ERR:
  If ErrMsgBox("DoDragScroll") = vbRetry Then Resume Next

End Function
Private Sub EnumFilesUnder(ByVal n As Node)

   On Error GoTo PROC_ERR
    Dim sPath As String
    Dim sExt As String
    Dim hFind As Long, L4 As Long
    Dim oldPath As String
    Dim W32FD As WIN32_FIND_DATA
    Dim n2 As Node
    Dim FolderPic As String
    
    TV.Visible = False
    oldPath = ""
    sPath = BuildFullPath(n) & "*.*"
    'old sPath = ucase$(n.FullPath & "\*.*")
    hFind = FindFirstFile(sPath, W32FD)
    Do
        ' Get the filename, if any.
        sPath = StripNull(W32FD.cFileName)
        If Len(sPath) = 0 Or StrComp(sPath, oldPath) = 0 Then
            ' Nothing found?
            Exit Do
        ElseIf Asc(sPath) <> 46 Then
           'do we have a folder?
           If (W32FD.dwFileAttributes And vbDirectory) Then 'Yes
               FolderPic = "cl"
               Set n2 = TV.Nodes.Add(n, tvwChild, , sPath, FolderPic)
               n2.ExpandedImage = "op"
               'causes duplicate keys in My Documents
               'n2.Key = BuildFullPath(n2)
               ' Add a dummy item so the + sign is
               ' displayed
               If hasSubDirectory(BuildFullPath(n) & sPath & "\") Then
                  TV.Nodes.Add n2, tvwChild
               End If
           Else  'do we have a matching file?
              sExt = GetExt(sPath)
              For L4 = 0 To UBound(FvFilter)
                 If sPath Like FvFilter(L4) Then 'Yes
                    Select Case sExt
                       Case "zip", "cab", "ace", "rar"
                          FolderPic = sExt
                       Case Else
                          FolderPic = "new"
                    End Select
                    Set n2 = TV.Nodes.Add(n, tvwChild, , sPath, FolderPic)
                    'n2.Key = BuildFullPath(n2)
                    ' TV.Nodes.Item(TV.Nodes.Count).Bold = True
                    '***Node colors don't work if you are using
                    '   background (wallpaper) in Treeview
                    TV.Nodes.Item(TV.Nodes.Count).BackColor = vbBlue '&H98CCD0   '&HE0E0E0    'grey
                    TV.Nodes.Item(TV.Nodes.Count).ForeColor = vbWhite    'RGB(248, 240, 136) 'Tree ylw
                    Exit For
                 End If
              Next
           End If
        End If
        FindNextFile hFind, W32FD
        oldPath = sPath
    Loop
    FindClose hFind
    TV.Visible = True
    Exit Sub

PROC_EXIT:
  Exit Sub
PROC_ERR:
  If ErrMsgBox("EnumFilesUnder") = vbRetry Then Resume Next

End Sub
Private Function GetExt(ByVal Name As String) As String
   On Error GoTo PROC_ERR
   Dim j As Integer
   j = InStrRev(Name, ".")
   If j > 0 And j < Len(Name) Then
      GetExt = LCase$(Mid$(Name, j + 1))
   End If

PROC_EXIT:
  Exit Function
PROC_ERR:
  If ErrMsgBox("GetExt") = vbRetry Then Resume Next

End Function
Private Function hasSubDirectory(ByVal sPath As String) As Boolean
    On Error GoTo PROC_ERR
    
    Dim hFind As Long
    Dim oldPath As String
    Dim W32FD As WIN32_FIND_DATA
    Dim L4 As Long
    
    oldPath = ""
    
    hFind = FindFirstFile(sPath & "*.*", W32FD)
    Do
        ' Get the filename, if any.
        sPath = StripNull(W32FD.cFileName)
        If Len(sPath) = 0 Or StrComp(sPath, oldPath) = 0 Then
            ' Nothing found?
            Exit Do
        ElseIf Asc(sPath) <> 46 Then
            ' return true if we have found a directory under this path
            If (W32FD.dwFileAttributes And vbDirectory) Then
                hasSubDirectory = True
                Exit Do
            End If
            For L4 = 0 To UBound(FvFilter)
                If sPath Like FvFilter(L4) Then
                  hasSubDirectory = True
                  Exit Do
                End If
            Next
        End If
        FindNextFile hFind, W32FD
        oldPath = sPath
    Loop
    FindClose hFind

PROC_EXIT:
  Exit Function
PROC_ERR:
  If ErrMsgBox("hasSubDirectory") = vbRetry Then Resume Next

End Function
Private Sub EndDrag1()
   On Error GoTo PROC_ERR

#If DBG Then
Debug.Print "EndDrag"
#End If

  If m_himlDrag Then
    ReleaseCapture
    Screen.MousePointer = vbDefault

    ImageList_DragLeave m_hwndTV
    ImageList_EndDrag

    ImageList_Destroy m_himlDrag
    m_himlDrag = 0

    tmrAutoScroll.Enabled = False
    tmrAutoExpand.Enabled = False
  End If

PROC_EXIT:
  Exit Sub
PROC_ERR:
  If ErrMsgBox("EndDrag1") = vbRetry Then Resume Next

End Sub
Private Sub EndDrag2()
   On Error GoTo PROC_ERR
  TreeView_SelectDropTarget m_hwndTV, 0
  TV.Refresh

  m_hitemDrag = 0
  Set m_nodDrag = Nothing

PROC_EXIT:
  Exit Sub
PROC_ERR:
  If ErrMsgBox("EndDrag2") = vbRetry Then Resume Next

End Sub
Private Sub DoDrag()
   On Error GoTo PROC_ERR
  Dim pt As POINTAPI
  Dim tvhti As TVHITTESTINFO
  Dim hitem As Long
  Dim rcClient As RECT
  Static hitemPrev As Long


 ' trying to fix jump problem in tree

#If DBG Then
Debug.Print "DoDrag"
#End If

  ' Get the cursor postion in TreeView client coords
  GetCursorPos pt
  ScreenToClient m_hwndTV, pt

  ' Unlock the treeview's painting
  ImageList_DragLeave m_hwndTV

  ' Highlights the new drop target if the cursor is over an item,
  ' removes the drop highlight otherwise...
  tvhti.pt = pt
  hitem = TreeView_HitTest(m_hwndTV, tvhti)
  TreeView_SelectDropTarget m_hwndTV, hitem

  ' If the cursor is over an item and TreeView has a horizontal scrollbar, assume
  ' that the item's tooltip is shown, and update the TreeView on each drag move
  ' (for some reason the TreeView's tooltip does not erase correctly...).
  If CBool(tvhti.flags And TVHT_ONITEMLINE) And _
        (GetWindowLong(m_hwndTV, GWL_STYLE) And WS_HSCROLL) Then
    UpdateWindow m_hwndTV
  End If

  ' Convert the item label's origin, which is relative to the
  ' Treeview's client area, to coods relative to the TreeView's
  ' window rect origin, for ImageList_DragEnter
  ClientToWindow m_hwndTV, pt

  ' Erase the old drag image and draw a new one.
  ImageList_DragMove pt.x, pt.y

  ' Lock the treeview's painting again
  ImageList_DragEnter m_hwndTV, pt.x, pt.y

  ' Modify the cursor to provide visual feedback to the user.
  ' Note: It's important to do this AFTER the call to DragMove. (so says Jeff...??)
  If hitem Then
    Screen.MousePointer = vbDefault
  Else
'    Screen.MousePointer = vbNoDrop
  End If

  ' If the cursor is still over same item as it was on the previous call,
  ' the cursor is over button, label, or icon of a collapsed parent item,
  ' start the auto expand timer, disable the timer otherwise.
  If (hitem = hitemPrev) And _
        (tvhti.flags And (TVHT_ONITEMBUTTON Or TVHT_ONITEM)) And _
         IsTVItemCollapsedParent(m_hwndTV, hitem) Then
    tmrAutoExpand.Enabled = True
  Else
    tmrAutoExpand.Enabled = False
  End If

  ' cache the current item's handle as the previous item's handle
  hitemPrev = hitem

  ' If the window is scrollable, and the cursor is within that auto scroll
  ' distance, start the auto scroll timer, disable the timer otherwise.
  GetClientRect m_hwndTV, rcClient
'Debug.Print "scroll: &H" & hex$(IsWindowScrollable(hwndTV)) & _
                    " , region &H" & hex$(PtInRectRegion(rc, m_cxyAutoScroll, tvhti.pt))
  If (IsWindowScrollable(m_hwndTV) And _
          PtInRectRegion(rcClient, m_cxyAutoScroll, tvhti.pt)) Then
    tmrAutoScroll.Enabled = True
  Else
    tmrAutoScroll.Enabled = False
  End If

PROC_EXIT:
  Exit Sub
PROC_ERR:
  If ErrMsgBox("DoDrag") = vbRetry Then Resume Next

End Sub
Private Function GetResourceStringFromFile(sModule As String, idString As Long) As String

   Dim hModule As Long
   Dim nChars As Long

   hModule = LoadLibrary(sModule)
   If hModule Then
      nChars = LoadString(hModule, idString, Buffer, MAX_PATH)
      If nChars Then
         GetResourceStringFromFile = Left$(Buffer, nChars)
      End If
      FreeLibrary hModule
   End If
End Function
Private Function QualifyPath(ByVal MyString As String) As String
   If Right$(MyString, 1) <> "\" Then
      QualifyPath = MyString & "\"
   Else
      QualifyPath = MyString
   End If
End Function
Private Function FolderLocation(lFolder As SHFolders) As String

   Dim lp As Long
   'Get the PIDL for this folder
  ' SHGetSpecialFolderLocation MyForm.hwnd, lFolder, lp
   SHGetSpecialFolderLocation 0&, lFolder, lp
   SHGetPathFromIDList lp, Buffer
   FolderLocation = StripNull(Buffer)
   'Free the PIDL
   CoTaskMemFree lp

End Function
Private Sub LoadTree6()

   Const Shell32$ = "Shell32.Dll"

   On Error GoTo PROC_ERR

   'Use API (MUCH faster than scripting)
'------------------------------
   Dim FirstFixed  As Integer
   Dim MaxPwr      As Integer
   Dim Pwr         As Integer
'------------------------------
   Dim DrvBitMask  As Long
   Dim DriveType   As Long
'------------------------------
   Dim MyDrive     As String
   Dim MyPic       As String
   Dim MyKey       As String
'------------------------------
   Dim nod1        As Node
   Dim si          As SHFILEINFO
   Dim RC          As RECT
'------------------------------

   TV.ImageList = ImageList1    ' Initialize ImageList.
   m_hwndTV = TV.hwnd
  ' Establish the distance in which auto-scrolling happens within
  ' the TreeView's client area (we need a root item for these calls)
  If TreeView_GetItemRect(m_hwndTV, TreeView_GetRoot(m_hwndTV), RC, True) Then
    m_cxyAutoScroll = (RC.Bottom - RC.TOp) * 2
  Else
    m_cxyAutoScroll = 32
  End If
  ' Initialize the auto expand and auto scroll timers
  ' already set in design properties
  ' tmrAutoExpand.Enabled = False
  ' tmrAutoExpand.Interval = 1000
  ' tmrAutoScroll.Enabled = False
  ' tmrAutoScroll.Interval = 100
  ' Store thet distance the cursor moves to initiate dragging.
  m_szDrag.cx = GetSystemMetrics(SM_CXDRAG)
  m_szDrag.cy = GetSystemMetrics(SM_CYDRAG)

'Private Const DRIVE_UNKNOWN       As Long = 0
'Private Const DRIVE_NO_ROOT_DIR   As Long = 1
'Private Const DRIVE_REMOVABLE     As Long = 2
'Private Const DRIVE_FIXED         As Long = 3
'Private Const DRIVE_REMOTE        As Long = 4
'Private Const DRIVE_CDROM         As Long = 5
'Private Const DRIVE_RAMDISK       As Long = 6

   MyDrive = GetResourceStringFromFile(Shell32, 4162) 'Desktop
   Set nod1 = TV.Nodes.Add(, , Desktop, MyDrive, "dt")
   '-----
   'MyDrive = GetResourceStringFromFile(Shell32, 9100) 'My Documents
   m_MyDocs = FolderLocation(CSIDL_PERSONAL)
   Set nod1 = TV.Nodes.Add(Desktop, tvwChild, QualifyPath(m_MyDocs), Mid(m_MyDocs, 4), "md")
   If hasSubDirectory(m_MyDocs) Then
      TV.Nodes.Add nod1, tvwChild
   End If
   '-----
   Set nod1 = TV.Nodes.Add(Desktop, tvwChild, "Ftp", "Ftp Client", "rte")
   '-----
   MyDrive = GetResourceStringFromFile(Shell32, 9216) 'My Computer
   Set nod1 = TV.Nodes.Add(Desktop, tvwChild, MyComputer, MyDrive, "mc")
   '-----
   DrvBitMask = GetLogicalDrives()
   ' DrvBitMask is a bitmask representing
   ' available disk drives. Bit position 0
   ' is drive A, bit position 2 is drive C, etc.
   ' If function fails, return value is zero.
   If DrvBitMask Then
    ' Get & search each available drive
      MaxPwr = Int(Log(DrvBitMask) / Log(2))   ' a little math...
      For Pwr = 0 To MaxPwr
         If 2 ^ Pwr And DrvBitMask Then
            MyDrive = Chr$(65 + Pwr) & ":\"
            DriveType = GetDriveType(MyDrive)
            Select Case DriveType
               Case 0, 1: MyPic = "dl"
               Case 2:
                  If Pwr < 2 Then 'A or B (Diskette)
                     MyPic = "f35"
                  Else 'other Removable
                     MyPic = "rem"
                  End If
               Case 3: MyPic = "hd"
               Case 4: MyPic = "rte"
               Case 5: MyPic = "cd"
               Case 6: MyPic = "ram"
            End Select
            'Get Drive DisplayName.
            SHGetFileInfo MyDrive, 0&, si, Len(si), SHGFI_DISPLAYNAME
            Set nod1 = TV.Nodes.Add(MyComputer, tvwChild, MyDrive, si.szDisplayName, MyPic)
            If (FirstFixed = 0) And (DriveType = 3) Then
               FirstFixed = TV.Nodes.Count
            End If
            TV.Nodes.Add nod1, tvwChild
         End If
      Next
   End If
   'Add Control Panel
   MyDrive = GetResourceStringFromFile(Shell32, 4161)
   Set nod1 = TV.Nodes.Add(MyComputer, tvwChild, "ControlPanel", MyDrive, "cp")
   TV.Nodes.Add nod1, tvwChild

   'expand first fixed drive
   Set nod1 = TV.Nodes(FirstFixed)
   nod1.Expanded = True
   nod1.EnsureVisible
   'ensure first entry (Desktop) is visible
   Set nod1 = TV.Nodes(1) 'Desktop
   nod1.EnsureVisible
   TV.Refresh
   Set nod1 = Nothing
   
PROC_EXIT:
  Exit Sub
PROC_ERR:
  If ErrMsgBox("LoadTree6") = vbRetry Then Resume Next

End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,1,0,0
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
   hwnd = TV.hwnd   '   m_hWnd
End Property
Public Property Let hwnd(ByVal New_hWnd As Long)
   If Ambient.UserMode Then Err.Raise 382
   m_hWnd = New_hWnd
   PropertyChanged "hWnd"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TV,TV,-1,Appearance
Public Property Get Appearance() As AppearanceConstants
Attribute Appearance.VB_Description = "Returns/sets whether or not controls, Forms or an MDIForm are painted at run time with 3-D effects."
   Appearance = TV.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
   TV.Appearance() = New_Appearance
   PropertyChanged "Appearance"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TV,TV,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
   BorderStyle = TV.BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConstants)
   TV.BorderStyle() = New_BorderStyle
   PropertyChanged "BorderStyle"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TV,TV,-1,Checkboxes
Public Property Get Checkboxes() As Boolean
Attribute Checkboxes.VB_Description = "Returns/sets a value which determines if the control displays a checkbox next to each item in the tree."
   Checkboxes = TV.Checkboxes
End Property
Public Property Let Checkboxes(ByVal New_Checkboxes As Boolean)
   TV.Checkboxes() = New_Checkboxes
   PropertyChanged "Checkboxes"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TV,TV,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
   Enabled = TV.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
   TV.Enabled() = New_Enabled
   PropertyChanged "Enabled"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TV,TV,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
   Set Font = TV.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
   Set TV.Font = New_Font
   PropertyChanged "Font"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TV,TV,-1,FullRowSelect
Public Property Get FullRowSelect() As Boolean
Attribute FullRowSelect.VB_Description = "Returns/sets a value which determines if the entire row of the selected item is highlighted and clicking anywhere on an item's row causes it to be selected."
   FullRowSelect = TV.FullRowSelect
End Property

Public Property Let FullRowSelect(ByVal New_FullRowSelect As Boolean)
   TV.FullRowSelect() = New_FullRowSelect
   PropertyChanged "FullRowSelect"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TV,TV,-1,HideSelection
Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "Determines whether the selected item will display as selected when the TreeView loses focus"
   HideSelection = TV.HideSelection
End Property

Public Property Let HideSelection(ByVal New_HideSelection As Boolean)
   TV.HideSelection() = New_HideSelection
   PropertyChanged "HideSelection"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TV,TV,-1,HotTracking
Public Property Get HotTracking() As Boolean
Attribute HotTracking.VB_Description = "Returns/sets a value which determines if items are highlighted as the mousepointer passes over them."
   HotTracking = TV.HotTracking
End Property

Public Property Let HotTracking(ByVal New_HotTracking As Boolean)
   TV.HotTracking() = New_HotTracking
   PropertyChanged "HotTracking"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TV,TV,-1,LabelEdit
Public Property Get LabelEdit() As LabelEditConstants
Attribute LabelEdit.VB_Description = "Returns/sets a value that determines if a user can edit the label of a ListItem or Node object."
   LabelEdit = TV.LabelEdit
End Property

Public Property Let LabelEdit(ByVal New_LabelEdit As LabelEditConstants)
   TV.LabelEdit() = New_LabelEdit
   PropertyChanged "LabelEdit"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TV,TV,-1,LineStyle
Public Property Get LineStyle() As TreeLineStyleConstants
Attribute LineStyle.VB_Description = "Returns/sets the style of lines displayed between Node objects."
   LineStyle = TV.LineStyle
End Property

Public Property Let LineStyle(ByVal New_LineStyle As TreeLineStyleConstants)
   TV.LineStyle() = New_LineStyle
   PropertyChanged "LineStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TV,TV,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
   Set MouseIcon = TV.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
   Set TV.MouseIcon = New_MouseIcon
   PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TV,TV,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
   MousePointer = TV.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
   TV.MousePointer() = New_MousePointer
   PropertyChanged "MousePointer"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TV,TV,-1,PathSeparator
Public Property Get PathSeparator() As String
Attribute PathSeparator.VB_Description = "Returns/sets the delimiter string used for the path returned by the FullPath property."
   PathSeparator = TV.PathSeparator
End Property
Public Property Let PathSeparator(ByVal New_PathSeparator As String)
   TV.PathSeparator() = New_PathSeparator
   PropertyChanged "PathSeparator"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,*.zip
Public Property Get FileFilter() As String
Attribute FileFilter.VB_Description = "Files filters separated by ';'\r\nExample *.zip;*.cab"
   FileFilter = m_FileFilter
End Property
Public Property Let FileFilter(ByVal New_FileFilter As String)
   m_FileFilter = New_FileFilter
   PropertyChanged "FileFilter"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
   Set Picture = UserControl.Picture
End Property
Public Property Set Picture(ByVal New_Picture As Picture)
   Set UserControl.Picture = New_Picture
   PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get DragDropEnable() As Boolean
   DragDropEnable = m_DragDropEnable
End Property
Public Property Let DragDropEnable(ByVal New_DragDropEnable As Boolean)
   m_DragDropEnable = New_DragDropEnable
   PropertyChanged "DragDropEnable"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
   BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
   UserControl.BackColor() = New_BackColor
   PropertyChanged "BackColor"
End Property
Public Property Get Path() As String
   Path = m_Path
End Property
Public Property Let Path(ByVal New_Path As String)
   Dim W32FD As WIN32_FIND_DATA
   W32FD = FileExistsW32FD(New_Path)
   If W32FD.dwReserved1 Then
      m_Path = New_Path
      PropertyChanged "Path"
      SetNodeVisible
   End If
End Property
Private Sub TV_NodeClick(ByVal Node As Node)
   RaiseEvent NodeClick(Node)
End Sub

Private Sub TV_NodeCheck(ByVal Node As Node)
   RaiseEvent NodeCheck(Node)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TV,TV,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a form or control."
   TV.Refresh
End Sub

