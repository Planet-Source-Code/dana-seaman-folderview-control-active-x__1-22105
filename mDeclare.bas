Attribute VB_Name = "mDeclares"
Option Explicit

Public Const HKEY_CLASSES_ROOT = &H80000000
'Public Const KEY_ALL_ACCESS = &H2003F
Public Const hNull& = 0
Public Const MAX_PATH = 260
Public Const NOERROR = 0
' Difference between day zero for VB dates and Win32 dates
' (or #12-30-1899# - #01-01-1601#)
Private Const rDayZeroBias As Double = 109205#    ' Abs(CDbl(#01-01-1601#))
' 10000000 nanoseconds * 60 seconds * 60 minutes * 24 hours / 10000
' comes to 86400000 (the 10000 adjusts for fixed point in Currency)
Private Const rMillisecondPerDay As Double = 10000000# * 60# * 60# * 24# / 10000#

Public Const LVM_FIRST = &H1000
Public Const LVS_SHAREIMAGELISTS = &H40&
Public Const GWL_STYLE = (-16)
Public Const LVM_SETIMAGELIST = (LVM_FIRST + 3)
'Public Const LVSIL_NORMAL = 0
'Public Const LVSIL_SMALL = 1
Public Const LVIF_IMAGE = &H2
Public Const LVM_SETITEM = (LVM_FIRST + 6)

'Public Const LARGE_ICON As Integer = 32
'Public Const SMALL_ICON As Integer = 16
'Public Const ILD_TRANSPARENT = &H1                                     'Display transparent
'ShellInfo Flags
'Public Const SHGFI_DISPLAYNAME = &H200
'Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_SYSICONINDEX = &H4000                               'System icon index
'Public Const SHGFI_LARGEICON = &H0                                     'Large icon
Public Const SHGFI_SMALLICON = &H1                                     'Small icon
'Public Const SHGFI_SHELLICONSIZE = &H4
'Public Const SHGFI_TYPENAME = &H400
'Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
'        Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
'        Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
Public Const SMALLSYS_SHGFI_FLAGS = SHGFI_SYSICONINDEX Or SHGFI_SMALLICON

Public Const YMDHMS As String = "yyyymmddhhnnss"
Public Const HMS As String = "hhnnss"
'----------------------------------
Public Const ico_ As String = "ico"
Public Const chk_ As String = "chk"
Public Const nam_ As String = "nam"
Public Const ext_ As String = "ext"
Public Const siz_ As String = "siz"
Public Const typ_ As String = "typ"
Public Const mod_ As String = "mod"
Public Const tim_ As String = "tim"
Public Const cre_ As String = "cre"
Public Const acc_ As String = "acc"
Public Const atr_ As String = "atr"
Public Const dos_ As String = "dos"
'----------------------------------
Public Const cmp_ As String = "cmp"
Public Const rat_ As String = "rat"
Public Const crc_ As String = "crc"
Public Const enc_ As String = "enc"
Public Const mtd_ As String = "mtd"
Public Const pth_ As String = "pth"
Public Const com_ As String = "com"
Public Const sig_ As String = "sig"
'----------------------------------
Public Const ace_      As String = "ace"
Public Const cab_      As String = "cab"
Public Const rar_      As String = "rar"
Public Const zip_      As String = "zip"
Public Const stardot   As String = "*."
'----------------------------------
Public Buffer As String * MAX_PATH
Public f_Type As String * 80

Type SHFILEINFO
        hicon As Long                      '  out: icon
        iIcon As Long                      '  out: icon index
        dwAttributes As Long               '  out: SFGAO_ flags
        szDisplayName As String * MAX_PATH '  out: display name (or path)
        szTypeName As String * 80          '  out: type name
End Type
Public SFI As SHFILEINFO
Public Const cbSFI As Long = 12 + MAX_PATH + 80 'size of SFI

Public Type LV_ITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long '(~ ItemData)
'#if (_WIN32_IE >= 0x0300)
    iIndent As Long
'#End If
End Type

   Public lvi As LV_ITEM 'api list item struc


'Modified for Faster Date Conversion & 64-bit NTFS Filesizes
Public Type WIN32_FIND_DATA
   dwFileAttributes  As Long
   ftCreationTime    As Currency
   ftLastAccessTime  As Currency
   ftLastWriteTime   As Currency
   nFileSizeBig      As Currency
   dwReserved0       As Long
   dwReserved1       As Long
   cFileName         As String * 260
   cAlternate        As String * 14
End Type

Public Type Master
   GridFormat  As Long
   Index       As Long
   Filename    As String
   Size        As Long
   CompSize    As Long
   Modified    As Date
   Created     As Date
   Accessed    As Date
   Attr        As Long
   Method      As Long
   flags       As Long
   Encypted    As Boolean
   Crc         As Long
   Sig         As Long
   Path        As String
   Comments    As String
End Type
Public Master As Master
Public Type FTs
   Ext As String
   Type As String
   IconIndex As Long
End Type
Public Enum GridFormat
   gfFiles = 1
   gfFtp = 2
   gfCab = 4
   gface = 8
   gfrar = 16
   gfzip = 32
End Enum
Public Enum SHFolders
    CSIDL_DESKTOP = &H0
    CSIDL_INTERNET = &H1
    CSIDL_PROGRAMS = &H2
    CSIDL_CONTROLS = &H3
    CSIDL_PRINTERS = &H4
    CSIDL_PERSONAL = &H5
    CSIDL_FAVORITES = &H6
    CSIDL_STARTUP = &H7
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_BITBUCKET = &HA
    CSIDL_STARTMENU = &HB
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_DRIVES = &H11
    CSIDL_NETWORK = &H12
    CSIDL_NETHOOD = &H13
    CSIDL_FONTS = &H14
    CSIDL_TEMPLATES = &H15
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    CSIDL_APPDATA = &H1A
    CSIDL_PRINTHOOD = &H1B
    CSIDL_ALTSTARTUP = &H1D '// DBCS
    CSIDL_COMMON_ALTSTARTUP = &H1E '// DBCS
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_COOKIES = &H21
    CSIDL_HISTORY = &H22
End Enum
'NOTE!! Some declares changed to 'As Any' to
'       accomodate Currency as well as Filetime
Declare Function DosDateTimeToFileTime Lib "kernel32" (ByVal wFatDate As Long, ByVal wFatTime As Long, lpFileTime As Any) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
'Declare Function GetLogicalDrives Lib "kernel32" () As Long
Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As Any, lpLocalFileTime As Any) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal uID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
'Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal flags&) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function CharLower Lib "user32" Alias "CharLowerA" (ByVal lpsz As String) As Long
Declare Function CharUpper Lib "user32" Alias "CharUpperA" (ByVal lpsz As String) As Long
'Declare Function SHGetMalloc Lib "shell32" (ppMalloc As IMalloc) As Long
Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Declare Function lstrlenptr Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Sub CopyMemoryLpToStr Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal lpvDest As String, lpvSource As Long, ByVal cbCopy As Long)

Public Function PointerToString(lPtr As Long) As String
Dim lLen As Long
Dim sR As String
    ' Get length of Unicode string to first null
    lLen = lstrlenptr(lPtr)
    ' Allocate a string of that length
    sR = String$(lLen, 0)
    ' Copy the pointer data to the string
    CopyMemoryLpToStr sR, ByVal lPtr, lLen
    PointerToString = sR
End Function
Public Function QualifyPath(ByVal MyString As String) As String
   If Right$(MyString, 1) <> "\" Then
      QualifyPath = MyString & "\"
   Else
      QualifyPath = MyString
   End If
End Function
Public Function GetMyDate(ZipDate As Integer, ZipTime As Integer) As Date
    Dim FTime As Currency 'Makes it much easier to convert
    'Convert the dos stamp into a file time
    DosDateTimeToFileTime CLng(ZipDate), CLng(ZipTime), FTime
    'Filetime to VbDate
    GetMyDate = UTCCurrToVbDate(FTime, False)
End Function
Public Function GetResourceStringFromFile(sModule As String, idString As Long) As String

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

Public Function ErrMsgBox(Msg As String) As Integer
    ErrMsgBox = MsgBox("Error: " & Err.Number & ". " & Err.Description, vbRetryCancel + vbCritical, Msg)
End Function
Public Function UTCCurrToVbDate(ByVal MyCurr As Currency, Optional TooLocal As Boolean = True) As Date
   Dim UTC As Currency
   ' Discrepancy in WIN32_FIND_DATA:
   ' Win2000 correctly reports 0 as 01-01-1980, Win98/ME does not.
   If MyCurr = 0 Then MyCurr = 11960017200000# ' 01-01-1980
   If TooLocal Then
      FileTimeToLocalFileTime MyCurr, UTC
   Else
      UTC = MyCurr
   End If
   UTCCurrToVbDate = (UTC / rMillisecondPerDay) - rDayZeroBias

End Function
Public Function CVC(ByVal Big As Currency) As Currency
    'Swap High/Low of Big
    'NOTE: Stores value as 64-bit integer (up to 8 Exabytes - 1)
    '      Scale * 10000 when retrieving value (VbCurrency)
    CopyMemory ByVal VarPtr(CVC) + 4, Big, 4
    CopyMemory CVC, ByVal VarPtr(Big) + 4, 4
End Function
Public Function DirSpace(sPath As String) As Currency
   Dim Win32Fd As WIN32_FIND_DATA
   Dim lHandle As Long
   Const FILE_ATTRIBUTE_DIRECTORY = &H10
   sPath = QualifyPath(sPath)
   lHandle = FindFirstFile(sPath & "*.*", Win32Fd)
   If lHandle > 0 Then
      Do
         If Asc(Win32Fd.cFileName) <> 46 Then  'skip . and .. entries
            If (Win32Fd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = 0 Then
               DirSpace = DirSpace + CVC(Win32Fd.nFileSizeBig)
            Else 'Recurse
               DirSpace = DirSpace + DirSpace(sPath & StripNull(Win32Fd.cFileName))
            End If
         End If
      Loop While FindNextFile(lHandle, Win32Fd) > 0
   End If
   FindClose (lHandle)

End Function
Public Sub ParseFullPath(ByVal FullPath As String, JustPath As String, JustName As String)
   
   Dim lSlash As Integer
   
   ' Given a full path, parse it and return
   ' the path and file name.
   lSlash = InStrRev(FullPath, "/")
   If lSlash = 0 Then
      lSlash = InStrRev(FullPath, "\")
   End If
   If lSlash > 0 Then
      JustName = Mid$(FullPath, lSlash + 1)
      JustPath = Left$(FullPath, lSlash)
   Else
      JustName = FullPath
      JustPath = ""
   End If

End Sub
Public Function StringToPointer(sStr As String, ByRef ByteArray() As Byte) As Long
    Dim x As Long
    Dim lstrlen As Long
    
    lstrlen = Len(sStr)
    For x = 1 To lstrlen
        ByteArray(x - 1) = AscB(Mid(sStr, x, 1))
    Next
    ByteArray(x - 1) = 0
    StringToPointer = VarPtr(ByteArray(LBound(ByteArray)))
End Function
Public Function StripNull(ByVal StrIn As String) As String
   On Error GoTo ProcedureError
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

ProcedureExit:
  Exit Function
ProcedureError:
     If ErrMsgBox("mDeclares.StripNull") = vbRetry Then Resume Next


End Function
Public Function FolderLocation(lFolder As SHFolders) As String

   Dim lp As Long
   'Get the PIDL for this folder
   SHGetSpecialFolderLocation 0&, lFolder, lp
   SHGetPathFromIDList lp, Buffer
   FolderLocation = StripNull(Buffer)
   'Free the PIDL
   CoTaskMemFree lp

End Function
Public Function FormatSize(ByVal Size As Variant) As String
   'Handles up to 999.9 Yottabytes.

   'MB = 1024 ^ 2 'Megabyte  2^20 or 1048576
   'GB = 1024 ^ 3 'Gigabyte  2^30 or 1073741824
   'TB = 1024 ^ 4 'Terabyte  2^40 or 1099511627776
   'PB = 1024 ^ 5 'Petabyte  2^50 or 1125899906842624
   'EB = 1024 ^ 5 'Exabyte   2^60 or 1152921504606846976
   'ZB = 1024 ^ 6 'Zettabyte 2^70 or 1180591620717411303424
   'YB = 1024 ^ 7 'Yottabyte 2^80 or 1208925819614629174706176
   'Formats as:
   '   #.###
   'or ##.##
   'or ###.#
   Dim Decimals As Integer, Group As Integer, Pwr As Integer
   Dim SizeKb
   Const KB& = 1024
   On Error GoTo PROC_ERR

   If Size < KB Then 'Return bytes
      FormatSize = FormatNumber(Size, 0) & " b"
      ' Vb5 FormatSize = Format(Size, "#,##0 b")
   Else
      SizeKb = Size / KB
      For Pwr = 0 To 23
         If SizeKb < 10 ^ (Pwr + 1) Then    ' Fits our criteria
            Group = Pwr \ 3                 ' Kb(0), Mb(1), etc.
            SizeKb = SizeKb / KB ^ Group    ' Scale to group
            Decimals = 4 - Len(Int(SizeKb)) ' NumDigitsAfterDecimal
            FormatSize = FormatNumber(SizeKb, Decimals) & " " & _
                         Mid("KMGTPEZY", Group + 1, 1) & "b"
            ' Vb5 FormatSize = Format(SizeKb, "#,###." & String(Decimals, 48)) & " " & _
                         Mid("KMGTPEZY", Group + 1, 1) & "b"
            Exit For
         End If
      Next
      If FormatSize = "" Then FormatSize = "Out of bounds"
   End If
    
PROC_EXIT:
  Exit Function
PROC_ERR:
   FormatSize = "Overflow"
End Function

