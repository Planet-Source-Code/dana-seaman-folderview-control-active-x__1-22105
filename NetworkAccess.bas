Attribute VB_Name = "NetworkAccess"
Option Explicit



' public Constants

Public Const INVALID_HANDLE_VALUE = -1
'
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const OPEN_EXISTING = 3
Public Const OPEN_ALWAYS = 4
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
Public Const GENERIC_WRITE = &H40000000
Public Const GENERIC_READ = &H80000000
'
Public Const FILE_BEGIN = 0
Public Const FILE_CURRENT = 1
Public Const FILE_END = 2

Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" _
                         (ByVal lpFileName As String, _
                         ByVal dwDesiredAccess As Long, _
                         ByVal dwShareMode As Long, _
                         lpSecurityAttributes As Any, _
                         ByVal dwCreationDisposition As Long, _
                         ByVal dwFlagsAndAttributes As Long, _
                         ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" _
                         (ByVal hObject As Long) As Long
Public Declare Function ReadFile Lib "kernel32" _
                         (ByVal hFile As Long, _
                         lpBuffer As Any, _
                         ByVal nNumberOfBytesToRead As Long, _
                         lpNumberOfBytesRead As Long, _
                         ByVal lpOverlapped As Long) As Long
Public Declare Function SetFilePointer Lib "kernel32.dll" _
                        (ByVal hFile As Long, _
                        ByVal lDistanceToMove As Long, _
                        lpDistanceToMoveHigh As Long, _
                        ByVal dwMoveMethod As Long) As Long

Public Declare Function GetFileSize Lib "kernel32" _
                        (ByVal hFile As Long, _
                        lpFileSizeHigh As Long) As Long

