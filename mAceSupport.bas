Attribute VB_Name = "mAceSupport"
Option Explicit

Private Type ACEOPENARCHIVEDATA
    Arcname As Long
    OpenMode As Long
    OpenResult As Long
    Flags As Long
    Host As Long
    AV As String * 51
    CmtBuf As Long      'Pointer to buffer
    CmtBufSize As Long
    CmtSize As Long
    CmtState As Long
    ChangeVolProc As Long
    ProcessDataProc As Long
End Type

Public Type ACEHEADERDATA
    Arcname As String * MAX_PATH
    FileName As String * MAX_PATH
    Flags As Long
    PackSize As Long
    UnpSize As Long
    FileCRC As Long
   'was FileTime As Long
    FileTime As Integer
    FileDate As Integer
    Method As Long
    QUAL As Long
    FileAttr As Long
    CmtBuf As Long      'Pointer to buffer
    CmtBufSize As Long
    CmtSize As Long
    CmtState As Long
End Type

Private Type typCHANGEVOLPROC
    Arcname As String
    Mode As Long
End Type

Private Type typPROCESSDATAPROC
    Addr As String
    Size As Long
End Type

'Private Const ACEERR_MEM = 1
'Private Const ACEERR_FILES = 2
'Private Const ACEERR_FOUND = 3
'Private Const ACEERR_FULL = 4
'Private Const ACEERR_OPEN = 5
'Private Const ACEERR_READ = 6
Public Const ACEERR_WRITE = 7
'Private Const ACEERR_CLINE = 8
Public Const ACEERR_CRC = 9
'Private Const ACEERR_OTHER = 10
'Private Const ACEERR_EXISTS = 11
'Private Const ACEERR_END = 128
'Private Const ACEERR_HANDLE = 129
'Private Const ACEERR_CONSTANT = 130
'Private Const ACEERR_NOPASSW = 131
'Private Const ACEERR_METHOD = 132
'Private Const ACEERR_USER = 255

'Const SUCCESS = 0&

Public Const ACEOPEN_LIST = 0
Public Const ACEOPEN_EXTRACT = 1

Public Const ACECMD_SKIP = 0
Private Const ACECMD_TEST = 1
Public Const ACECMD_EXTRACT = 2

'Private Const ACEVOL_REQUEST = 0
'Private Const ACEVOL_OPENED = 1

Private Declare Function ACEOpenArchive Lib "unACE.dll" _
                    (ByRef Archivedata As ACEOPENARCHIVEDATA) As Long
Public Declare Function ACEProcessFile Lib "unACE.dll" _
                    (ByVal hArcData As Long, _
                     ByVal Operation As Long, _
                     ByVal DestPath As String) As Long
Public Declare Function ACECloseArchive Lib "unACE.dll" _
                    (ByVal hArcData As Long) As Long
Public Declare Function ACEReadHeader Lib "unACE.dll" _
                    (ByVal hArcData As Long, _
                     ByRef Headerdata As ACEHEADERDATA) As Long


Public Function OpenACEArchive(sFilename As String, _
                OpenMode As Long, _
                ByRef bMultiVolume As Boolean) As Long
    Dim hArchive As Long
    Dim tArchiveData As ACEOPENARCHIVEDATA
    Dim ByteArray() As Byte
    
    ReDim ByteArray(0 To Len(sFilename)) As Byte
    tArchiveData.Arcname = StringToPointer(sFilename, ByteArray)
    tArchiveData.OpenMode = OpenMode ' parameter instead of constant
    tArchiveData.CmtBufSize = 0
    hArchive = ACEOpenArchive(tArchiveData)
    If tArchiveData.OpenResult <> 0 Then
        If hArchive <> 0 Then ACECloseArchive hArchive
        OpenACEArchive = 0
    Else
        bMultiVolume = CBool(tArchiveData.Flags & &H800)
        OpenACEArchive = hArchive
    End If
End Function

Public Function UnpackACE(sFilename As String, sDestin As String) As Boolean
    Dim hArchive As Long
    Dim tHeaderdata As ACEHEADERDATA
    Dim sFile As String
    Dim bMultiVolume As Boolean
    hArchive = OpenACEArchive(sFilename, ACEOPEN_EXTRACT, bMultiVolume)
    If hArchive = 0 Then Exit Function

    While ACEReadHeader(hArchive, tHeaderdata) = 0
        sFile = StripNull(tHeaderdata.FileName)
        Select Case ACEProcessFile(hArchive, ACECMD_EXTRACT, sDestin)
           Case ACEERR_WRITE
              MsgBox "Could not write file to disk", vbCritical
              ACECloseArchive hArchive
              Exit Function
           Case ACEERR_CRC
              MsgBox "Crc Error on File " & sFile, vbInformation
        End Select
        
        If tHeaderdata.FileAttr <> vbDirectory Then
           'Show progress
        End If
        DoEvents
    Wend
    ACECloseArchive hArchive
End Function

