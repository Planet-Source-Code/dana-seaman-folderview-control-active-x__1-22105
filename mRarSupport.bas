Attribute VB_Name = "mRarSupport"

Option Explicit

Private Type RAROPENARCHIVEDATA
    szArcName As Long               ' INPUT: Should point to a zero terminated string containing the archive name
    OpenMode As Long                ' INPUT: RAR_OM_LIST - Open archive for reading file headers only
                                    '        RAR_OM_EXTRACT - Open archive for testing and extracting files
    OpenResult As Long              ' OUTPUT: 0                 - Success
                                    '         ERAR_NO_MEMORY    - Not enough memory to initialize data structures
                                    '         ERAR_BAD_DATA     - Archive header broken
                                    '         ERAR_BAD_ARCHIVE  - File is not a valid RAR archive
                                    '         ERAR_EOPEN        - File open error
    szCmtBuf As Long                ' INPUT: Should point to a buffer for archive comments.
                                    '        Maximum comment size is limited to 64 KB. Comment text is zero termintad.
                                    '        If the comment text is larger than the buffer size, the comment text
                                    '        will be trunctated. If szCmtBuf is set to NULL, comments will not be read.
    CmtBufSize As Long              ' INPUT: Should contain size of buffer for archive comments
    CmtSize As Long                 ' OUTPUT: Containing size of comments actually read into the buffer.
                                    '         Cannot exceed CmtBufSize.
    CmtState As Long                ' State:
                                    ' 0                 - absent comments
                                    ' 1                 - Comments read completely
                                    ' ERAR_NO_MEMORY    - Not enough memory to extract comment
                                    ' ERAR_BAD_DATA     - Broken comment
                                    ' ERAR_UNKNOWN_FORMAT - Unknown comment format
                                    ' ERAR_SMALL_BUF    - Buffer too small, comments not completely read
End Type

Public Type RARHEADERDATA
    Arcname As String * MAX_PATH         ' Contains the zero terminated string of the current archive name.
                                    ' Maybe used to determine the current volume name
    FileName As String * MAX_PATH        ' Contains the zero terminated string of the file name
    Flags As Long                   ' Flags
                                    ' bits 7 6 5 4 3 2 1 0
                                    '      0 0 0 0 0 0 0 1  &H1&    - file continued from previous volume
                                    '      0 0 0 0 0 0 1 0  &H2&    - file continues on next volume
                                    '      0 0 0 0 0 1 0 0  &H4&    - file encrypted with password
                                    '      0 0 0 0 1 0 0 0  &H8&    - file comment present
                                    '      0 0 0 1 0 0 0 0  &H10&   - compression of previous files is used
                                    '                                 (solid flag)
                                    '      0 0 0 0 0 0 0 0  &H00&   - dictionary size    64 KB
                                    '      0 0 1 0 0 0 0 0  &H20&   - dictionary size   128 KB
                                    '      0 1 0 0 0 0 0 0  &H40&   - dictionary size   256 KB
                                    '      0 1 1 0 0 0 0 0  &H60&   - dictionary size   512 KB
                                    '      1 0 0 0 0 0 0 0  &H80&   - dictionary size  1024 KB
                                    '      1 0 1 0 0 0 0 0  &HA0&   - reserved
                                    '      1 1 0 0 0 0 0 0  &HC0&   - reserved
                                    '      1 1 1 0 0 0 0 0  &HE0&   - file is directory
    PackSize As Long                ' Packed file size or size of the file part if file was split between volumes
    UnpSize As Long                 ' UnPacked file size
    HostOS As Long                  ' Operating system used for archiving
                                    ' 0 - MS DOS
                                    ' 1 - OS/2
                                    ' 2 - Win32
                                    ' 3 - Unix
    FileCRC As Long                 ' unpacked CRC of file. '
                                    ' It should not be used for file parts which were split between volumes.
    'was  FILETIME As Long                ' Date & Time in standardMS-DOS format
    FileTime As Integer
    FileDate As Integer
                                    ' First 16 bits contain date
                                    '   Bits 0 - 4  : day (1-31)
                                    '   Bits 5 - 8  : month (1=January,12=December)
                                    '   Bits 9 - 15 : year (0=1980)
                                    ' Second 16 bits contain time
                                    '   Bits 0 - 4  : number of seconds divided by two
                                    '   Bits 5 - 10 : number of minutes (0-59)
                                    '   Bits 11 - 15: numer of hours (0-23)
    UnpVer As Long                  ' RAR version required to extract the file
                                    ' It is encoded as 10 * Major version + minor version
    Method As Long                  ' Packing method
    FileAttr As Long                ' File attributes
    CmtBuf As Long                  ' INPUT: Should point to a buffer for file comments.
                                    '        Maximum comment size is limited to 64 KB. Comment text is zero termintad.
                                    '        If the comment text is larger than the buffer size, the comment text
                                    '        will be trunctated. If szCmtBuf is set to NULL, comments will not be read.
    CmtBufSize As Long              ' INPUT: Should contain size of buffer for file comments
    CmtSize As Long                 ' OUTPUT: Containing size of comments actually read into the buffer.
                                    '         Should not exceed CmtBufSize.
    CmtState As Long                ' State:
                                    ' 0                 - absent comments
                                    ' 1                 - Comments read completely
                                    ' ERAR_NO_MEMORY    - Not enough memory to extract comment
                                    ' ERAR_BAD_DATA     - Broken comment
                                    ' ERAR_UNKNOWN_FORMAT - Unknown comment format
                                    ' ERAR_SMALL_BUF    - Buffer too small, comments not completely read
End Type

' Error constants
'Private Const ERAR_END_ARCHIVE = 10&    ' end of archive
'Private Const ERAR_NO_MEMORY = 11&      ' not enough memory to initialize data structures
'Private Const ERAR_BAD_DATA = 12&       ' Archive header broken
'Private Const ERAR_BAD_ARCHIVE = 13&    ' File is not valid RAR archive
'Private Const ERAR_UNKNOWN_FORMAT = 14& ' Unknown comment format
'Private Const ERAR_EOPEN = 15&          ' File open error
'Private Const ERAR_ECREATE = 16&        ' File create error
'Private Const ERAR_ECLOSE = 17&         ' file close error
Private Const ERAR_EREAD = 18&          ' Read error
Private Const ERAR_EWRITE = 19&         ' Write error
' Private Const ERAR_SMALL_BUF = 20&      ' Buffer too small, comment weren't read completely

' OpenMode values
Public Const RAR_OM_LIST = 0&           ' Open archive for reading file headers only
Private Const RAR_OM_EXTRACT = 1        ' Open archive for testing and extracting files

' Operation values
Public Const RAR_SKIP = 0&              ' Move to the next file in archive
                                        ' Warning: If the archive is solid and
                                        ' RAR_OM_EXTRACT mode was set when the archive
                                        ' was opened, the current file will be processed and
                                        ' the operation will be performed slower than a simple seek
'Private Const RAR_TEST = 1&             ' Test the current file and move to the next file in
                                        ' the archive. If the archive was opened with the
                                        ' RAR_OM_LIST mode, the operation is equal to RAR_SKIP
Private Const RAR_EXTRACT = 2&          ' Extract the current file and move to the next file.
                                        ' If the archive was opened with the RAR_OM_LIST mode,
                                        ' the operation is equal to RAR_SKIP

' ChangeVolProc-Mode-parameter-values
'Private Const RAR_VOL_ASK = 0&          ' Required volume is absent. The function should
                                        ' prompt the user and return non-zero value to retry the
                                        ' operation. The function may also specify a new
                                        ' volume name, placing it to ArcName parameter
'Private Const RAR_VOL_NOTIFY = 1&       ' Required volume is successfully opened. This is a
                                        ' notification call and ArcName modification is NOT
                                        ' allowed. The function should return non-zero value
                                        ' to continue or a zero value to terminate operation

' Open RAR archive and allocate memory structures (about 1MB)
' parameters:   ArchiveData     - points to RAROpenArchiveData structure
' returns:  Archive handle or NULL in case of error
Private Declare Function RAROpenArchive Lib "unrar.dll" _
                (ByRef Archivedata As RAROPENARCHIVEDATA) As Long
                
    
' Close RAR archive and release allocated memory.
' Is must be called when archive processing is finished, even if the archive processing
' was stopped due to an error
' parameters:   hAcrData        - contains the archive handle obtained from the
'                                 RAROpenArchive function call
' returns:  0 on success or ERAR_ECLOSE on Archive close error
Public Declare Function RARCloseArchive Lib "unrar.dll" _
                (ByVal hArcData As Long) As Long
                
' Read header of file in archive
' parameters:   hAcrData        - contains the archive handle obtained from the
'                                 RAROpenArchive function call
'               HeaderData      - points to RARHeaderData structure
' returns:  0                   - Success
'           ERAR_END_ARCHIVE    - End of archive
'           ERAR_BAD_ARCHIVE    - File header broken
Public Declare Function RARReadHeader Lib "unrar.dll" _
                (ByVal hArcData As Long, _
                 ByRef Headerdata As RARHEADERDATA) As Long
                 
' Performs action and moves the current position in the archive to the next file.
' Extract or test the current file from the archive opened in RAR_OM_EXTRACT mode.
' If the mode RAR_OM_LIST is set, then a call to this function will simply skip
' the archive position to the next file
' parameters:   hAcrData        - contains the archive handle obtained from the
'                                 RAROpenArchive function call
'               Operation       - RAR_SKIP  : Move to the next file in the archive.
'                                   If the archive is solid and RAR_OM_EXTRACT mode
'                                   was set when the archive was opened, the current
'                                   file will be processed and the operation will be
'                                   performed slower than a simple seek.
'                                 RAR_TEST  : Test the current file and move to the
'                                   next file in the archive. If the archive was opened
'                                   with RAR_OM_LIST mode, the operation is equal to
'                                   RAR_SKIP
'                                 RAR_EXTRACT: Extract the current file and move to
'                                   the next file. If the file was opened with
'                                   RAR_OM_LIST mode, the operation is equal to RAR_SKIP
'               DestPath        - points to a zero-terminated string containing the
'                                 destination directory to which to extract files to.
'                                 If DestPath is equal to NULL it means extract to the
'                                 current directory. This parameters has meaning only
'                                 if DestName is NULL
'               DestName        - points to a string containing the full path and name
'                                 of the file to be extracted of NULL as default. If
'                                 DestName is defined (not NULL) it overrides the original
'                                 file name saved in the archive and DestPath setting
' returns:  0                   - Success
'           ERAR_BAD_DATA       - File CRC error
'           ERAR_BAD_ARCHIVE    - Volume is not a valid RAR archive
'           ERAR_UNKOWN_FORMAT  - Unknown archive format
'           ERAR_EOPEN          - Volume open error
'           ERAR_ECREATE        - File create error
'           ERAR_ECLOSE         - File close error
'           ERAR_EREAD          - Read error
'           ERAR_EWRITE         - Write error
Public Declare Function RARProcessFile Lib "unrar.dll" _
                (ByVal hArcData As Long, _
                 ByVal Operation As Long, _
                 ByVal DestPath As String, _
                 ByVal DestName As Long) As Long

' Set a user-defined function to process volume changing
' parameters:   hAcrData        - contains the archive handle obtained from the
'                                 RAROpenArchive function call
'               lpChangeVolProc - should point to a user-defined "volume change processing" function
'                   This function will be passed two parameters:
'                   ArcName     - points to a zero-terminated name of the next volume
'                   Mode        - The function call mode
'                                 RAR_VOL_ASK   : required volume is absent. The function should prompt the
'                                       user and return a non-zero value to retry or return a zero value to
'                                       terminate the operation. The function may also specify a new volume
'                                       name, placing it to the ArcName parameter
'                                 RAR_VOL_NOTIFY: Required volume is successfully opened. This is a notification
'                                       call and ArcName modification is not allowed. The function should
'                                       return a non-zero value to continue or a zero value to terminate operation.
'                   Other functions of UNRAR.DLL should not be called from the ChangeVolProc function
Private Declare Sub RARSetChangeVolProc Lib "unrar.dll" _
                (ByVal hArcData As Long, _
                 ByVal lpChangeVolProc As Long)
                 
' Set a user-defined function to process unpacked data.
' It may be used to read a file while it is being extracted or tested without
' actual extracting file to disk.
' parameters:   hAcrData        - contains the archive handle obtained from the
'                                 RAROpenArchive function call
'               lpProcessDataProc - should point to a user-defined "data processing" function
'                   This function is called each time when the next data portion is unpacked.
'                   It will be passed two parameters:
'                   Addr        - The address pointing to the unpacked data. The function may refer to the
'                                 the data but must not change it.
'                   Size        - The size of the unpacked data. It is guaranteed only the size will not
'                                 exceed 1 MB (1.048.576 bytes). Any other presumptions may not be correct
'                                 for future implementations of UNRAR.DLL
'                   The function should return a non-zero value to continue process or a zero value to
'                   cancel the archive operation.
'                   Other functions of UNRAR.DLL should not be called from the ChangeVolProc function
Private Declare Sub RARSetProcessDataProc Lib "unrar.dll" _
                (ByVal hArcData As Long, _
                 ByVal lpProcessDataProc As Long)
                 
' Set a password to decrypt files
' It may be used to read a file while it is being extracted or tested without
' actual extracting file to disk.
' parameters:   hAcrData        - contains the archive handle obtained from the
'                                 RAROpenArchive function call
'               Password - should point to a string containing a zero terminated password
Private Declare Sub RARSetPassword Lib "unrar.dll" _
                (ByVal hArcData As Long, _
                 ByVal sPassword As String)
Public Function StringToPointer(sStr As String, ByRef ByteArray() As Byte) As Long
    Dim X As Long
    Dim lstrlen As Long
    
    lstrlen = Len(sStr)
    For X = 1 To lstrlen
        ByteArray(X - 1) = AscB(Mid(sStr, X, 1))
    Next
    ByteArray(X - 1) = 0
    StringToPointer = VarPtr(ByteArray(LBound(ByteArray)))
End Function
Public Function OpenRARArchive(sFilename As String, _
                OpenMode As Long, _
                ByRef bMultiVolume As Boolean) As Long
    Dim hArchive As Long
    Dim tArchiveData As RAROPENARCHIVEDATA
    Dim ByteArray() As Byte
    
    ReDim ByteArray(0 To Len(sFilename)) As Byte
    tArchiveData.szArcName = StringToPointer(sFilename, ByteArray)
    tArchiveData.OpenMode = OpenMode
    tArchiveData.CmtBufSize = 0
    hArchive = RAROpenArchive(tArchiveData)
    If tArchiveData.OpenResult <> 0 Then
        If hArchive <> 0 Then RARCloseArchive hArchive
        OpenRARArchive = 0
    Else
        OpenRARArchive = hArchive
    End If
End Function

Public Function UnpackRAR(sFilename As String, sDestin As String) As Boolean
    Dim hArchive As Long
    Dim tHeaderdata As RARHEADERDATA
    Dim sFile As String
    Dim bMultiVolume As Boolean
    
    hArchive = OpenRARArchive(sFilename, RAR_OM_EXTRACT, bMultiVolume)
    If hArchive = 0 Then Exit Function

'    RARSetChangeVolProc hArchive, FnPtr(AddressOf ChangeVolProc)
'    RARSetProcessDataProc hArchive, FnPtr(AddressOf ProcessDataProc)
    
    sDestin = QualifyPath(sDestin)
    
    While RARReadHeader(hArchive, tHeaderdata) = 0
        sFile = StripNull(tHeaderdata.FileName)
        Select Case RARProcessFile(hArchive, RAR_EXTRACT, sDestin, 0&)
           Case ERAR_EWRITE
              MsgBox "Write error", vbCritical
              ACECloseArchive hArchive
              Exit Function
           Case ERAR_EREAD
                MsgBox "Archive " & sFile & " Read Error.", vbInformation + vbOKOnly
        End Select
        
        If tHeaderdata.FileAttr <> vbDirectory Then
          'Show progress here
        End If
        
        DoEvents
    Wend
    RARCloseArchive hArchive
End Function

Public Function ChangeVolProc(ByRef sArcName As String, ByVal lMode As Long) As Long
    Debug.Print sArcName & " " & CStr(lMode)
    ChangeVolProc = 1&
End Function

Public Function ProcessDataProc(ByVal lAddr As Long, ByVal lSize As Long) As Long
    Debug.Print "SIZE: " & CStr(lSize)
    ProcessDataProc = 1&
End Function


