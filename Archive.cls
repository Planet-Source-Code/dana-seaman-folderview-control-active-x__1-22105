VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Major rewrite 30-May-2001. Handles Ace, Cab, Rar, Zip archives
' in just one class (easily expandable to accomodate other formats)
' Common event now returns Total # of files for use in progress bars.
' Zip and Cab are enumerated using NO 3rd party Dll's.
' Ace and Rar use UnAce.Dll and UnRar.Dll respectively.
'
' Dec 2001. Added network support.

'To Do:
' Complete Zip/Unzip support(Done - not included in this demo)
' Finish UnAce/UnRar decompress.
' Write code to compress Ace/Rar (need Dll's)
' Write compress/decompress for Cab using Cabinet.Dll
' instead of setupapi.dll
'
'Set ArchiveName/ArchiveExt after instantiating class
Public ArchiveName As String  'Compressed FileName to open
Public ArchiveExt As String    'Ext of Archive (ace,cab,rar,zip)
'Returns total # of files - included in all FileFound events
Public FileCount As Long      'Total

Private Type ZipFileCentralHeader
  VersionMadeBy            As Integer
  VersionNeededToExtract   As Integer
  Flag                     As Integer
  CompressionMethod        As Integer
  time                     As Integer
  date                     As Integer
  CRC32                    As Long
  CompressedSize           As Long
  UncompressedSize         As Long
  FileNameLength           As Integer
  ExtraFieldLength         As Integer
  FileCommentLength        As Integer
  DiskNumberStart          As Integer
  InternalAttr             As Integer
  ExternalAttr             As Long
  RelOffsetLocHdr          As Long
  Filename                 As String
  ExtraField               As String
  FileComment              As String
End Type
Private Type ZipFileEndCentralHeader
  DiskNumberThis             As Integer
  DiskNumberCentralDir       As Integer
  CentralDirEntriesThisDisk  As Integer
  CentralDirEntriesTotal     As Integer
  SizeCentralDir             As Long
  CentralDirOffset           As Long
  FileCommentLength          As Integer
  FileComment                As String
End Type
Private Type ZipDigitalSignature
   SignatureSize  As Integer
   signature      As String
End Type

'Private Type ZipFileLocalHeader
'  Version            As Integer
'  Flag               As Integer
'  CompressionMethod  As Integer
'  Time               As Integer
'  Date               As Integer
'  CRC32              As Long
'  CompressedSize     As Long
'  UncompressedSize   As Long
'  FileNameLength     As Integer
'  ExtraFieldLength   As Integer
'  FileName           As String
'  ExtraField         As String
'End Type
Private Type ACEOPENARCHIVEDATA
    Arcname As Long
    OpenMode As Long
    OpenResult As Long
    flags As Long
    Host As Long
    AV As String * 51
    CmtBuf As Long      'Pointer to buffer ??
    CmtBufSize As Long
    CmtSize As Long
    CmtState As Long
    ChangeVolProc As Long
    ProcessDataProc As Long
End Type

Private Type ACEHEADERDATA
    Arcname As String * MAX_PATH
    Filename As String * MAX_PATH
    flags As Long
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
Private Const ACEERR_WRITE = 7
'Private Const ACEERR_CLINE = 8
Private Const ACEERR_CRC = 9
'Private Const ACEERR_OTHER = 10
'Private Const ACEERR_EXISTS = 11
'Private Const ACEERR_END = 128
'Private Const ACEERR_HANDLE = 129
'Private Const ACEERR_CONSTANT = 130
'Private Const ACEERR_NOPASSW = 131
'Private Const ACEERR_METHOD = 132
'Private Const ACEERR_USER = 255

'Const SUCCESS = 0&

Private Const ACEOPEN_LIST = 0
Private Const ACEOPEN_EXTRACT = 1

Private Const ACECMD_SKIP = 0
Private Const ACECMD_TEST = 1
Private Const ACECMD_EXTRACT = 2

'Private Const ACEVOL_REQUEST = 0
'Private Const ACEVOL_OPENED = 1

Private Declare Function ACEOpenArchive Lib "unACE.dll" _
                    (ByRef Archivedata As ACEOPENARCHIVEDATA) As Long
Private Declare Function ACEProcessFile Lib "unACE.dll" _
                    (ByVal hArcData As Long, _
                     ByVal Operation As Long, _
                     ByVal DestPath As String) As Long
Private Declare Function ACECloseArchive Lib "unACE.dll" _
                    (ByVal hArcData As Long) As Long
Private Declare Function ACEReadHeader Lib "unACE.dll" _
                    (ByVal hArcData As Long, _
                     ByRef Headerdata As ACEHEADERDATA) As Long
Private Type CabFileHeader
    signature     As String * 4  ' MSCF (cabinet file signature )
    reserved1     As Long        '
    cbCabinet     As Long        'size of this cabinet file in bytes
    reserved2     As Long        '
    coffFiles     As Long        'offset of the first CFFILE entry
    reserved3     As Long        '
    versionMinor  As Byte        'cabinet file format version, minor
    versionMajor  As Byte        'cabinet file format version, major
    cFolders      As Integer     'number of CFFOLDER entries in this cabinet
    cFiles        As Integer     'number of CFFILE entries in this cabinet
    flags         As Integer     'cabinet file option indicators
    setID         As Integer     'must be the same for all cabinets in a set
    iCabinet      As Integer     'number of this cabinet file in a set
  '  cbCFHeader    As Integer     '(optional) size of per-cabinet reserved area
  '  cbCFFolder    As Byte        '(optional) size of per-folder reserved area
  '  cbCFData      As Byte        '(optional) size of per-datablock reserved area
  '  abReserve     As Byte        '(optional) per-cabinet reserved area
  '  szCabinetPrev As Byte        '(optional) name of previous cabinet file
  '  szDiskPrev    As Byte        '(optional) name of previous disk
  '  szCabinetNext As Byte        '(optional) name of next cabinet file
  '  szDiskNext    As Byte        '(optional) name of next disk
End Type

Private Type CFFOLDER
    coffCabStart  As Long     'offset of the first CFDATA block in this folder
    cCFData       As Integer  'number of CFDATA blocks in this folder
    typeCompress  As Integer  'compression type indicator
End Type

Private Type CFFILE
    uSize            As Long     'uncompressed size of this file in bytes
    uoffFolderStart  As Long     'uncompressed offset of this file in the folder
    iFolder          As Integer  'index into the CFFOLDER area
    date             As Integer  'date stamp for this file
    time             As Integer  'time stamp for this file
    attribs          As Integer  'attribute flags for this file
    'szName is variable length string with Chr$(0) terminator
    'See GetInfo to see how seek is adjusted for block alignment
    szName           As String * 260  'name of this file
End Type

'Would have been nice if the Crc and
'Compressed size were in CFFILE above

Private Type CFDATA
    csum       As Long    'checksum of this CFDATA entry
    cbData     As Integer 'number of compressed bytes in this block
    cbUncomp   As Integer 'number of uncompressed bytes in this block
  '  abReserve  As Byte    '(optional) per-datablock reserved area
  '  ab[cbData] As Byte    'compressed data bytes
End Type
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

Private Type RARHEADERDATA
    Arcname As String * MAX_PATH         ' Contains the zero terminated string of the current archive name.
                                    ' Maybe used to determine the current volume name
    Filename As String * MAX_PATH        ' Contains the zero terminated string of the file name
    flags As Long                   ' Flags
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
Private Const RAR_OM_LIST = 0&           ' Open archive for reading file headers only
Private Const RAR_OM_EXTRACT = 1        ' Open archive for testing and extracting files

' Operation values
Private Const RAR_SKIP = 0&              ' Move to the next file in archive
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
Private Declare Function RARCloseArchive Lib "unrar.dll" _
                (ByVal hArcData As Long) As Long
                
' Read header of file in archive
' parameters:   hAcrData        - contains the archive handle obtained from the
'                                 RAROpenArchive function call
'               HeaderData      - points to RARHeaderData structure
' returns:  0                   - Success
'           ERAR_END_ARCHIVE    - End of archive
'           ERAR_BAD_ARCHIVE    - File header broken
Private Declare Function RARReadHeader Lib "unrar.dll" _
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
Private Declare Function RARProcessFile Lib "unrar.dll" _
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
Public Event FileFound( _
        ByVal Index As Long, _
        ByVal Total As Long, _
        ByVal Filename As String, _
        ByVal ArchiveExt As String, _
        ByVal Modified As Date, _
        ByVal Size As Long, _
        ByVal CompSize As Long, _
        ByVal Method As Long, _
        ByVal Attr As Long, _
        ByVal Path As String, _
        ByVal flags As Long, _
        ByVal Crc As Long, _
        ByVal Comments As String)
        
        
Private Sub GetZip()
    'Copyright 2001 Dana Seaman
    'Rewritten to:
    '1: Get ZipFileEndCentralHeader
    '2: Go direct to ZipFileCentralHeader
    '3: Enumerate the entries
    '4: Add to Listview
    '5: 31Dec2001 Add network support
    '   Replace VB Binary File I/O with API
    
        '<EhHeader>
        On Error GoTo GetZip_Err
        '</EhHeader>

    Dim Sig           As Long
    Dim LenFile       As Long
    Dim Index         As Long
    Dim sPath         As String
    Dim Filename      As String
    Dim Temp          As String * 4096
    Dim MyDate        As Date
    Dim hFile         As Long
    Dim bBuffer(10)   As Byte
    Dim lResult       As Long
    Dim lowbyte       As Long 'low dword of file pointer position
    Dim highbyte      As Long 'high dword of file pointer position
    Dim Ret           As Long
    Dim MyPos         As Long
    '-------------------------------------------------
    'Dim zFile         As ZipFileLocalHeader
    Dim zCentral      As ZipFileCentralHeader
    Dim zEndCentral   As ZipFileEndCentralHeader
    'Dim zSignature    As ZipDigitalSignature
    '-------------------------------------------------
    'Zip Signatures                          'a.k.a.
    'Const LocalFileHeaderSig = &H4034B50    'PK 03 04
    'Const CentralFileHeaderSig = &H2014B50  'PK 01 02
    'Const EndCentralHeaderSig = &H6054B50   'PK 05 06
    'Const DigitalSig = &H5054B50            'PK 05 05
    'Const SpanSig = &H8074B50               'PK 07 08
    Const Offset As Long = 4096

100 hFile = CreateFile(ArchiveName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0&)
102 If hFile = INVALID_HANDLE_VALUE Then Exit Sub
104 LenFile = GetFileSize(hFile, 0)
    If LenFile = 0 Then Exit Sub

110 If LenFile > Offset Then
114    MyPos = LenFile - Offset
    Else
118    MyPos = 1
    End If
    lowbyte = MyPos
    highbyte = 0
    lowbyte = SetFilePointer(hFile, lowbyte, highbyte, FILE_BEGIN)
    ReadFile hFile, ByVal Temp, Offset, Ret, ByVal 0&
122 Sig = InStrRev(Temp, "PK" & Chr$(5) & Chr$(6))
124 If Sig Then
128    lowbyte = MyPos + Sig + 3
130    highbyte = 0
132    lowbyte = SetFilePointer(hFile, lowbyte, highbyte, FILE_BEGIN)
134    With zEndCentral
138       ReadFile hFile, .DiskNumberThis, 2, Ret, ByVal 0&
142       ReadFile hFile, .DiskNumberCentralDir, 2, Ret, ByVal 0&
146       ReadFile hFile, .CentralDirEntriesThisDisk, 2, Ret, ByVal 0&
150       ReadFile hFile, .CentralDirEntriesTotal, 2, Ret, ByVal 0&
154       ReadFile hFile, .SizeCentralDir, 4, Ret, ByVal 0&
158       ReadFile hFile, .CentralDirOffset, 4, Ret, ByVal 0&
         ' Get ZipStream, , .FileCommentLength        'Integer
         ' .FileComment = String$(.FileCommentLength, vbKeySpace)
         ' Get ZipStream, , .FileComment              'String
160      ' Seek ZipStream, .CentralDirOffset + 1
162       lowbyte = .CentralDirOffset '+ 1
164       highbyte = 0
166       lowbyte = SetFilePointer(hFile, lowbyte, highbyte, FILE_BEGIN)
168       FileCount = .CentralDirEntriesThisDisk
       End With
170    For Index = 1 To FileCount
172       With zCentral 'This has all the goodies
174          ReadFile hFile, Sig, 4, Ret, ByVal 0&
176          ReadFile hFile, .VersionMadeBy, 2, Ret, ByVal 0&
178          ReadFile hFile, .VersionNeededToExtract, 2, Ret, ByVal 0&
180          ReadFile hFile, .Flag, 2, Ret, ByVal 0&
182          ReadFile hFile, .CompressionMethod, 2, Ret, ByVal 0&
184          ReadFile hFile, .time, 2, Ret, ByVal 0&
186          ReadFile hFile, .date, 2, Ret, ByVal 0&
188          ReadFile hFile, .CRC32, 4, Ret, ByVal 0&
190          ReadFile hFile, .CompressedSize, 4, Ret, ByVal 0&
192          ReadFile hFile, .UncompressedSize, 4, Ret, ByVal 0&
194          ReadFile hFile, .FileNameLength, 2, Ret, ByVal 0&
196          ReadFile hFile, .ExtraFieldLength, 2, Ret, ByVal 0&
198          ReadFile hFile, .FileCommentLength, 2, Ret, ByVal 0&
200          ReadFile hFile, .DiskNumberStart, 2, Ret, ByVal 0&
202          ReadFile hFile, .InternalAttr, 2, Ret, ByVal 0&
204          ReadFile hFile, .ExternalAttr, 4, Ret, ByVal 0&
206          ReadFile hFile, .RelOffsetLocHdr, 4, Ret, ByVal 0&
208          .Filename = String$(.FileNameLength, vbKeySpace)
210          ReadFile hFile, ByVal .Filename, Len(.Filename), Ret, ByVal 0&
212          If .ExtraFieldLength Then
214             .ExtraField = String$(.ExtraFieldLength, vbKeySpace)
216             ReadFile hFile, ByVal .ExtraField, Len(.ExtraField), Ret, ByVal 0&
             End If
218          If .FileCommentLength Then
220             .FileComment = String$(.FileCommentLength, vbKeySpace)
222             ReadFile hFile, ByVal .FileComment, Len(.FileComment), Ret, ByVal 0&
             End If
          End With
224       ParseFullPath zCentral.Filename, sPath, Filename
226       With zCentral
228          MyDate = GetMyDate(.date, .time)
230          RaiseEvent FileFound(Index, FileCount, Filename, ArchiveExt, MyDate, .UncompressedSize, .CompressedSize, .CompressionMethod, .ExternalAttr, sPath, .Flag, .CRC32, .FileComment)
          End With
       Next
    End If

234 CloseHandle (hFile)

        '<EhFooter>
        Exit Sub

GetZip_Err:
        MsgBox Err.Description & vbCrLf & _
               "in cArchive.GetZip component " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Private Sub GetAce()
On Error GoTo ProcedureError
Dim hArchive As Long
Dim bMultiVolume As Boolean
Dim sPath As String, Filename As String
Dim Index As Long
Dim sFile As String
Dim zCentral As ACEHEADERDATA
Dim MyDate As Date
Dim Comments As String

'----Step thru just to get the total FileCount
   hArchive = OpenACEArchive(ArchiveName, ACEOPEN_LIST, bMultiVolume)
   If hArchive Then
      While ACEReadHeader(hArchive, zCentral) = 0
         sFile = StripNull(zCentral.Filename)
         FileCount = FileCount + 1
         ACEProcessFile hArchive, ACECMD_SKIP, vbNull
      Wend
      ACECloseArchive hArchive
   End If
'-----

   hArchive = OpenACEArchive(ArchiveName, ACEOPEN_LIST, bMultiVolume)
   If hArchive Then
      While ACEReadHeader(hArchive, zCentral) = 0
         sFile = StripNull(zCentral.Filename)
         Index = Index + 1
         ParseFullPath sFile, sPath, Filename
         With zCentral
            MyDate = GetMyDate(.FileDate, .FileTime)
            Comments = PointerToString(.CmtBuf)
            RaiseEvent FileFound(Index, FileCount, Filename, ArchiveExt, MyDate, .UnpSize, .PackSize, .Method, .FileAttr, sPath, .flags, .FileCRC, Comments)
         End With
         ACEProcessFile hArchive, ACECMD_SKIP, vbNull
      Wend
      ACECloseArchive hArchive
   End If


ProcedureExit:
   Exit Sub
ProcedureError:
   If ErrMsgBox("GetAce") = vbRetry Then Resume Next
End Sub

Private Sub GetCab()
'Copyright 2001 Dana Seaman
'   31Dec2001 Add network support
'   Replace VB Binary File I/O with API
    
        '<EhHeader>
        On Error GoTo GetCab_Err
        '</EhHeader>

Dim Sig        As Long
Dim Index      As Long
Dim hFile         As Long
Dim CabStream  As Integer
Dim sPath      As String
Dim Filename   As String
Dim Temp       As String
Dim MyDate     As Date
Dim SeekPos    As Long
Dim lResult    As Long
Dim lowbyte    As Long 'low dword of file pointer position
Dim highbyte   As Long 'high dword of file pointer position
Dim Ret        As Long
'-------------------------------------------------

Dim zCentral      As CabFileHeader
Dim zFile         As CFFILE
Dim zFolder       As CFFOLDER
'-------------------------------------------------

100 hFile = CreateFile(ArchiveName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0&)
102 If hFile = INVALID_HANDLE_VALUE Then Exit Sub

CabStream = FreeFile

Open ArchiveName For Binary As CabStream

Get CabStream, , zCentral
ReadFile hFile, zCentral, 36, Ret, ByVal 0&

If zCentral.signature = "MSCF" Then
   FileCount = zCentral.cFiles
   If FileCount Then
      If zCentral.cFolders Then
         Get CabStream, , zFolder
         ReadFile hFile, zFolder, 8, Ret, ByVal 0&
      End If
      SeekPos = zCentral.coffFiles '+ 1
      Seek CabStream, SeekPos
      lowbyte = SeekPos
      highbyte = 0
      lowbyte = SetFilePointer(hFile, lowbyte, highbyte, FILE_BEGIN)
      For Index = 1 To FileCount
         Get CabStream, , zFile
         ReadFile hFile, zFile, 276, Ret, ByVal 0&
         With zFile
            Temp = StripNull(.szName)
            ParseFullPath Temp, sPath, Filename
            MyDate = GetMyDate(.date, .time)
            RaiseEvent FileFound(Index, FileCount, Filename, ArchiveExt, MyDate, .uSize, 0, zFolder.typeCompress, .attribs, sPath, zCentral.flags, 0, "")
            '260 bytes were read for .szname
            'synchronize SeekPos for next block
            SeekPos = SeekPos + Len(Temp) + 17
            Seek CabStream, SeekPos
            lowbyte = SeekPos
            highbyte = 0
            lowbyte = SetFilePointer(hFile, lowbyte, highbyte, FILE_BEGIN)
         End With
      Next
   End If

End If

    Close CabStream
234 CloseHandle (hFile)
        
        '<EhFooter>
        Exit Sub

GetCab_Err:
        MsgBox Err.Description & vbCrLf & _
               "in cArchive.GetCab component " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Private Sub GetRar()
On Error GoTo ProcedureError

Dim hArchive As Long
Dim bMultiVolume As Boolean
Dim sPath As String, Filename As String
Dim MyDate As Date
Dim zCentral As RARHEADERDATA
Dim Index As Long
Dim Comments As String

'--------Step thru just to get the total FileCount
    hArchive = OpenRARArchive(ArchiveName, RAR_OM_LIST, bMultiVolume)
    If hArchive = 0 Then Exit Sub
    While RARReadHeader(hArchive, zCentral) = 0
        RARProcessFile hArchive, RAR_SKIP, vbNull, 0&
        If (zCentral.flags And &H1000) = 0 Then
            ' file continued flag not set
           FileCount = FileCount + 1
        End If
    Wend
    RARCloseArchive hArchive
'--------
    hArchive = OpenRARArchive(ArchiveName, RAR_OM_LIST, bMultiVolume)
    If hArchive = 0 Then Exit Sub
    
    While RARReadHeader(hArchive, zCentral) = 0
        
        RARProcessFile hArchive, RAR_SKIP, vbNull, 0&
        If (zCentral.flags And &H1000) = 0 Then
            ' file continued flag not set
           Index = Index + 1
           With zCentral
              ParseFullPath StripNull(.Filename), sPath, Filename
              MyDate = GetMyDate(.FileDate, .FileTime)
              Comments = PointerToString(.CmtBuf)
              RaiseEvent FileFound(Index, FileCount, Filename, ArchiveExt, MyDate, .UnpSize, .PackSize, .Method, .FileAttr, sPath, .flags, .FileCRC, Comments)
            End With
        End If
    Wend
           
    RARCloseArchive hArchive

ProcedureExit:
   Exit Sub
ProcedureError:
   If ErrMsgBox("GetRar") = vbRetry Then Resume Next
End Sub
Public Function GetInfo() As Boolean

   Select Case ArchiveExt
      Case ace_: GetAce
      Case cab_: GetCab
      Case rar_: GetRar
      Case zip_: GetZip
   End Select

End Function
Public Function OpenACEArchive(sFileName As String, _
                OpenMode As Long, _
                ByRef bMultiVolume As Boolean) As Long
    Dim hArchive As Long
    Dim tArchiveData As ACEOPENARCHIVEDATA
    Dim ByteArray() As Byte
    
    ReDim ByteArray(0 To Len(sFileName)) As Byte
    tArchiveData.Arcname = StringToPointer(sFileName, ByteArray)
    tArchiveData.OpenMode = OpenMode ' parameter instead of constant
    tArchiveData.CmtBufSize = 0
    hArchive = ACEOpenArchive(tArchiveData)
    If tArchiveData.OpenResult <> 0 Then
        If hArchive <> 0 Then ACECloseArchive hArchive
        OpenACEArchive = 0
    Else
        bMultiVolume = CBool(tArchiveData.flags & &H800)
        OpenACEArchive = hArchive
    End If
End Function

Public Function UnpackACE(sFileName As String, sDestin As String) As Boolean
    Dim hArchive As Long
    Dim tHeaderdata As ACEHEADERDATA
    Dim sFile As String
    Dim bMultiVolume As Boolean
    hArchive = OpenACEArchive(sFileName, ACEOPEN_EXTRACT, bMultiVolume)
    If hArchive = 0 Then Exit Function

    While ACEReadHeader(hArchive, tHeaderdata) = 0
        sFile = StripNull(tHeaderdata.Filename)
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
Public Function OpenRARArchive(sFileName As String, _
                OpenMode As Long, _
                ByRef bMultiVolume As Boolean) As Long
    Dim hArchive As Long
    Dim tArchiveData As RAROPENARCHIVEDATA
    Dim ByteArray() As Byte
    
    ReDim ByteArray(0 To Len(sFileName)) As Byte
    tArchiveData.szArcName = StringToPointer(sFileName, ByteArray)
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

Public Function UnpackRAR(sFileName As String, sDestin As String) As Boolean
    Dim hArchive As Long
    Dim tHeaderdata As RARHEADERDATA
    Dim sFile As String
    Dim bMultiVolume As Boolean
    
    hArchive = OpenRARArchive(sFileName, RAR_OM_EXTRACT, bMultiVolume)
    If hArchive = 0 Then Exit Function

'    RARSetChangeVolProc hArchive, FnPtr(AddressOf ChangeVolProc)
'    RARSetProcessDataProc hArchive, FnPtr(AddressOf ProcessDataProc)
    
    sDestin = QualifyPath(sDestin)
    
    While RARReadHeader(hArchive, tHeaderdata) = 0
        sFile = StripNull(tHeaderdata.Filename)
        Select Case RARProcessFile(hArchive, RAR_EXTRACT, sDestin, 0&)
           Case ERAR_EWRITE
              MsgBox "Write error", vbCritical
              RARCloseArchive hArchive
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
