VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFvLvDemo 
   AutoRedraw      =   -1  'True
   Caption         =   "FolderView Active-X, Listview Demo"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   502
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   725
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin FvLvDemo.Splitter Splitter1 
      Height          =   7635
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   13467
      RatioFromTop    =   0.33
      Child1          =   "FolderView1"
      Child2          =   "ListView1"
      LiveUpdate      =   0   'False
      Begin FvLvDemo.FolderView FolderView1 
         Height          =   7575
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   13361
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "FolderViewListViewDemo.frx":0000
         BackColor       =   12632256
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   7605
         Left            =   3990
         TabIndex        =   2
         Top             =   -15
         Width           =   7740
         _ExtentX        =   13653
         _ExtentY        =   13414
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         PictureAlignment=   5
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
         Picture         =   "FolderViewListViewDemo.frx":001C
      End
   End
   Begin VB.PictureBox Splitter 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   7395
      Left            =   3360
      ScaleHeight     =   7395
      ScaleWidth      =   585
      TabIndex        =   0
      Top             =   0
      Width           =   585
   End
End
Attribute VB_Name = "frmFvLvDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------
Private ArqExt       As String
Private WinDir       As String
Private SysDir       As String
Private TempDir      As String
Private SourcePath   As String
Private sFolder      As String
'Private sFile        As String
Private sName        As String
Private sExtension   As String
Private sSize        As String
Private sType        As String
Private sModified    As String
Private sTime        As String
Private sCreated     As String
Private sAccessed    As String
Private sAttribute   As String
Private sMsDos       As String
Private sNone        As String
Private m_MyDocs     As String
'------------------------------
Private Start        As Long
Private FvFilter     As Variant
Private IsFAT        As Boolean
Private InCab        As Boolean
Private InZip        As Boolean
Private Nodx         As Node
Private TypeNew()    As FTs
'------------------------------
Private WithEvents Archive  As cArchive
Attribute Archive.VB_VarHelpID = -1
'------------------------------
Const MyComputer$ = "MyComputer"
Const Desktop$ = "Desktop"
Private Function NiceCase(ByVal Nam As String) As String
   Dim UNam As String, LNam As String
   On Error GoTo ProcedureError
   UNam = Nam: LNam = Nam
   CharUpper UNam: CharLower LNam
   If Nam = UNam Or Nam = LNam Then
  ' If Nam = UCase$(Nam) Or Nam = LCase$(Nam) Then
      NiceCase = StrConv(Nam, vbProperCase)
   Else
      NiceCase = Nam 'already mixed case so leave alone
   End If

ProcedureExit:
  Exit Function
ProcedureError:
     If ErrMsgBox(Me.Name & ".NiceCase") = vbRetry Then Resume Next
   
End Function

Private Function BinarySearchTypeNew(sExt As String) As Integer

   Dim iLow As Integer
   Dim iHigh As Integer
   Dim iMid As Integer
  
   On Error Resume Next

   BinarySearchTypeNew = -1
   iLow = 1 '0 is reserved
   iHigh = UBound(TypeNew) - LBound(TypeNew)
  
   Do
      iMid = (iLow + iHigh) \ 2
      'always LCase so let's use faster binary compare
      Select Case StrComp(sExt, TypeNew(iMid).Ext, vbBinaryCompare)
         Case -1 '< Less than
            iHigh = iMid - 1
         Case 1  '> Greater than
            iLow = iMid + 1
         Case 0  '= Equal
            BinarySearchTypeNew = iMid
            Exit Do
      End Select
   Loop Until iHigh < iLow

End Function
Private Sub ShellSortTypeNewArray()

  Dim iLowBound As Integer
  Dim iHighBound As Integer
  Dim iX As Integer
  Dim iY As Integer
  Dim Temp As FTs

  On Error GoTo ProcedureError
  
  ' Get array bounds
  iLowBound = LBound(TypeNew)
  iHighBound = UBound(TypeNew)
  ' Get array middle
  iY = (iHighBound - iLowBound + 1) \ 2

  Do While iY > 0
    ' Sort lower portion
    For iX = iLowBound To iHighBound - iY
      If TypeNew(iX).Ext > TypeNew(iX + iY).Ext Then
        Temp = TypeNew(iX)
        TypeNew(iX) = TypeNew(iX + iY)
        TypeNew(iX + iY) = Temp
      End If
    Next iX
    ' Sort upper portion
    For iX = iHighBound - iY To iLowBound Step -1
      If TypeNew(iX).Ext > TypeNew(iX + iY).Ext Then
        Temp = TypeNew(iX)
        TypeNew(iX) = TypeNew(iX + iY)
        TypeNew(iX + iY) = Temp
      End If
    Next iX
    ' Divide array
    iY = iY \ 2
  Loop

ProcedureExit:
   Exit Sub
ProcedureError:
   If ErrMsgBox(Me.Name & ".ShellSortTypeNewArray") = vbRetry Then Resume Next

End Sub





Private Sub ace_FileFound(ByVal Count As Long, ByVal Filename As String, ByVal DateTime As Date, ByVal Size As Variant, ByVal CompSize As Variant, ByVal Method As Long, ByVal Attr As Variant, ByVal Path As String, ByVal flags As Long, ByVal Crc As Long, ByVal Comments As String)
   With Master
     .GridFormat = gface
     .Index = Count
     .Filename = Filename
     .Size = Size
     .Modified = DateTime
     .Created = DateTime
     .Accessed = DateTime
     .Attr = Attr
     .Path = Path
     .CompSize = CompSize
     .Method = Method
     .flags = flags
     .Encypted = (flags And 1) * -1 'Make it Boolean
     .Crc = Crc
     '.Sig = 0
     .Comments = Comments
   End With
   LVAddCommon Master
End Sub

Private Sub cab_FileFound(ByVal Count As Long, ByVal Filename As String, ByVal MyDate As Date, ByVal Size As Variant, ByVal Attr As Variant, ByVal Path As String)

   With Master
      .GridFormat = gfCab
      .Index = Count
      .Filename = Filename
      .Modified = MyDate
      .Size = Size
      .Attr = Attr
      .Path = Path
   End With
   LVAddCommon Master

End Sub


Private Sub Archive_FileFound(ByVal Index As Long, ByVal Total As Long, ByVal Filename As String, ByVal ArchiveExt As String, ByVal Modified As Date, ByVal Size As Long, ByVal CompSize As Long, ByVal Method As Long, ByVal Attr As Long, ByVal Path As String, ByVal flags As Long, ByVal Crc As Long, ByVal Comments As String)
        '<EhHeader>
        On Error GoTo Archive_FileFound_Err
        '</EhHeader>
    Dim sMethod As String, sExt As String, FakePath As String
    Dim fType As String
    Dim MyIcon As Long
    Dim FakeFile As Integer
    Dim Ratio As Single
    Dim Encrypt As Boolean
    Dim Item As ListItem

       On Error GoTo ProcedureError
100    If Index = Total Then
102       ArqExt = ArchiveExt
       End If
104    sExt = GetExt(Filename)
106    If LenB(sExt) Then
108       FakeFile = FreeFile
          'Create fake 0 byte file in TempDir
          'Ex: "C:\Windows\Temp\~FileName.Ext"
110       FakePath = QualifyPath(TempDir) & "~" & Filename
112       Open FakePath For Binary As FakeFile
114       Close FakeFile
116       fType = GetFileType(sExt, FakePath, MyIcon)
118       lvi.iImage = MyIcon 'index in System ImageList
120       Kill FakePath
       End If
122    Set Item = ListView1.ListItems.Add()
124    Item.SubItems(LVIdx(nam_)) = Filename
126    lvi.iItem = Item.Index - 1 ' adjusts to 0-based
128    lvi.mask = LVIF_IMAGE 'set image mask
130    SendMessage ListView1.hwnd, LVM_SETITEM, 0&, lvi 'Assign
132    Item.SubItems(LVIdx(ext_)) = sExt
134    Item.SubItems(LVIdx(siz_)) = FormatNumber$(Size, 0)
136    Item.SubItems(LVIdx("siz2")) = Right$(String(10, 48) & Size, 10)
138    Item.SubItems(LVIdx(typ_)) = fType
140    Item.SubItems(LVIdx(mod_)) = FormatDateTime(Modified, vbShortDate)
142    Item.SubItems(LVIdx("mod2")) = Format$(Modified, YMDHMS)
144    Item.SubItems(LVIdx(tim_)) = FormatDateTime(Modified, vbLongTime)
146    Item.SubItems(LVIdx("tim2")) = Format$(Modified, HMS)
      
       'If .GridFormat Then 'ace, cab, rar, zip
148       Item.SubItems(LVIdx(cmp_)) = FormatNumber$(CompSize, 0)
150       Item.SubItems(LVIdx("cmp2")) = Right$(String(10, 48) & CompSize, 10)
          'Trap division by zero
152       If Size Then
154          Ratio = 1 - CompSize / Size
             'Don't allow negative values (per PkZip/WinZip)
             'Occurs on stored+encrypted files
156          If Ratio < 0 Then Ratio = 0
          Else
158          Ratio = 0
          End If
          'Ratio is single. Format as desired
160       Item.SubItems(LVIdx(rat_)) = Format$(Ratio, "00.0%")
162       Item.SubItems(LVIdx(cre_)) = FormGenDatTim(Modified)
164       Item.SubItems(LVIdx("cre2")) = Format$(Modified, YMDHMS)
166       Item.SubItems(LVIdx(acc_)) = FormGenDatTim(Modified)
168       Item.SubItems(LVIdx("acc2")) = Format$(Modified, YMDHMS)

170       Select Case ArchiveExt
             Case ace_
                sMethod = MethodVerboseAce(Method, flags)
172             Encrypt = (flags And 4) * -1
174          Case cab_
                Select Case Method
                   Case 0: sMethod = "None"
                   Case 1: sMethod = "MsZip"
                   Case 2: sMethod = "Lzx"
                End Select
182             Encrypt = False
184          Case rar_
                sMethod = MethodVerboseRar(Method, flags)
                'Flag bit 2 is Encryption True/False
186             Encrypt = (flags And 4) * -1
188          Case zip_
                sMethod = MethodVerboseZip(Method, flags)
190             Encrypt = (flags And 1) * -1
          End Select
192       Item.SubItems(LVIdx(mtd_)) = sMethod
194       Item.SubItems(LVIdx(enc_)) = Encrypt
196       Item.SubItems(LVIdx(crc_)) = Hex$(Crc)
          'Digital Signature extract not yet coded
198       Item.SubItems(LVIdx(sig_)) = "na"
200       Item.SubItems(LVIdx(pth_)) = Path
202       Item.SubItems(LVIdx(com_)) = Comments
      ' End If
204    Item.SubItems(LVIdx(atr_)) = GetAttrString(Attr)
       'Save the item number for
       'other operations
206    Item.Tag = Index
'208    ProgressPanel Index, Total
       'TotalSize = TotalSize + .Size
'210    If Index = Total Then Preview.Cls
ProcedureExit:
      Exit Sub
ProcedureError:
212      If ErrMsgBox(Me.Name & "Archive_FileFound") = vbRetry Then Resume Next
   
        '<EhFooter>
        Exit Sub

Archive_FileFound_Err:
        Select Case ErrMsgBox("FolderviewDemo.frmFolderviewDemo.Archive_FileFound")
           Case vbAbort
              Screen.MousePointer = vbDefault
              Exit Sub
           Case vbRetry
              Resume
           Case vbIgnore
              Resume Next
       End Select
        '</EhFooter>
End Sub

Private Sub FolderView1_NodeClick(ByVal Node As MSComctlLib.Node)
        '<EhHeader>
        On Error GoTo FolderView1_NodeClick_Err
        '</EhHeader>
       Dim Path As String, sExt As String
       Dim Start As Long
   
100    Set Nodx = Node
'102    Tip.Hide
'104    ShowTip = False
       'Reset array of Ext/Type/SysIconIndex
       'Retaining Folder entry if any
106    ReDim Preserve TypeNew(0)
   
108    Select Case Node.Key
          Case "Ftp"
110          ListView1.ListItems.Clear
112       Case "Desktop"
114          LoadFiles QualifyPath(FolderLocation(CSIDL_DESKTOP))
116       Case "MyComputer"
      
118       Case "MyDocuments"
120          LoadFiles QualifyPath(FolderLocation(CSIDL_PERSONAL))
122       Case "ControlPanel"
124          Shell "rundll32.exe shell32.dll,Control_RunDLL", vbNormalFocus
126       Case Else
128          Path = BuildFullPath(Node)
130          sExt = GetExt(Node.Text)
132          Start = GetTickCount()
'134          ShowTip = True
136          Select Case sExt
                Case ace_, cab_, rar_, zip_
                   SourcePath = Path
138                InZip = True
'140                Tip.MouseNotify FolderView1.hWnd, tipMouseMove
                   'LoadStart
142                Screen.MousePointer = vbHourglass
144                LVColumnHeaders
146                Set Archive = New cArchive
148                Archive.ArchiveName = Path
150                Archive.ArchiveExt = sExt
152                Archive.GetInfo
154                LoadCleanup 1
156                Me.Caption = Path
158             Case Else
160                SourcePath = QualifyPath(Path)
'162                Tip.MouseNotify FolderView1.hWnd, tipMouseMove
164                LoadFiles (QualifyPath(Path))
             End Select
       End Select
        '<EhFooter>
        Exit Sub

FolderView1_NodeClick_Err:
        Select Case ErrMsgBox("FolderviewDemo.frmFolderviewDemo.FolderView1_NodeClick")
           Case vbAbort
              Screen.MousePointer = vbDefault
              Exit Sub
           Case vbRetry
              Resume
           Case vbIgnore
              Resume Next
       End Select
        '</EhFooter>
End Sub

Private Sub Form_Load()

   ' Copyright 2001 Dana Seaman, Natal, Brazil
   ' E-Mail:  dseaman@ieg.com.br
   
   Const Shell32$ = "Shell32.Dll"
   FvFilter = Split(LCase(FolderView1.FileFilter), ";")
   m_MyDocs = FolderLocation(CSIDL_PERSONAL)
   'If Not AssignSysIL Then MsgBox "Error in AssignSysIl"
 
'------------------------------
   'Get Win, Sys, & Temp directory paths
   WinDir = Left$(Buffer, GetWindowsDirectory(Buffer, MAX_PATH))
   SysDir = Left$(Buffer, GetSystemDirectory(Buffer, MAX_PATH))
   TempDir = Left$(Buffer, GetTempPath(MAX_PATH, Buffer))
'---- Rip resource strings from Windows Dll's ----
   sFolder = GetResourceStringFromFile(Shell32, 4131) '"(" & GetResourceStringFromFile(Shell32, 4131) & ")"
   'sFile = GetResourceStringFromFile(Shell32, 4130)
   sName = GetResourceStringFromFile(Shell32, 8976)
   sExtension = StrConv(ext_, vbProperCase)
   sSize = GetResourceStringFromFile(Shell32, 8978)
   sType = GetResourceStringFromFile(Shell32, 8979)
   sModified = GetResourceStringFromFile(Shell32, 8980)
   sTime = GetResourceStringFromFile("Intl.Cpl", 25)
   sCreated = GetResourceStringFromFile(Shell32, 8996)
   sAccessed = GetResourceStringFromFile(Shell32, 8997)
   sAttribute = GetResourceStringFromFile(Shell32, 8987)
   sMsDos = "MsDos 8.3"
   sNone = GetResourceStringFromFile(Shell32, 9808)
'-------------------------------
   FolderView1.Enumerate

   LVColumnHeaders
   ListView1.Visible = True
   ListView1.Refresh
   ReDim Preserve TypeNew(0) 'init array of Ext/Type/IconIdx

   'FolderView1.Path = "C:\_AceRarSamples"
   'LoadFiles (FolderView1.Path)


End Sub
Sub LoadFiles(ByVal Path As String)
       
   On Error GoTo ProcedureError
   Dim Win32Fd As WIN32_FIND_DATA
   Dim lHandle As Long
   Dim Item As ListItem
   Dim MyName As String
   Dim sExt As String
   Dim MyDate As Date
   Dim MySize As Currency
   Dim MyIcon As Long
   Dim Start As Long
   Dim MyCount As Long
   
   Const MustGet$ = "exe|ico|lnk|pif|cur"
   
   Start = GetTickCount()
   Screen.MousePointer = vbHourglass
   InZip = False
   SourcePath = QualifyPath(Path)
   LVColumnHeaders

   IsFAT = CheckFAT
   lHandle = FindFirstFile(SourcePath & "*.*", Win32Fd)
   If lHandle > 0 Then
      Do
         If Asc(Win32Fd.cFileName) <> 46 Then  'skip . and .. entries
            MyName = StripNull(Win32Fd.cFileName)
            Set Item = ListView1.ListItems.Add()
            Item.SubItems(LVIdx(nam_)) = NiceCase(MyName)
            sExt = GetExt(MyName)
            Item.SubItems(LVIdx(ext_)) = sExt
            If Win32Fd.dwFileAttributes And vbDirectory Then
               Item.SubItems(LVIdx(typ_)) = sFolder 'from Registry
               If TypeNew(0).Type <> sFolder Then
                  'Get/Store Folder Icon (Only once)
                  SHGetFileInfo Path & MyName, 0&, SFI, cbSFI, SMALLSYS_SHGFI_FLAGS
                  TypeNew(0).Type = sFolder
                  TypeNew(0).IconIndex = SFI.iIcon
               End If
               lvi.iImage = TypeNew(0).IconIndex
            Else
               MySize = CVC(Win32Fd.nFileSizeBig) * 10000
               Item.SubItems(LVIdx(siz_)) = FormatSize(MySize)
               Item.SubItems(LVIdx("siz2")) = Right(String(10, 48) & MySize, 10)
               Item.SubItems(LVIdx(typ_)) = GetFileType(sExt, Path & MyName, MyIcon)
               If Len(sExt) = 3 And InStr(MustGet, sExt) Then
                  'Always obtain Icons for exe,ico,lnk,pif,cur
                  SHGetFileInfo Path & MyName, 0&, SFI, cbSFI, SMALLSYS_SHGFI_FLAGS
                  lvi.iImage = SFI.iIcon
               Else 'Use associated Icon
                  lvi.iImage = MyIcon 'index in system imagelist
               End If
            End If
            lvi.iItem = Item.Index - 1 'item index (-1 adjusts to 0-based api index)
            lvi.mask = LVIF_IMAGE 'just setting image index
            'assign item's image index via api...
            SendMessage ListView1.hwnd, LVM_SETITEM, 0&, lvi
            
            MyDate = UTCCurrToVbDate(Win32Fd.ftLastWriteTime)
            Item.SubItems(LVIdx(mod_)) = FormatDateTime(MyDate, vbShortDate)
            Item.SubItems(LVIdx("mod2")) = Format(MyDate, YMDHMS)
            Item.SubItems(LVIdx(tim_)) = FormatDateTime(MyDate, vbLongTime)
            Item.SubItems(LVIdx("tim2")) = Format(MyDate, HMS)
            MyDate = UTCCurrToVbDate(Win32Fd.ftCreationTime)
            'Don't use VbGeneralDate since it ignores 00:00:00
            Item.SubItems(LVIdx(cre_)) = FormGenDatTim(MyDate)
            Item.SubItems(LVIdx("cre2")) = Format(MyDate, YMDHMS)
            MyDate = UTCCurrToVbDate(Win32Fd.ftLastAccessTime)
            If IsFAT Then ' FAT (Just date)
               Item.SubItems(LVIdx(acc_)) = FormatDateTime(MyDate, vbShortDate)
            Else 'NTFS (Date + Time)
               Item.SubItems(LVIdx(acc_)) = FormGenDatTim(MyDate)
            End If
            Item.SubItems(LVIdx("acc2")) = Format(MyDate, YMDHMS)
            Item.SubItems(LVIdx(atr_)) = GetAttrString(Win32Fd.dwFileAttributes)
            'Get Filename MsDos 8.3
            If InStr(Win32Fd.cAlternate, vbNullChar) = 1 Then
               CharUpper MyName
               Item.SubItems(LVIdx(dos_)) = MyName
            Else
               Item.SubItems(LVIdx(dos_)) = StripNull(Win32Fd.cAlternate)
            End If
         End If
         MyCount = MyCount + 1
         If MyCount Mod 50 = 0 Then
            ShowProgress Start, MyCount, Path
         End If
      Loop While FindNextFile(lHandle, Win32Fd) > 0
   End If
   FindClose lHandle
'------------
   LoadCleanup 3
   ShowProgress Start, MyCount, Path

ProcedureExit:
   Exit Sub
ProcedureError:
   If ErrMsgBox(Me.Name & ".LoadFiles") = vbRetry Then Resume Next
    
End Sub
Private Sub ShowProgress(Start, Count, Path)
   Me.Caption = Format((GetTickCount() - Start) / 1000, "#,##0.00") & " seconds, " & _
                Count & " Objects in " & Path
End Sub

  
Private Function GetFileType(ByVal sExt As String, ByVal FullPath As String, ByRef MyIcon As Long) As String
   On Error GoTo ProcedureError
   Dim sName As String
   Dim lRegKey As Long, L4 As Long
   
   If sExt <> "" Then
      'NOTE: Array must be sorted for binary search
      L4 = BinarySearchTypeNew(sExt)
      If L4 <> -1 Then
         GetFileType = TypeNew(L4).Type
         MyIcon = TypeNew(L4).IconIndex
         Exit Function
      End If
      'Not a duplicate so get info from registry
      If RegOpenKey(HKEY_CLASSES_ROOT, ByVal "." & sExt, lRegKey) = 0 Then
         'Get type of file (Not to be confused with actual FileType )
         RegQueryValueEx lRegKey, ByVal "", 0&, 1, ByVal Buffer, MAX_PATH
         sName = StripNull(Buffer)
         RegCloseKey lRegKey
         If Len(sName) Then
            'Get FileType
            If RegOpenKey(HKEY_CLASSES_ROOT, sName, lRegKey) = 0 Then
               RegQueryValueEx lRegKey, ByVal "", 0&, 1, ByVal f_Type, 80
               GetFileType = StripNull(f_Type)
               RegCloseKey lRegKey
            End If
         End If
      End If
      'Bump array and add new extension/type
      L4 = UBound(TypeNew()) + 1
      ReDim Preserve TypeNew(L4)
      TypeNew(L4).Ext = sExt
      If GetFileType = "" Then 'No associated type
         GetFileType = sNone 'was sFile & " " & UCase$(sExt)
         TypeNew(L4).IconIndex = 0
      Else 'New Ext, get this Icon
         SHGetFileInfo FullPath, 0&, SFI, cbSFI, SMALLSYS_SHGFI_FLAGS
         TypeNew(L4).IconIndex = SFI.iIcon  'index in system imagelist
      End If
      TypeNew(L4).Type = GetFileType
      MyIcon = TypeNew(L4).IconIndex
      ShellSortTypeNewArray 'So we can use a binary search
   End If
   
ProcedureExit:
  Exit Function

ProcedureError:
     If ErrMsgBox(Me.Name & ".GetFileType") = vbRetry Then Resume Next

End Function
Private Function GetExt(ByVal Name As String) As String
   On Error GoTo ProcedureError
   Dim j As Integer
   j = InStrRev(Name, ".")
   If j > 0 And j < (Len(Name)) Then
      GetExt = Mid$(Name, j + 1)
      CharLower GetExt
   End If

ProcedureExit:
  Exit Function
ProcedureError:
     If ErrMsgBox(Me.Name & ".GetExt") = vbRetry Then Resume Next

End Function

Private Function MethodVerboseZip(ByVal Method, ByVal BitFlag) As String
   On Error Resume Next
   'Conforms to PkZip 2.04g Specifications
'Methods are
'0    Stored (None)
'1    Shrunk
'2-5  Reduced:1,2,3,4
'(For Method 6 - Imploding)
' general purpose bit flag: (2 bytes)
'Bit 1: If the compression method used was type 6,
'       Imploding, then this bit, if set, indicates
'       an 8K sliding dictionary was used.  If clear,
'       then a 4K sliding dictionary was used.
'Bit 2: If the compression method used was type 6,
'       Imploding, then this bit, if set, indicates
'       3 Shannon-Fano trees were used to encode the
'       sliding dictionary output.  If clear, then 2
'       Shannon-Fano trees were used.
'6    Imploded:8kDict/4kDict:3Tree/2Tree
'7    Tokenized
'(For Method 8 - Deflating)
' general purpose bit flag: (2 bytes)
'Bit 2  Bit 1
'  0      0    Normal (-en) compression option was used.
'  0      1    Maximum (-ex) compression option was used.
'  1      0    Fast (-ef) compression option was used.
'  1      1    Super Fast (-es) compression option was used.
'8    Deflated:N,X,F,S
'9    EnhDefl
'10   ImplDCL
   
   BitFlag = (BitFlag \ 2) And 3 'Isolate bits 1, 2

   Select Case Method
      'Since deflated is the most common check for it first
      Case 8
         MethodVerboseZip = "Deflated:" & Choose(BitFlag + 1, "N", "X", "F", "S")
      Case 0
         MethodVerboseZip = "Stored"
      Case 1
         MethodVerboseZip = "Shrunk"
      Case 2 To 5
         MethodVerboseZip = "Reduced:" & Method - 1
      Case 6
         MethodVerboseZip = "Imploded:" & Choose(BitFlag + 1, "8KDict:2Tree", "4KDict:2Tree", "8KDict:3Tree", "4KDict:3Tree")
      Case 7
         MethodVerboseZip = "Tokenized"
      Case 9
         MethodVerboseZip = "EnhDef"
      Case 10
         MethodVerboseZip = "ImplDCL"
      Case Else
         MethodVerboseZip = "Unknown"
   End Select

End Function
Private Function MethodVerboseRar(ByVal Method As Long, ByVal BitFlag As Long) As String
   On Error Resume Next
   Dim Dict As Integer
  'Flags
  '      0 0 0 0 0 0 0 0  &H00&   - dictionary size    64 KB
  '      0 0 1 0 0 0 0 0  &H20&   - dictionary size   128 KB
  '      0 1 0 0 0 0 0 0  &H40&   - dictionary size   256 KB
  '      0 1 1 0 0 0 0 0  &H60&   - dictionary size   512 KB
  '      1 0 0 0 0 0 0 0  &H80&   - dictionary size  1024 KB
  Dict = 2 ^ (6 + ((BitFlag \ 32) And 7)) 'Isolate bits 5 to 7
   Select Case Method
      Case 48
         MethodVerboseRar = "Stored"
      Case 51
         MethodVerboseRar = "Deflated:" & Dict & "Kb"
      Case Else
         MethodVerboseRar = "Unknown"
   End Select

End Function
Private Function MethodVerboseAce(ByVal Method, ByVal BitFlag) As String
   On Error Resume Next
   Dim Dict As Integer

   Dict = 1024 'Need to confirm this!!!
   Select Case Method
      Case 0
         MethodVerboseAce = "Stored"
      Case 1
         MethodVerboseAce = "Deflated:" & Dict & "Kb"
      Case Else
         MethodVerboseAce = "Unknown"
   End Select

End Function
Private Function BuildFullPath(Node As Node) As String
   On Error GoTo PROC_ERR
   Dim iPos As Integer
   Dim sExt As String
   Dim MyPath As String
   Dim MyDocs2 As String
   
   MyPath = Node.FullPath
   
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
   sExt = GetExt(Node.Text)
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
     If ErrMsgBox(Me.Name & ".BuildFullPath") = vbRetry Then Resume Next

End Function
Private Function CheckFAT() As Boolean

   Dim m_VolName As String * MAX_PATH
   Dim m_FileSys As String * MAX_PATH
   Dim m_VolSN As Long
   Dim m_MaxLen As Long
   Dim m_Flags As Long
    
   CheckFAT = True ' Assume True
   If GetVolumeInformation(Left(SourcePath, 3), m_VolName, MAX_PATH, m_VolSN, m_MaxLen, m_Flags, m_FileSys, MAX_PATH) Then
      If Left(m_FileSys, 3) <> "FAT" Then
         CheckFAT = False
      End If
   End If
End Function
Private Function GetAttrString(ByVal Attr As Variant) As String
   On Error GoTo ProcedureError
   Dim j As Integer
   
   Const sFill As String = "..............."
   Const sAttr As String = "rhsvdalnt?lco?e"

'00 r   0001  "Read Only"
'01 h   0002  "Hidden"
'02 s   0004  "System"
'03 v   0008  "Volume Label"
'04 f   0016  "Folder"
'05 a   0032  "Archive"
'06 l   0064  "Alias"
'07 n   0128  "Normal"
'08 t   0256  "Temporary"
'09 ?   0512   ??
'10 l   1024  "Alias"
'11 c   2048  "Compressed"
'12 o   4096  "Offline"
'13 ?   8192   ??
'14 e  16384  "Encrypted"

   GetAttrString = sFill
   If Attr Then
      For j = 0 To 14
         If Attr And (2 ^ j) Then ' Set letter
            Mid$(GetAttrString, j + 1, 1) = Mid$(sAttr, j + 1, 1)
         End If
      Next
  End If

ProcedureExit:
  Exit Function
ProcedureError:
     If ErrMsgBox(Me.Name & ".GetAttrString") = vbRetry Then Resume Next

End Function

Private Sub LoadStart()
   Start = GetTickCount
   Screen.MousePointer = vbHourglass
   LVColumnHeaders
End Sub

Private Sub LoadCleanup(SortCol As Integer)
With ListView1
   .SortKey = SortCol
   .Sorted = True
   .Visible = True
   .Refresh
   .Sorted = False
End With
InCab = False
'AssignSysIL
Screen.MousePointer = vbDefault

End Sub
Private Function FormGenDatTim(MyDate As Date) As String
   FormGenDatTim = FormatDateTime(MyDate, vbShortDate) & " " & _
                   FormatDateTime(MyDate, vbLongTime)
End Function
Private Function LVIdx(sKey As String) As Long
   LVIdx = ListView1.ColumnHeaders(sKey).SubItemIndex
End Function
Private Sub LVAddCommon(Master As Master)

Dim sMethod As String, sExt As String, FakePath As String
Dim fType As String
Dim MyIcon As Long
Dim FakeFile As Integer
Dim Ratio As Single
Dim Encrypt As Boolean
Dim Item As ListItem

   On Error GoTo ProcedureError

   With Master
      sExt = GetExt(.Filename)
      If sExt <> "" Then
         FakeFile = FreeFile
         'Create fake 0 byte file in TempDir
         'Ex: "C:\Windows\Temp\~FileName.Ext"
         FakePath = QualifyPath(TempDir) & "~" & .Filename
         Open FakePath For Binary As FakeFile
         Close FakeFile
         fType = GetFileType(sExt, FakePath, MyIcon)
         lvi.iImage = MyIcon 'index in System ImageList
         Kill FakePath
      End If
      Set Item = ListView1.ListItems.Add()
      Item.SubItems(LVIdx(nam_)) = .Filename
      lvi.iItem = Item.Index - 1 ' adjusts to 0-based
      lvi.mask = LVIF_IMAGE 'set image mask
      SendMessage ListView1.hwnd, LVM_SETITEM, 0&, lvi 'Assign
      Item.SubItems(LVIdx(ext_)) = sExt
      Item.SubItems(LVIdx(siz_)) = FormatNumber$(.Size, 0)
      Item.SubItems(LVIdx("siz2")) = Right(String(10, 48) & .Size, 10)
      Item.SubItems(LVIdx(typ_)) = fType
      Item.SubItems(LVIdx(mod_)) = FormatDateTime(.Modified, vbShortDate)
      Item.SubItems(LVIdx("mod2")) = Format(.Modified, YMDHMS)
      Item.SubItems(LVIdx(tim_)) = FormatDateTime(.Modified, vbLongTime)
      Item.SubItems(LVIdx("tim2")) = Format(.Modified, HMS)
      
      If .GridFormat <> gfCab Then ' skip cab
         Item.SubItems(LVIdx(cmp_)) = FormatNumber$(.CompSize, 0)
         Item.SubItems(LVIdx("cmp2")) = Right(String(10, 48) & .CompSize, 10)
         'Trap division by zero
         If .Size Then
            Ratio = 1 - .CompSize / .Size
            'Don't allow negative values (per PkZip/WinZip)
            'Occurs on stored/encrypted files
            If Ratio < 0 Then Ratio = 0
         Else
            Ratio = 0
         End If
         'Ratio is single. Format as desired
         Item.SubItems(LVIdx(rat_)) = Format(Ratio, "00.0%")
         Item.SubItems(LVIdx(cre_)) = FormGenDatTim(.Created)
         Item.SubItems(LVIdx("cre2")) = Format(.Created, YMDHMS)
         Item.SubItems(LVIdx(acc_)) = FormGenDatTim(.Accessed)
         Item.SubItems(LVIdx("acc2")) = Format(.Accessed, YMDHMS)
         Select Case .GridFormat
            Case gface
               sMethod = MethodVerboseAce(.Method, .flags)
            Case gfrar
               sMethod = MethodVerboseRar(.Method, .flags)
            Case gfzip
               sMethod = MethodVerboseZip(.Method, .flags)
         End Select
         Item.SubItems(LVIdx(mtd_)) = sMethod
         'Flag bit 0 is Encryption True/False
         Encrypt = (.flags And 1) * -1 'Make it Boolean
         Item.SubItems(LVIdx(enc_)) = Encrypt
         Item.SubItems(LVIdx(crc_)) = Hex$(.Crc)
         'Digital Signature extract not yet coded
         Item.SubItems(LVIdx(sig_)) = "na"
         Item.SubItems(LVIdx(pth_)) = .Path
         Item.SubItems(LVIdx(com_)) = .Comments
      End If
      Item.SubItems(LVIdx(atr_)) = GetAttrString(.Attr)
      'Save the item number for
      'other operations
      Item.Tag = .Index
      'TotalSize = TotalSize + .Size
   End With
   
ProcedureExit:
  Exit Sub
ProcedureError:
     If ErrMsgBox(Me.Name & ".LVAddCommon") = vbRetry Then Resume Next
   
End Sub
Private Sub LVColumnHeaders()
   Dim L4 As Long
   'Set L5 to 1000 for debug, 0 to hide column
   Const L5 As Long = 0
   ListView1.ListItems.Clear
   ListView1.ColumnHeaders.Clear
   ListView1.Visible = False
   ListView1.ColumnHeaders.Add , chk_, , 300
   ListView1.ColumnHeaders.Add , nam_, sName, 1800
   ListView1.ColumnHeaders.Add , ext_, sExtension, 500
   ListView1.ColumnHeaders.Add , siz_, sSize, 1100, 1
   ListView1.ColumnHeaders.Add , typ_, sType, 1400
   If InZip Then
      InCab = False
      ListView1.ColumnHeaders.Add , cmp_, "Comp", 800, 1
      ListView1.ColumnHeaders.Add , rat_, "Ratio", 700, 1
   End If
   ListView1.ColumnHeaders.Add , mod_, sModified, 1000
   ListView1.ColumnHeaders.Add , tim_, sTime, 1100
   If InCab Then
      ListView1.ColumnHeaders.Add , atr_, sAttribute, 1000
      ListView1.ColumnHeaders.Add , pth_, "Path", 4000
      GoTo JustCab
   End If
   ListView1.ColumnHeaders.Add , cre_, sCreated, 1700
   If CheckFAT Then
      L4 = 1200
   Else
      L4 = 1700
   End If
   ListView1.ColumnHeaders.Add , acc_, sAccessed, L4
   ListView1.ColumnHeaders.Add , atr_, sAttribute, 1000
   
   If InZip Then
      ListView1.ColumnHeaders.Add , mtd_, "Method", 1000
      ListView1.ColumnHeaders.Add , enc_, "Encoded", 1000
      ListView1.ColumnHeaders.Add , crc_, "Hex CRC", 1000
      ListView1.ColumnHeaders.Add , sig_, "Signature", 1000
      ListView1.ColumnHeaders.Add , pth_, "Path", 4000
      ListView1.ColumnHeaders.Add , com_, "Comments", 4000
   Else
      ListView1.ColumnHeaders.Add , dos_, sMsDos, 1600
   End If
   'Invisible columns for sort
   ListView1.ColumnHeaders.Add , "cmp2", , L5
   ListView1.ColumnHeaders.Add , "cre2", , L5
   ListView1.ColumnHeaders.Add , "acc2", , L5
JustCab:
   ListView1.ColumnHeaders.Add , "siz2", , L5
   ListView1.ColumnHeaders.Add , "mod2", , L5
   ListView1.ColumnHeaders.Add , "tim2", , L5
End Sub
Private Sub Form_Resize()
   On Error Resume Next
   
    Splitter1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight

End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

   Static Ascend As Boolean ' Remember Ascend
   Ascend = Not Ascend      ' Toggle Ascend
   With ListView1
      .SortOrder = Abs(Ascend) ' Boolean to 0/1
      Select Case ColumnHeader.Key
         Case mod_, tim_, siz_, cre_, acc_, cmp_
            ' Map to hidden column
            .SortKey = .ColumnHeaders(ColumnHeader.Key & "2").SubItemIndex
         Case Else ' Sort alphanumeric
            .SortKey = ColumnHeader.Index - 1
      End Select
   End With
   ' Sort
   ListView1.Sorted = True
   
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim MySize, lRow As Long
   lRow = ListView1.SelectedItem.Index
   If ListView1.ListItems(lRow).SubItems(3) = sFolder Then
      If ListView1.ListItems(lRow).SubItems(2) = "" Then
         Screen.MousePointer = vbHourglass
         MySize = DirSpace(SourcePath & ListView1.HitTest(x, y)) * 10000
         ListView1.ListItems(lRow).SubItems(2) = FormatSize(MySize)
         ListView1.ListItems(lRow).SubItems(10) = Right(String(10, 48) & MySize, 10)
         Screen.MousePointer = vbDefault
      End If
   End If
End Sub
Public Function AssignSysIL() As Boolean

    On Error GoTo ER
    '// Get handles to system imagelists and assign them to the listview.
    '// Call this method when changing the ListView.View property.
    '// Apparently, the VB ListView checks for a VB ImageList when
    '// changing View styles. Finding none, it sets its ImageList handle
    '// to zero, effectively giving our api-assigned system imagelist the boot.
    Dim hSysIL(0 To 1) As Long 'sys IL handles
    Dim SFI As SHFILEINFO
    Dim L As Long
  '  Dim listview1 As ListView
  '  Set listview1 = ListView
    
    '// Normally an imagelist is destroyed when it's no longer needed.
    '// However since the system image lists are global shared objects, it's
    '// important not to do so. If you destroy a system image list, you'll
    '// trash the icon display for anything that uses it, including the Desktop,
    '// StartMenu, Explorer, etc. Restoring things requires re-starting the
    '// Windows shell.
    '// The next statement sets a ListView-specific flag so that it does not
    '// destroy its imagelist(s) when the ListView itself is destroyed. Although
    '// the VB ListView sets this flag by default, the API ListView does not.
    '// Rather than gamble that future versions of the VB ListView adopt the
    '// same behavior, take the safe route and make sure it's set.
    SetWindowLong ListView1.hwnd, _
                    GWL_STYLE, _
                GetWindowLong(ListView1.hwnd, GWL_STYLE) Or _
            LVS_SHAREIMAGELISTS
    
    
    For L = 0 To 1
        '// get sysIL handle.
        '// AFIK, the only time a sysIL handle is likely to change is
        '// when the user changes display settings, so it may not be
        '// necessary to grab them each time. (But it can't hurt).
        hSysIL(L) = SHGetFileInfo(App.Path, 0&, SFI, Len(SFI), SHGFI_SYSICONINDEX Or L)
        '// assign / re-assign sys IL to listview...
        SendMessage ListView1.hwnd, LVM_SETIMAGELIST, L, ByVal hSysIL(L)
    Next
    
    AssignSysIL = True
    
    Exit Function '0
ER:
    AssignSysIL = False
    Debug.Print Err.Description & "   AssignSysIL"
    
End Function


Private Sub rar_FileFound(ByVal Count As Long, ByVal Filename As String, ByVal DateTime As Date, ByVal Size As Variant, ByVal CompSize As Variant, ByVal Method As Long, ByVal Attr As Variant, ByVal Path As String, ByVal flags As Long, ByVal Crc As Long, ByVal Comments As String)
   With Master
     .GridFormat = gfrar
     .Index = Count
     .Filename = Filename
     .Size = Size
     .Modified = DateTime
     .Created = DateTime
     .Accessed = DateTime
     .Attr = Attr
     .Path = Path
     .CompSize = CompSize
     .Method = Method
     .flags = flags
     .Encypted = (flags And 4) * -1 'Make it Boolean
     .Crc = Crc
     '.Sig = 0
     .Comments = Comments
   End With
   LVAddCommon Master
End Sub

Private Sub zip_FileFound(ByVal Count As Long, ByVal Filename As String, ByVal DateTime As Date, ByVal Size As Variant, ByVal CompSize As Variant, ByVal Method As Long, ByVal Attr As Variant, ByVal Path As String, ByVal flags As Long, ByVal Crc As Long, ByVal Comments As String)
   With Master
     .GridFormat = gfzip
     .Index = Count
     .Filename = Filename
     .Size = Size
     .Modified = DateTime
     .Created = DateTime
     .Accessed = DateTime
     .Attr = Attr
     .Path = Path
     .CompSize = CompSize
     .Method = Method
     .flags = flags
     .Encypted = (flags And 1) * -1 'Make it Boolean
     .Crc = Crc
     '.Sig = 0
     .Comments = Comments
   End With
   LVAddCommon Master

End Sub
