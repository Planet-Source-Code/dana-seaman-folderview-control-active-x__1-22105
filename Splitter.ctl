VERSION 5.00
Begin VB.UserControl Splitter 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3165
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   195
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   211
   ToolboxBitmap   =   "Splitter.ctx":0000
   Begin VB.PictureBox Spiral 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   0
      Picture         =   "Splitter.ctx":00D0
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Spiral2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   60
      Picture         =   "Splitter.ctx":030A
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   3
      Top             =   180
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox Spiral2R 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   60
      Picture         =   "Splitter.ctx":0544
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox SpiralR 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   0
      Picture         =   "Splitter.ctx":077E
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   1
      Top             =   540
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox picSplitter 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   2895
      Left            =   1560
      ScaleHeight     =   193
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   0
      Top             =   0
      Width           =   270
   End
End
Attribute VB_Name = "Splitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------
' Splitter control by Tim Humphrey
' 3/14/2001
' zzhumphreyt@techie.com
'
' modified dseaman@ieg.com.br
' spiral binder image tiled into splitter
' Added TileArea so no more external dependencies except API.
' ----------
'
' Usage: Add controls to the Splitter control
' the same way you do for a frame;
' Set Child1 and/or Child2 to the names of the
' controls you want to be resized.
'
' - All sizes are in pixels (was twips).
'
' - A default representation of the splitter bar is displayed in design mode
'
' - The list box can only be resized in certain increments, 195 twips,
'   if you add one resize the Splitter control to match it.
'
' - While resizing, the escape key may be pressed to cancel and undo
'   the dragging
'--------------------------------------------------

Option Explicit
Option Compare Text

'-------------------- Notes --------------------
' There is a dependency between the MinSizeAux property and the Orientation property,
' initialize Orientation first

'-------------------- Enumerations --------------------
Public Enum AppearanceConstants
   vbFlat = 0
   vb3D = 1
End Enum

Public Enum BorderConstants
   vbBSNone = 0
   vbFixedSingle = 1
End Enum

Public Enum OrientationConstants
   OC_HORIZONTAL = 0
   OC_VERTICAL = 1
End Enum

'-------------------- Constants --------------------
'----- Property strings
Const cStr_SplitterAppearance As String = "SplitterAppearance"
Const cStr_SplitterBorder As String = "SplitterBorder"
Const cStr_SplitterColor As String = "SplitterColor"
Const cStr_Orientation As String = "Orientation"
Const cStr_SplitterSize As String = "SplitterSize"
Const cStr_RatioFromTop As String = "RatioFromTop"
Const cStr_Child1 As String = "Child1"
Const cStr_Child2 As String = "Child2"
Const cStr_MinSize1 As String = "MinSize1"
Const cStr_MinSize2 As String = "MinSize2"
Const cStr_MinSizeAux As String = "MinSizeAux"
Const cStr_LiveUpdate As String = "LiveUpdate"

'----- Defaults
Const kDefSplitterAppearance As Integer = vb3D
Const kDefSplitterBorder As Integer = vbFixedSingle
Const kDefSplitterColor As Long = &H404040
Const kDefOrientation As Integer = OC_HORIZONTAL
Const kDefSplitterSize As Integer = 18 ' 270
Const kDefRatioFromTop As Single = 0.5
Const kDefChild1 As String = ""
Const kDefChild2 As String = ""
Const kDefMinSize1 As Long = 17 ' 255
Const kDefMinSize2 As Long = 17 ' 255
Const kDefMinSizeAux As Long = 17 '255
Const kDefLiveUpdate As Boolean = True

'-------------------- Variables --------------------
'----- Public properties
Private m_SplitterAppearance  As AppearanceConstants
Private m_SplitterBorder      As BorderConstants
Private m_SplitterColor       As Long
Private m_Orientation         As OrientationConstants
Private m_SplitterSize        As Integer
Private m_RatioFromTop        As Single
Private m_Child1              As String
Private m_Child2              As String
Private m_MinSize1            As Long
Private m_MinSize2            As Long
Private m_MinSizeAux          As Long
Private m_LiveUpdate          As Boolean
Private m_MinRequiredSpace    As Long
Private m_AvailableAuxSpace   As Long

'----- Control use
Private gUpdateDisplay  As Boolean
Private gMoving         As Boolean
Private gOrigRatio      As Single
Private gOrigPos        As Single

'-------------------- Events --------------------
Public Event Resize()

'-------------------- API Types --------------------
Private Type Point
   x                    As Long
   y                    As Long
End Type

'-------------------- API Functions --------------------
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As Point) As Boolean
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As Point) As Boolean
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Public Property Get SplitterColor() As Long
   SplitterColor = m_SplitterColor
End Property

Public Property Let SplitterColor(Value As Long)
   m_SplitterColor = Value
   PropertyChanged cStr_SplitterColor
End Property

Public Property Get Orientation() As OrientationConstants
   Orientation = m_Orientation
End Property

Public Property Let Orientation(Value As OrientationConstants)
   m_Orientation = Value
   PropertyChanged cStr_Orientation
   UserControl_Resize
End Property

Public Property Get SplitterSize() As Integer
   SplitterSize = m_SplitterSize
End Property

Public Property Let SplitterSize(Value As Integer)
   If Value > 0 Then
      m_SplitterSize = Value
      PropertyChanged cStr_SplitterSize
      MinRequiredSpace = CalcMinRequiredSpace
      UserControl_Show
   End If
End Property

Public Property Get RatioFromTop() As Single
   RatioFromTop = m_RatioFromTop
End Property

Public Property Let RatioFromTop(Value As Single)
   If (Value >= 0) And (Value <= 1) Then
      m_RatioFromTop = Value
      PropertyChanged cStr_RatioFromTop
      UserControl_Show
   End If
End Property

Public Property Get Child1() As String
   Child1 = m_Child1
End Property

Public Property Let Child1(Value As String)
   m_Child1 = Value
   PropertyChanged cStr_Child1
   UserControl_Show
End Property

Private Function ObjectFromName(Name As String) As Object

   Dim i   As Integer

   If LenB(Name) Then
      For i = 0 To UserControl.ContainedControls.Count - 1
         'Step through controls and look for a match.
         If UserControl.ContainedControls(i).Name = Name Then
            Set ObjectFromName = UserControl.ContainedControls(i)
            Exit For
         End If
      Next
   End If
   
End Function

Public Property Get Child2() As String
   Child2 = m_Child2
End Property

Public Property Let Child2(Value As String)
   m_Child2 = Value
   PropertyChanged cStr_Child2
   UserControl_Show
End Property

Public Property Get MinSize1() As Long
   MinSize1 = m_MinSize1
End Property

Public Property Let MinSize1(Value As Long)
   If Value >= 0 Then
      m_MinSize1 = Value
      PropertyChanged cStr_MinSize1
      MinRequiredSpace = CalcMinRequiredSpace
      UserControl_Show
   End If
End Property

Public Property Get MinSize2() As Long
   MinSize2 = m_MinSize2
End Property

Public Property Let MinSize2(Value As Long)
   If Value >= 0 Then
      m_MinSize2 = Value
      PropertyChanged cStr_MinSize2
      MinRequiredSpace = CalcMinRequiredSpace
      UserControl_Show
   End If
End Property

Public Property Get MinSizeAux() As Long
   MinSizeAux = m_MinSizeAux
   PropertyChanged cStr_MinSizeAux
   UserControl_Show
End Property

Public Property Let MinSizeAux(Value As Long)
   If Value >= 0 Then
      m_MinSizeAux = Value
      AvailableAuxSpace = CalcAvailableAuxSpace
   End If
End Property

Public Property Get LiveUpdate() As Boolean
   LiveUpdate = m_LiveUpdate
End Property

Public Property Let LiveUpdate(Value As Boolean)
   m_LiveUpdate = Value
   PropertyChanged cStr_LiveUpdate
End Property

Private Property Get MinRequiredSpace() As Long
   MinRequiredSpace = m_MinRequiredSpace
End Property

Private Property Let MinRequiredSpace(Value As Long)
   m_MinRequiredSpace = Value
End Property

Private Function CalcMinRequiredSpace() As Long
   CalcMinRequiredSpace = MinSize1 + SplitterSize + MinSize2
End Function

Private Property Get AvailableAuxSpace() As Long
   AvailableAuxSpace = m_AvailableAuxSpace
End Property

Private Property Let AvailableAuxSpace(Value As Long)
   m_AvailableAuxSpace = Value
End Property

Private Function CalcAvailableAuxSpace() As Long

   Select Case Orientation
      Case OC_HORIZONTAL
         If UserControl.ScaleHeight > MinSizeAux Then
            CalcAvailableAuxSpace = UserControl.ScaleHeight
         Else
            CalcAvailableAuxSpace = MinSizeAux
         End If
      Case OC_VERTICAL
         If UserControl.ScaleWidth > MinSizeAux Then
            CalcAvailableAuxSpace = UserControl.ScaleWidth
         Else
            CalcAvailableAuxSpace = MinSizeAux
         End If
   End Select

End Function

Private Sub picSplitter_KeyPress(KeyAscii As Integer)
   If gMoving And KeyAscii = 27 Then 'if user pressed escape
      RatioFromTop = gOrigRatio

      If LiveUpdate Then
         MoveSplitter
         MoveChildren
      Else
         '   picSplitter.BackColor = vbButtonFace
         '   picSplitter.BorderStyle = vbBSNone
         MoveSplitter
      End If

      gMoving = False
   End If
End Sub

Private Sub picSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   '-------------------- Variables --------------------
   Dim originalPoint    As Point
   Dim modifiedPoint    As Point
   Dim offsetX          As Long
   Dim offsetY          As Long

   '-------------------- Code --------------------
   If Button = vbLeftButton Then
      gOrigRatio = RatioFromTop
      Select Case Orientation
         Case OC_HORIZONTAL
            gOrigPos = x
         Case OC_VERTICAL
            gOrigPos = y
      End Select

      picSplitter.ZOrder 0
      picSplitter.Appearance = SplitterAppearance
      '  picSplitter.BackColor = SplitterColor
      '  picSplitter.BorderStyle = SplitterBorder

      If Not LiveUpdate Then
         'Changing the picture box to include a border will alter the scalewidth/scaleheight,
         'thereby immediately triggering a mouse moved event; we must compensate for this
         GetCursorPos originalPoint
         ScreenToClient picSplitter.hwnd, originalPoint

         '  picSplitter.Appearance = SplitterAppearance
         '  picSplitter.BackColor = SplitterColor
         '  picSplitter.BorderStyle = SplitterBorder

         GetCursorPos modifiedPoint
         ScreenToClient picSplitter.hwnd, modifiedPoint

         Select Case Orientation
            Case OC_HORIZONTAL
               offsetX = (originalPoint.x - modifiedPoint.x) ' * Screen.TwipsPerPixelX
               gOrigPos = gOrigPos - offsetX
            Case OC_VERTICAL
               offsetY = (originalPoint.y - modifiedPoint.y) ' * Screen.TwipsPerPixelY
               gOrigPos = gOrigPos - offsetY
         End Select
      End If
      gMoving = True
   End If

   TileSplitter True

End Sub

Private Sub picSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   '-------------------- Variables --------------------
   Dim blnResize        As Boolean
   Dim sngCurrPos       As Single
   Dim sngNewPos        As Single
   Dim sngMinPos        As Single
   Dim sngMaxPos        As Single
   Dim availableSpace   As Single

   '-------------------- Code --------------------
   If gMoving And Button = vbLeftButton Then
      Select Case Orientation
         Case OC_HORIZONTAL
            sngCurrPos = picSplitter.Left
            sngNewPos = picSplitter.Left + (x - gOrigPos) 'only add offset from original position
            sngMinPos = MinSize1
            sngMaxPos = UserControl.ScaleWidth - MinSize2 - SplitterSize
         Case OC_VERTICAL
            sngCurrPos = picSplitter.Top
            sngNewPos = picSplitter.Top + (y - gOrigPos) 'only add offset from original position
            sngMinPos = MinSize1
            sngMaxPos = UserControl.ScaleHeight - MinSize2 - SplitterSize
      End Select

      blnResize = False
      If (sngNewPos > sngMinPos) And (sngNewPos < sngMaxPos) Then
         blnResize = True
      Else
         If sngNewPos <= sngMinPos Then 'too low
            If sngCurrPos <> sngMinPos Then
               blnResize = True
               sngNewPos = sngMinPos
            End If
         Else 'too high
            If sngCurrPos <> sngMaxPos Then
               blnResize = True
               sngNewPos = sngMaxPos
            End If
         End If
      End If

      If blnResize Then
         gUpdateDisplay = False

         'Move splitter
         Select Case Orientation
            Case OC_HORIZONTAL
               picSplitter.Left = sngNewPos
               availableSpace = UserControl.ScaleWidth
            Case OC_VERTICAL
               picSplitter.Top = sngNewPos
               availableSpace = UserControl.ScaleHeight
         End Select

         'Determine new ratio
         If availableSpace <> 0 Then
            RatioFromTop = (sngNewPos + (SplitterSize \ 2)) / availableSpace
         Else
            RatioFromTop = 0
         End If

         gUpdateDisplay = True

         If LiveUpdate Then
            MoveChildren
         End If
      End If
   End If
End Sub

Private Sub picSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbLeftButton Then
      ' If Not LiveUpdate Then
      '  picSplitter.BackColor = vbButtonFace
      '  picSplitter.BorderStyle = vbBSNone
      MoveChildren
      picSplitter.Cls
      picSplitter.AutoRedraw = True
      TileSplitter
      ' End If
      gMoving = False
   End If
End Sub

Private Sub TileSplitter(Optional Down As Boolean = False)
   Dim hDCSource        As Long
   Dim W                As Long
   Dim H                As Long

   Select Case Orientation
      Case OC_HORIZONTAL
         W = Spiral.ScaleWidth
         H = Spiral.ScaleHeight
         If Down Then
            hDCSource = SpiralR.hDC
         Else
            hDCSource = Spiral.hDC
         End If
      Case OC_VERTICAL
         W = Spiral.ScaleHeight
         H = Spiral.ScaleWidth
         If Down Then
            hDCSource = Spiral2R.hDC
         Else
            hDCSource = Spiral2.hDC
         End If
   End Select
   TileArea picSplitter.hDC, _
      0, _
      0, _
      picSplitter.ScaleWidth, _
      picSplitter.ScaleHeight, _
      hDCSource, _
      W, _
      H

End Sub

Public Sub TileArea( _
   ByVal hDCDestin As Long, _
   ByVal x As Long, _
   ByVal y As Long, _
   ByVal Width As Long, _
   ByVal Height As Long, _
   ByVal hDCSource As Long, _
   ByVal SrcWidth As Long, _
   ByVal SrcHeight As Long _
   )
   Dim lSrcX            As Long
   Dim lSrcY            As Long
   Dim lSrcStartX       As Long
   Dim lSrcStartY       As Long
   Dim lSrcStartWidth   As Long
   Dim lSrcStartHeight  As Long
   Dim lDstX            As Long
   Dim lDstY            As Long
   Dim lDstWidth        As Long
   Dim lDstHeight       As Long

   lSrcStartX = (x Mod SrcWidth)
   lSrcStartY = (y Mod SrcHeight)
   lSrcStartWidth = (SrcWidth - lSrcStartX)
   lSrcStartHeight = (SrcHeight - lSrcStartY)
   lSrcX = lSrcStartX
   lSrcY = lSrcStartY

   lDstY = y
   lDstHeight = lSrcStartHeight

   Do While lDstY < (y + Height)
      If (lDstY + lDstHeight) > (y + Height) Then
         lDstHeight = y + Height - lDstY
      End If
      lDstWidth = lSrcStartWidth
      lDstX = x
      lSrcX = lSrcStartX
      Do While lDstX < (x + Width)
         If (lDstX + lDstWidth) > (x + Width) Then
            lDstWidth = x + Width - lDstX
            If (lDstWidth = 0) Then
               lDstWidth = 4
            End If
         End If
         'If (lDstWidth > Width) Then lDstWidth = Width
         'If (lDstHeight > Height) Then lDstHeight = Height
         BitBlt hDCDestin, lDstX, lDstY, lDstWidth, lDstHeight, hDCSource, lSrcX, lSrcY, vbSrcCopy
         lDstX = lDstX + lDstWidth
         lSrcX = 0
         lDstWidth = SrcWidth
      Loop
      lDstY = lDstY + lDstHeight
      lSrcY = 0
      lDstHeight = SrcHeight
   Loop
End Sub

Private Sub UserControl_InitProperties()
   gUpdateDisplay = False

   'Set controls
   If Not Ambient.UserMode Then
      picSplitter.BorderStyle = vbBSSolid
   Else
      picSplitter.BorderStyle = vbBSNone
   End If

   'Public properties
   SplitterAppearance = kDefSplitterAppearance
   SplitterBorder = kDefSplitterBorder
   SplitterColor = kDefSplitterColor
   Orientation = kDefOrientation
   ' SplitterSize = kDefSplitterSize
   SplitterSize = 18 '* Screen.TwipsPerPixelX
   RatioFromTop = kDefRatioFromTop
   Child1 = kDefChild1
   Child2 = kDefChild2
   MinSize1 = kDefMinSize1
   MinSize2 = kDefMinSize2
   MinSizeAux = kDefMinSizeAux
   LiveUpdate = kDefLiveUpdate

   gUpdateDisplay = True
   gMoving = False
   UserControl_Show
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   gUpdateDisplay = False

   If Not Ambient.UserMode Then
      picSplitter.BorderStyle = vbBSSolid
   Else
      picSplitter.BorderStyle = vbBSNone
   End If

   With PropBag
      SplitterAppearance = .ReadProperty(cStr_SplitterAppearance, kDefSplitterAppearance)
      SplitterBorder = .ReadProperty(cStr_SplitterBorder, kDefSplitterBorder)
      SplitterColor = .ReadProperty(cStr_SplitterColor, kDefSplitterColor)
      Orientation = .ReadProperty(cStr_Orientation, kDefOrientation)
      SplitterSize = .ReadProperty(cStr_SplitterSize, kDefSplitterSize)
      RatioFromTop = .ReadProperty(cStr_RatioFromTop, kDefRatioFromTop)
      Child1 = .ReadProperty(cStr_Child1, kDefChild1)
      Child2 = .ReadProperty(cStr_Child2, kDefChild2)
      MinSize1 = .ReadProperty(cStr_MinSize1, kDefMinSize1)
      MinSize2 = .ReadProperty(cStr_MinSize2, kDefMinSize2)
      MinSizeAux = .ReadProperty(cStr_MinSizeAux, kDefMinSizeAux)
      LiveUpdate = .ReadProperty(cStr_LiveUpdate, kDefLiveUpdate)
      UserControl.Appearance = .ReadProperty("Appearance", 1)
      UserControl.BorderStyle = .ReadProperty("BorderStyle", 1)
   End With

   gUpdateDisplay = True
   UserControl_Show

End Sub

Private Sub UserControl_Resize()
   AvailableAuxSpace = CalcAvailableAuxSpace
   picSplitter.Height = UserControl.ScaleHeight
   TileSplitter
   UserControl_Show
End Sub

Private Sub UserControl_Show()
   MoveSplitter
   MoveChildren
End Sub

Private Sub MoveSplitter()
   If gUpdateDisplay Then
      Select Case Orientation
         Case OC_HORIZONTAL
            picSplitter.Move SplitterTop(UserControl.ScaleWidth), 0, SplitterSize, AvailableAuxSpace
            picSplitter.MousePointer = vbSizeWE
         Case OC_VERTICAL
            picSplitter.Move 0, SplitterTop(UserControl.ScaleHeight), AvailableAuxSpace, SplitterSize
            picSplitter.MousePointer = vbSizeNS
      End Select
   End If
End Sub
Private Sub MoveWindow2(obj As Object, _
   ByVal L As Integer, _
   ByVal T As Integer, _
   ByVal W As Integer, _
   ByVal H As Integer)
   
   Dim SX As Single
   Dim SY As Single
   
   SX = Screen.TwipsPerPixelX
   SY = Screen.TwipsPerPixelY
   obj.Move L * SX, T * SY, W * SX, H * SY
   
End Sub
Private Sub MoveChildren()
   '-------------------- Variables --------------------
   Dim vObjChild1       As Object
   Dim vObjChild2       As Object
   Dim newLeft1         As Integer
   Dim newTop1          As Integer
   Dim newWidth1        As Integer
   Dim newHeight1       As Integer
   Dim newLeft2         As Integer
   Dim newTop2          As Integer
   Dim newWidth2        As Integer
   Dim newHeight2       As Integer

   '-------------------- Code --------------------
   If gUpdateDisplay Then
      UserControl.AutoRedraw = False

      Set vObjChild1 = ObjectFromName(m_Child1)
      Set vObjChild2 = ObjectFromName(m_Child2)

      'Hack around evil ListView control
      If Not (vObjChild1 Is Nothing) And (TypeName(vObjChild1) = "ListView") Then
         newLeft1 = -1
         newTop1 = -1
         newWidth1 = 2
         newHeight1 = 2
      End If

      If Not (vObjChild2 Is Nothing) And (TypeName(vObjChild2) = "ListView") Then
         newLeft2 = -1
         newTop2 = -1
         newWidth2 = 2
         newHeight2 = 2
      End If

      Select Case Orientation
         Case OC_HORIZONTAL
            If Not (vObjChild1 Is Nothing) Then
               MoveWindow2 vObjChild1, _
                  newLeft1, _
                  newTop1, _
                  newWidth1 + picSplitter.Left, _
                  newHeight1 + AvailableAuxSpace
            End If

            If Not (vObjChild2 Is Nothing) Then
               newLeft2 = newLeft2 + (picSplitter.Left + m_SplitterSize)
               newTop2 = newTop2 + 0
               newHeight2 = newHeight2 + AvailableAuxSpace

               If UserControl.ScaleWidth - (picSplitter.Left + m_SplitterSize) >= MinSize2 Then
                  newWidth2 = newWidth2 + (UserControl.ScaleWidth - (picSplitter.Left + m_SplitterSize))
               Else
                  newWidth2 = newWidth2 + MinSize2
               End If
               MoveWindow2 vObjChild2, _
                  newLeft2, _
                  newTop2, _
                  newWidth2, _
                  newHeight2
            End If
         Case OC_VERTICAL
            If Not (vObjChild1 Is Nothing) Then
               MoveWindow2 vObjChild1, _
                  newLeft1, _
                  newTop1, _
                  newWidth1 + AvailableAuxSpace, _
                  newHeight1 + picSplitter.Top
            End If

            If Not (vObjChild2 Is Nothing) Then
               newLeft2 = newLeft2 + 0
               newTop2 = newTop2 + picSplitter.Top + m_SplitterSize
               newWidth2 = newWidth2 + AvailableAuxSpace

               If UserControl.ScaleHeight - (picSplitter.Top + m_SplitterSize) >= MinSize2 Then
                  newHeight2 = newHeight2 + UserControl.ScaleHeight - (picSplitter.Top + m_SplitterSize)
               Else
                  newHeight2 = newHeight2 + MinSize2
               End If
               MoveWindow2 vObjChild2, _
                  newLeft2, _
                  newTop2, _
                  newWidth2, _
                  newHeight2
            End If
      End Select

      RaiseEvent Resize
      UserControl.AutoRedraw = True

   End If

End Sub

Public Property Get SplitterAppearance() As AppearanceConstants
   SplitterAppearance = m_SplitterAppearance
End Property

Public Property Let SplitterAppearance(Value As AppearanceConstants)
   m_SplitterAppearance = Value
   PropertyChanged cStr_SplitterAppearance
End Property

Public Property Get SplitterBorder() As BorderConstants
   SplitterBorder = m_SplitterBorder
End Property

Public Property Let SplitterBorder(Value As BorderConstants)
   m_SplitterBorder = Value
   PropertyChanged cStr_SplitterBorder
End Property

Private Function SplitterTop(availableSpace As Integer) As Integer
   '-------------------- Variables --------------------
   Dim newPos           As Integer
   Dim size1Violated    As Boolean

   '-------------------- Code --------------------
   If availableSpace > MinRequiredSpace Then
      newPos = (availableSpace * RatioFromTop) - (SplitterSize \ 2)

      'Correct bounds if needed
      If newPos < 0 Then
         newPos = 0
      End If
      If (newPos + SplitterSize) > availableSpace Then
         newPos = availableSpace - SplitterSize
      End If

      'See if Child1 bounds violated
      If newPos < MinSize1 Then
         newPos = MinSize1
         size1Violated = True
      Else
         size1Violated = False
      End If

      'See if Child2 bounds violated
      If ((newPos + SplitterSize) > (availableSpace - MinSize2)) And Not size1Violated Then
         newPos = availableSpace - MinSize2 - SplitterSize
      End If
   Else
      newPos = MinSize1
   End If

   SplitterTop = newPos
End Function

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   With PropBag
      .WriteProperty cStr_SplitterAppearance, SplitterAppearance, kDefSplitterAppearance
      .WriteProperty cStr_SplitterBorder, SplitterBorder, kDefSplitterBorder
      .WriteProperty cStr_SplitterColor, SplitterColor, kDefSplitterColor
      .WriteProperty cStr_Orientation, m_Orientation, kDefOrientation
      .WriteProperty cStr_SplitterSize, m_SplitterSize, kDefSplitterSize
      .WriteProperty cStr_RatioFromTop, m_RatioFromTop, kDefRatioFromTop
      .WriteProperty cStr_Child1, m_Child1, kDefChild1
      .WriteProperty cStr_Child2, m_Child2, kDefChild2
      .WriteProperty cStr_MinSize1, m_MinSize1, kDefMinSize1
      .WriteProperty cStr_MinSize2, m_MinSize2, kDefMinSize2
      .WriteProperty cStr_MinSizeAux, m_MinSizeAux, kDefMinSizeAux
      .WriteProperty cStr_LiveUpdate, m_LiveUpdate, kDefLiveUpdate
      .WriteProperty "Appearance", UserControl.Appearance, 1
      .WriteProperty "BorderStyle", UserControl.BorderStyle, 1
   End With

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As Integer
   Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
   UserControl.Appearance() = New_Appearance
   PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
   BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
   UserControl.BorderStyle() = New_BorderStyle
   PropertyChanged "BorderStyle"
End Property

