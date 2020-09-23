VERSION 5.00
Begin VB.UserControl tjProgress 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox picSegment 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   99
      Left            =   3720
      Picture         =   "tjPBAR.ctx":0000
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   7
      TabIndex        =   11
      Top             =   1800
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   2400
      Picture         =   "tjPBAR.ctx":0222
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   7
      TabIndex        =   10
      Top             =   3000
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   240
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   261
      TabIndex        =   9
      Top             =   2280
      Visible         =   0   'False
      Width           =   3915
   End
   Begin VB.PictureBox picSegment 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   480
      Picture         =   "tjPBAR.ctx":0444
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox picSegment 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   900
      Picture         =   "tjPBAR.ctx":058A
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox picSegment 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   1320
      Picture         =   "tjPBAR.ctx":06D0
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox picSegment 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   1740
      Picture         =   "tjPBAR.ctx":0816
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox picSegment 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   2160
      Picture         =   "tjPBAR.ctx":095C
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox picSegment 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   2580
      Picture         =   "tjPBAR.ctx":0AA2
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox picSegment 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   3000
      Picture         =   "tjPBAR.ctx":0BE8
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox picSegment 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   7
      Left            =   3360
      Picture         =   "tjPBAR.ctx":0D2E
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   7
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox picProgress 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   8
      Top             =   0
      Width           =   3855
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   1680
      Stretch         =   -1  'True
      Top             =   2640
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "tjProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetPixel Lib "GDI32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "GDI32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public Enum ProgressTypes
    Standard = 0
    Cool = 1
    Cooler = 2
    CoolBlue = 3
    Yellow = 4
    Fire = 5
    Bark = 6
    BarkLarge = 7
    Custom = 99
End Enum

Public Enum CaptionTypes
    Percentage = 0
    Text = 1
End Enum

Public Enum CaptionX
    pbarLeft = 0
    pbarCenter = 1
    pbarRight = 2
End Enum

Public Enum CaptionY
    pbarTop = 0
    pbarMiddle = 1
    pbarBottom = 2
End Enum

Dim uCustomPicture As StdPicture
Dim uFont As StdFont
Dim uRounded As Boolean
Dim uCaptionColor As Long
Dim uCaptionPositionX As Integer
Dim uCaptionPositionY As Integer
Dim uShowCaption As Boolean
Dim uCaptionType As Integer
Dim uCaption As String
Dim uValue As Long
Dim uType As Integer
Dim uMin As Long
Dim uMax As Long

Public Property Set Font(ByRef new_font As StdFont)
    SetFont new_font
    PropertyChanged "Font"
    SetProgressBarValue picProgress, uMin, uMax, uValue, Picture2, uShowCaption
End Property

Public Property Let Font(ByRef new_font As StdFont)
    SetFont new_font
    PropertyChanged "Font"
    SetProgressBarValue picProgress, uMin, uMax, uValue, Picture2, uShowCaption
End Property
Public Property Get Font() As StdFont
    Set Font = uFont
End Property

Public Property Get CustomPicture() As StdPicture
    Set CustomPicture = uCustomPicture
End Property

Public Property Set CustomPicture(ByVal new_pict As StdPicture)
    Set uCustomPicture = new_pict
    picSegment(99).Cls
    picSegment(99).Width = new_pict.Width
    picSegment(99).Height = new_pict.Height
    Set picSegment(99).Picture = new_pict
    Call SetImage
    SetProgressBarValue picProgress, uMin, uMax, uValue, Picture2, uShowCaption
    PropertyChanged "CustomPicture"
End Property

Private Sub SetFont(ByRef new_font As StdFont)
    With uFont
        .Bold = new_font.Bold
        .Italic = new_font.Italic
        .Name = new_font.Name
        .Size = new_font.Size
    End With
End Sub

Public Property Get Rounded() As Boolean
    Rounded = uRounded
End Property

Public Property Let Rounded(New_Value As Boolean)
    uRounded = New_Value
    PropertyChanged "Rounded"
    SetProgressBarValue picProgress, uMin, uMax, uValue, Picture2, uShowCaption
End Property

Public Property Get CaptionColor() As OLE_COLOR
    CaptionColor = uCaptionColor
End Property

Public Property Let CaptionColor(ByRef New_Value As OLE_COLOR)
    uCaptionColor = New_Value
    PropertyChanged "CaptionColor"
'    SetProgressBarValue picProgress, uMin, uMax, uValue, picSegment(uType), uShowCaption
    SetProgressBarValue picProgress, uMin, uMax, uValue, Picture2, uShowCaption
End Property

Public Property Get CaptionYPosition() As CaptionY
    CaptionYPosition = uCaptionPositionY
End Property

Public Property Let CaptionYPosition(New_Value As CaptionY)
    uCaptionPositionY = New_Value
    PropertyChanged "CaptionYPosition"
'    SetProgressBarValue picProgress, uMin, uMax, uValue, picSegment(uType), uShowCaption
    SetProgressBarValue picProgress, uMin, uMax, uValue, Picture2, uShowCaption
End Property

Public Property Get CaptionXPosition() As CaptionX
    CaptionXPosition = uCaptionPositionX
End Property

Public Property Let CaptionXPosition(New_Value As CaptionX)
    uCaptionPositionX = New_Value
    PropertyChanged "CaptionXPosition"
'    SetProgressBarValue picProgress, uMin, uMax, uValue, picSegment(uType), uShowCaption
    SetProgressBarValue picProgress, uMin, uMax, uValue, Picture2, uShowCaption
End Property

Public Property Get Caption() As String
    Caption = uCaption
End Property

Public Property Let Caption(ByVal New_Value As String)
    uCaption = New_Value
    PropertyChanged "Caption"
'    SetProgressBarValue picProgress, uMin, uMax, uValue, picSegment(uType), uShowCaption
    SetProgressBarValue picProgress, uMin, uMax, uValue, Picture2, uShowCaption

End Property

Public Property Get ShowCaption() As Boolean
    ShowCaption = uShowCaption
End Property

Public Property Let ShowCaption(ByVal New_Value As Boolean)
    uShowCaption = New_Value
    PropertyChanged "ShowCaption"
    SetProgressBarValue picProgress, uMin, uMax, uValue, Picture2, uShowCaption
End Property

Public Property Get CaptionType() As CaptionTypes
    CaptionType = uCaptionType
End Property

Public Property Let CaptionType(ByVal New_Value As CaptionTypes)
    uCaptionType = New_Value
    PropertyChanged "CaptionType"
    SetProgressBarValue picProgress, uMin, uMax, uValue, Picture2, uShowCaption
End Property


Public Property Get Min() As Long

    Min = uMin
    
End Property

Public Property Let Min(ByVal New_Value As Long)

    uMin = New_Value
    PropertyChanged "Min"
    SetProgressBarValue picProgress, uMin, uMax, uValue, Picture2, uShowCaption
End Property

Public Property Get Max() As Long

    Max = uMax
    
End Property

Public Property Let Max(ByVal New_Value As Long)

    uMax = New_Value
    PropertyChanged "Max"
    SetProgressBarValue picProgress, uMin, uMax, uValue, Picture2, uShowCaption
End Property

Public Property Get Value() As Long

    Value = uValue
    
End Property

Public Property Let Value(ByVal New_Value As Long)

    uValue = New_Value
'    SetProgressBarValue picProgress, uMin, uMax, uValue, picSegment(uType), uShowCaption
    SetProgressBarValue picProgress, uMin, uMax, uValue, Picture2, uShowCaption
    PropertyChanged "Value"
    
End Property

Public Property Get ProgressType() As ProgressTypes
    ProgressType = uType
End Property

Public Property Let ProgressType(ByVal New_Value As ProgressTypes)
    uType = New_Value
    Call SetImage
    SetProgressBarValue picProgress, uMin, uMax, uValue, Picture2, uShowCaption
    PropertyChanged "ProgressType"
End Property

Private Sub SetImage()
    If uType = Custom Then
        If uCustomPicture Is Nothing Then Exit Sub
        picSegment(99).Cls
        Set picSegment(99).Picture = uCustomPicture
    End If
    
        Image1.Width = picSegment(uType).Width
        Image1.Height = picProgress.Height
        Image1.Picture = picSegment(uType).Picture
        Picture2.Width = Image1.Width
        Picture2.Height = Image1.Height
        Picture2.Refresh
        Picture2.PaintPicture Image1.Picture, 0, 0, Image1.Width \ Screen.TwipsPerPixelX, Image1.Height \ Screen.TwipsPerPixelY
End Sub

Private Sub SetProgressBarValue(pBar As PictureBox, vMin As Long, vMax As Long, vValue As Long, SourceSegment As PictureBox, Optional ShowCaption As Boolean = False)
Dim NumSegments As Long
Dim NumSegmentsToDraw As Long
Dim PercentageValue As Long
Dim ix As Integer
Dim ContainerColor As Long
Dim BarTop As Long
Dim BarLeft As Long
Dim BarWidth As Long
Dim BarHeight As Long
Dim TextMiddle As Long
Dim TextTop As Long

    If vMin > vValue Then Exit Sub
    If vValue > vMax Then Exit Sub
    If vMin < 0 Or vMax < 0 Or vValue < 0 Then Exit Sub
    If vMin = 0 And vMax = 0 Then Exit Sub
    
    picBuffer.Cls
    
    BarTop = pBar.Top \ Screen.TwipsPerPixelY
    BarLeft = pBar.Left \ Screen.TwipsPerPixelX
    BarWidth = (pBar.Width \ Screen.TwipsPerPixelX) - 1
    BarHeight = (pBar.Height \ Screen.TwipsPerPixelY) - 1
    
    If vValue >= vMin Then
    
        'segments on a bar
        NumSegments = BarWidth \ ((SourceSegment.Width \ Screen.TwipsPerPixelX) + 1)
        
        'Percentage value
        PercentageValue = Int((100 / (vMax - vMin)) * (vValue - vMin))
        
        'number of segments to draw for the percentage
        NumSegmentsToDraw = PercentageValue / (100 / NumSegments)
        
        For ix = 0 To NumSegmentsToDraw - 1
            BitBlt picBuffer.hDC, 2 + (ix * (1 + (SourceSegment.Width \ Screen.TwipsPerPixelX))), 0, SourceSegment.Width \ Screen.TwipsPerPixelX, SourceSegment.Height \ Screen.TwipsPerPixelY, SourceSegment.hDC, 0, 0, vbSrcCopy
        Next ix
        
        
        BitBlt pBar.hDC, 0, 0, picBuffer.Width, picBuffer.Height, picBuffer.hDC, 0, 0, vbSrcCopy
        
    End If
    
    'look at the container to remove the progress bar's corner pixels (for a rounded look)
 
    If uRounded Then
        BitBlt pBar.hDC, 0, 0, 1, 1, UserControl.hDC, BarLeft, BarTop, vbSrcCopy
        BitBlt pBar.hDC, 0, BarHeight, 1, 1, UserControl.hDC, BarLeft, BarTop + BarHeight, vbSrcCopy
        BitBlt pBar.hDC, BarWidth, 0, 1, 1, UserControl.hDC, BarLeft + BarWidth, BarTop, vbSrcCopy
        BitBlt pBar.hDC, BarWidth, BarHeight, 1, 1, UserControl.hDC, BarLeft + BarWidth, BarTop + BarHeight, vbSrcCopy
    End If

'    BitBlt pBar.hDC, 0, 0, 1, 1, pBar.Container.hDC, BarLeft, BarTop, vbSrcCopy
'    BitBlt pBar.hDC, 0, BarHeight, 1, 1, pBar.Container.hDC, BarLeft, BarTop + BarHeight, vbSrcCopy
'    BitBlt pBar.hDC, BarWidth, 0, 1, 1, pBar.Container.hDC, BarLeft + BarWidth, BarTop, vbSrcCopy
'    BitBlt pBar.hDC, BarWidth, BarHeight, 1, 1, pBar.Container.hDC, BarLeft + BarWidth, BarTop + BarHeight, vbSrcCopy

    If ShowCaption Then
        'pBar.Font = "MS Sans Serif"
        'pBar.FontSize = 8
        
        pBar.FontName = uFont.Name
        pBar.FontBold = uFont.Bold
        pBar.FontItalic = uFont.Italic
        pBar.FontSize = uFont.Size
        pBar.ForeColor = vbBlack
        
        
        Select Case uCaptionPositionX
            Case pbarLeft
                TextMiddle = 4
            Case pbarCenter
                If uCaptionType = 0 Then
                    TextMiddle = (BarWidth - pBar.TextWidth(PercentageValue & "%")) \ 2
                Else
                    TextMiddle = (BarWidth - pBar.TextWidth(uCaption)) \ 2
                End If
            Case pbarRight
                If uCaptionType = 0 Then
                    TextMiddle = BarWidth - (pBar.TextWidth(PercentageValue & "%") + 4)
                Else
                    TextMiddle = BarWidth - (pBar.TextWidth(uCaption) + 4)
                End If
        End Select
        
        Select Case uCaptionPositionY
            Case pbarTop
                TextTop = 0
            Case pbarMiddle
                If uCaptionType = 0 Then
                    TextTop = (BarHeight - pBar.TextHeight(PercentageValue & "%")) \ 2
                Else
                    TextTop = (BarHeight - pBar.TextHeight(uCaption)) \ 2
                End If
            Case pbarBottom
                If uCaptionType = 0 Then
                    TextTop = BarHeight - pBar.TextHeight(PercentageValue & "%")
                Else
                    TextTop = BarHeight - pBar.TextHeight(uCaption)
                End If
        End Select
        
        If uCaptionType = 0 Then
            pBar.CurrentX = TextMiddle - 1
            pBar.CurrentY = TextTop
            pBar.Print PercentageValue & "%"
        
            pBar.CurrentX = TextMiddle + 1
            pBar.CurrentY = TextTop
            pBar.Print PercentageValue & "%"
        
            pBar.CurrentX = TextMiddle
            pBar.CurrentY = TextTop - 1
            pBar.Print PercentageValue & "%"
        
            pBar.CurrentX = TextMiddle
            pBar.CurrentY = TextTop + 1
            pBar.Print PercentageValue & "%"
        
            pBar.ForeColor = uCaptionColor
        
            pBar.CurrentX = TextMiddle
            pBar.CurrentY = TextTop
            pBar.Print PercentageValue & "%"
        Else
            pBar.CurrentX = TextMiddle - 1
            pBar.CurrentY = TextTop
            pBar.Print uCaption
        
            pBar.CurrentX = TextMiddle + 1
            pBar.CurrentY = TextTop
            pBar.Print uCaption
        
            pBar.CurrentX = TextMiddle
            pBar.CurrentY = TextTop - 1
            pBar.Print uCaption
        
            pBar.CurrentX = TextMiddle
            pBar.CurrentY = TextTop + 1
            pBar.Print uCaption
        
            pBar.ForeColor = uCaptionColor
        
            pBar.CurrentX = TextMiddle
            pBar.CurrentY = TextTop
            pBar.Print uCaption
        
        End If
    End If
    pBar.Refresh

End Sub

Private Sub UserControl_Initialize()
    Set uFont = New StdFont
    uRounded = True
    uCaptionColor = vbCyan
    uCaptionPositionX = pbarLeft
    uCaptionPositionY = pbarMiddle
    uShowCaption = True
    uCaptionType = Percentage
    uCaption = ""
    uValue = 0
    uType = Standard
    uMin = 0
    uMax = 100
    Call SetImage
    SetProgressBarValue picProgress, uMin, uMax, uValue, Picture2, uShowCaption
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Set uFont = .ReadProperty("Font", Ambient.Font)
        Set uCustomPicture = .ReadProperty("CustomPicture", Nothing)
        uRounded = .ReadProperty("Rounded", True)
        uCaptionColor = .ReadProperty("CaptionColor", vbCyan)
        uCaptionPositionX = .ReadProperty("CaptionXPosition", pbarLeft)
        uCaptionPositionY = .ReadProperty("CaptionYPosition", pbarMiddle)
        uShowCaption = .ReadProperty("ShowCaption", True)
        uCaptionType = .ReadProperty("CaptionType", Percentage)
        uCaption = .ReadProperty("Caption", "")
        uValue = .ReadProperty("Value", 0)
        uType = .ReadProperty("ProgressType", Standard)
        uMin = .ReadProperty("Min", 0)
        uMax = .ReadProperty("Max", 100)
    End With
    If Not (uCustomPicture Is Nothing) Then
        picSegment(99).Cls
        picSegment(99).Width = uCustomPicture.Width
        picSegment(99).Height = uCustomPicture.Height
        Set picSegment(99).Picture = uCustomPicture
    End If
    Call SetImage
    SetProgressBarValue picProgress, uMin, uMax, uValue, Picture2, uShowCaption
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Font", uFont, Ambient.Font
        .WriteProperty "CustomPicture", uCustomPicture, Nothing
        .WriteProperty "Rounded", uRounded, True
        .WriteProperty "CaptionColor", uCaptionColor, vbCyan
        .WriteProperty "CaptionXPosition", uCaptionPositionX, pbarLeft
        .WriteProperty "CaptionYPosition", uCaptionPositionY, pbarMiddle
        .WriteProperty "ShowCaption", uShowCaption, True
        .WriteProperty "CaptionType", uCaptionType, Percentage
        .WriteProperty "Caption", uCaption, ""
        .WriteProperty "Value", uValue, 0
        .WriteProperty "ProgressType", uType, Standard
        .WriteProperty "Min", uMin, 0
        .WriteProperty "Max", uMax, 100
    End With
End Sub
Private Sub UserControl_Resize()
    picBuffer.Height = UserControl.Height
    picBuffer.Width = UserControl.Width
    picProgress.Move 0, 0, UserControl.Width, UserControl.Height
    Call SetImage
    SetProgressBarValue picProgress, uMin, uMax, uValue, Picture2, uShowCaption
End Sub
