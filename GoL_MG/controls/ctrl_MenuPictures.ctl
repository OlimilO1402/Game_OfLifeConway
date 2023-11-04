VERSION 5.00
Begin VB.UserControl ctrl_MenuPictures 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   1650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1605
   InvisibleAtRuntime=   -1  'True
   MaskPicture     =   "ctrl_MenuPictures.ctx":0000
   Picture         =   "ctrl_MenuPictures.ctx":04E8
   ScaleHeight     =   110
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   107
   ToolboxBitmap   =   "ctrl_MenuPictures.ctx":07ED
   Begin VB.PictureBox BitmapsChecked 
      BorderStyle     =   0  'Kein
      Height          =   480
      Index           =   0
      Left            =   0
      Picture         =   "ctrl_MenuPictures.ctx":0AFF
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   0
      Width           =   480
   End
   Begin VB.PictureBox BitmapsUnchecked 
      BorderStyle     =   0  'Kein
      Height          =   480
      Index           =   0
      Left            =   0
      Picture         =   "ctrl_MenuPictures.ctx":0E04
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "ctrl_MenuPictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
        
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, _
                                                            ByVal nPosition As Long, _
                                                            ByVal wFlags As Long, _
                                                            ByVal hBitmapUnchecked As Long, _
                                                            ByVal hBitmapChecked As Long) As Long

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
        
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, _
                                                                        ByVal wFlags As Long, _
                                                                        ByVal wIDNewItem As Long, _
                                                                        ByVal lpNewItem As String) As Long

Private Const MF_STRING = &H0&
Private Const MF_SEPARATOR = &H800&
Private Const MF_BYPOSITION = &H400&
Private Const MF_BITMAP = &H4&
Private Const MF_BYCOMMAND = &H0&

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal X As Long, _
                                                    ByVal Y As Long, _
                                                    ByVal nWidth As Long, _
                                                    ByVal nHeight As Long, _
                                                    ByVal hSrcDC As Long, _
                                                    ByVal xSrc As Long, _
                                                    ByVal ySrc As Long, _
                                                    ByVal nSrcWidth As Long, _
                                                    ByVal nSrcHeight As Long, _
                                                    ByVal dwRop As Long) As Long


Private m_lng_MenubarHandle As Long
Private m_lng_MenuHandles(0 To 255, 0 To 255) As Long
Private m_lng_BitmapsCount As Long

Private m_lng_SRCWidth As Long
Private m_lng_SRCHeight As Long
Private m_lng_SRCLeft As Long
Private m_lng_SRCTop As Long
Private m_lng_DSTWidth As Long
Private m_lng_DSTHeight As Long

Public Property Get DestWidth() As Long
    DestWidth = m_lng_DSTWidth
End Property
Public Property Let DestWidth(ByVal New_Value As Long)
    m_lng_DSTWidth = New_Value
End Property

Public Property Get DestHeight() As Long
    DestHeight = m_lng_DSTHeight
End Property
Public Property Let DestHeight(ByVal New_Value As Long)
    m_lng_DSTHeight = New_Value
End Property

Public Property Get SourceWidth() As Long
    SourceWidth = m_lng_SRCWidth
End Property
Public Property Let SourceWidth(ByVal New_Value As Long)
    m_lng_SRCWidth = New_Value
End Property

Public Property Get SourceHeight() As Long
    SourceHeight = m_lng_SRCHeight
End Property
Public Property Let SourceHeight(ByVal New_Value As Long)
    m_lng_SRCHeight = New_Value
End Property


Public Property Get SourceLeft() As Long
    SourceLeft = m_lng_SRCLeft
End Property
Public Property Let SourceLeft(ByVal New_Value As Long)
    m_lng_SRCLeft = New_Value
End Property

Public Property Get SourceTop() As Long
    SourceTop = m_lng_SRCTop
End Property
Public Property Let SourceTop(ByVal New_Value As Long)
    m_lng_SRCTop = New_Value
End Property


Public Property Get BitmapsCount() As Long
    BitmapsCount = m_lng_BitmapsCount
End Property

Private Sub UserControl_Initialize()
    m_lng_DSTWidth = 14
    m_lng_DSTHeight = 14
    
    m_lng_SRCWidth = 32
    m_lng_SRCHeight = 32
    m_lng_SRCLeft = 0
    m_lng_SRCTop = 0
    
    BitmapsUnchecked(0).Width = m_lng_DSTWidth
    BitmapsUnchecked(0).Height = m_lng_DSTHeight
    BitmapsChecked(0).Width = m_lng_DSTWidth
    BitmapsChecked(0).Height = m_lng_DSTHeight
    m_lng_BitmapsCount = 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_lng_DSTWidth = 14
    m_lng_DSTHeight = 14
    
    m_lng_SRCWidth = 32
    m_lng_SRCHeight = 32
    m_lng_SRCLeft = 0
    m_lng_SRCTop = 0
    
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 32 * Screen.TwipsPerPixelX
    UserControl.Height = 32 * Screen.TwipsPerPixelY
End Sub

Public Sub Add_MenuIcon(ByVal MenuID As String, ByRef BitmapUnchecked As IPictureDisp, Optional ByRef BitmapChecked As IPictureDisp)
'ByVal MenuIndex As Long, ByVal ItemIndex As Long, ByRef BitmapUnchecked As IPictureDisp, Optional ByRef BitmapChecked As IPictureDisp)
On Error Resume Next 'ByVal Depth As Long,
Dim MenuTree() As String
Dim i As Long
Dim Level As Long

    If Not BitmapUnchecked Is Nothing Then
        m_lng_MenubarHandle = GetMenu(UserControl.Parent.hwnd)
        MenuTree = Split(MenuID, ",")
        m_lng_MenuHandles(BitmapsCount, 0) = m_lng_MenubarHandle
        Level = UBound(MenuTree())
        For i = 0 To Level - 1
            m_lng_MenuHandles(BitmapsCount, i + 1) = GetSubMenu(Val(m_lng_MenuHandles(BitmapsCount, i)), Val(MenuTree(i)))
        Next i
      
        If BitmapChecked Is Nothing Then Set BitmapChecked = BitmapUnchecked

        Load BitmapsUnchecked(BitmapsCount + 1)

        m_lng_BitmapsCount = BitmapsUnchecked.UBound
        Load BitmapsChecked(BitmapsCount)
        With BitmapsUnchecked(BitmapsCount)
            .AutoRedraw = True
            .PaintPicture BitmapUnchecked, 0, 0, m_lng_DSTWidth, m_lng_DSTHeight, SourceLeft, SourceTop, SourceWidth, SourceHeight, vbSrcCopy
            Set .Picture = .Image
        End With
        With BitmapsChecked(BitmapsUnchecked.UBound)
            .AutoRedraw = True
            .PaintPicture BitmapChecked, 0, 0, m_lng_DSTWidth, m_lng_DSTHeight, SourceLeft, SourceTop, SourceWidth, SourceHeight, vbSrcCopy
            Set .Picture = .Image
        End With

        Call SetMenuItemBitmaps(m_lng_MenuHandles(BitmapsCount - 1, Level), Val(MenuTree(Level)), MF_BYPOSITION, BitmapsUnchecked(BitmapsCount).Picture, BitmapsChecked(BitmapsCount).Picture)

    End If
End Sub

