VERSION 5.00
Begin VB.UserControl dShadow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   930
   ScaleHeight     =   555
   ScaleWidth      =   930
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   480
      Top             =   0
   End
   Begin VB.PictureBox picdc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   465
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "dShadow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'  this idea was expanded from Brian Kurtyka's
' Self updaing transparent / translucent form
' using the alphablend (THX a mil brian, you taught
' me alphablending just from that one examble...
' after seeing that i messed aroung to try to draw
' directly to the screen, and then ofcourse had to
' learn how to repaint the screen..
' im a big fan of 3d-ish techniques, so how about
' a form that seems to be floating above the desktop?
' ok so its not perfectly fluid, but i optimized
' it to the best of my abilities, although i invite
' any of our more advanced coders to make it even more
' err..."fluidic"(?)

' The alphablend technique is nowhere near new to PSC
' but the idea of a form's dropshadow to be unique,
' and the RedrawWindow API is new as well, i found
' no examples of it on PSC, nor any working implintations
' in the entire net (probly cause of my poor searching
' skills ;)

'anyway, if you like it, let me know, although voting
' is pretty much pointless, this really wont compete with the bigboys.
' putting my name in the creds would be cool..
' and last but not least, you can check out more of
'my work at www.bbsdaze.com/~dark
' and my other interests (music) as mp3.com/Contrast (my band)


Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Dim tRect As RECT, DeskHDc As Long, TimeX As Boolean, secX As Integer
Dim AlfB1 As RECT, AlfB2 As RECT
'Default Property Values:
Const m_def_cthru = 150
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
Const m_def_AutoRedraw = 0
Const m_def_hDC = 0
Const m_def_hWnd = 0
Const m_def_ScaleWidth = 0
Const m_def_ScaleTop = 0
Const m_def_ScaleMode = 0
Const m_def_ScaleLeft = 0
Const m_def_ScaleHeight = 0
Const m_def_width = 5
Const m_def_height = 5
'Property Variables:
Dim m_cthru As Integer
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
Dim m_AutoRedraw As Boolean
Dim m_hDC As Long
Dim m_hWnd As Long
Dim m_Picture As Picture
Dim m_ScaleWidth As Single
Dim m_ScaleTop As Single
Dim m_ScaleMode As Integer
Dim m_ScaleLeft As Single
Dim m_ScaleHeight As Single
Dim m_width As Long
Dim m_height As Long
'Event Declarations:
Event ablend()
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Paint()
Attribute Paint.VB_Description = "Occurs when any part of a form or PictureBox control is moved, enlarged, or exposed."
Dim xx As Integer, yy As Integer





'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get BackColor() As Long
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
     
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Returns/sets the output from a graphics method to a persistent bitmap."
    AutoRedraw = m_AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    m_AutoRedraw = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get hdc() As Long
Attribute hdc.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hdc = m_hDC
End Property

Public Property Let hdc(ByVal New_hDC As Long)
    m_hDC = New_hDC
    PropertyChanged "hDC"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = m_hWnd
End Property

Public Property Let hwnd(ByVal New_hWnd As Long)
    m_hWnd = New_hWnd
    PropertyChanged "hWnd"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set m_Picture = New_Picture
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12
Public Function ScaleY(ByVal height As Single, Optional ByVal FromScale As Variant, Optional ByVal ToScale As Variant) As Single
Attribute ScaleY.VB_Description = "Converts the value for the height of a Form, PictureBox, or Printer from one unit of measure to another."

End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12
Public Function ScaleX(ByVal width As Single, Optional ByVal FromScale As Variant, Optional ByVal ToScale As Variant) As Single
Attribute ScaleX.VB_Description = "Converts the value for the width of a Form, PictureBox, or Printer from one unit of measure to another."

End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_Description = "Returns/sets the number of units for the horizontal measurement of an object's interior."
    ScaleWidth = m_ScaleWidth
End Property

Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
    m_ScaleWidth = New_ScaleWidth
    PropertyChanged "ScaleWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get ScaleTop() As Single
Attribute ScaleTop.VB_Description = "Returns/sets the vertical coordinates for the top edges of an object."
    ScaleTop = m_ScaleTop
End Property

Public Property Let ScaleTop(ByVal New_ScaleTop As Single)
    m_ScaleTop = New_ScaleTop
    PropertyChanged "ScaleTop"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get ScaleMode() As Integer
Attribute ScaleMode.VB_Description = "Returns/sets a value indicating measurement units for object coordinates when using graphics methods or positioning controls."
    ScaleMode = m_ScaleMode
End Property

Public Property Let ScaleMode(ByVal New_ScaleMode As Integer)
    m_ScaleMode = New_ScaleMode
    PropertyChanged "ScaleMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get ScaleLeft() As Single
Attribute ScaleLeft.VB_Description = "Returns/sets the horizontal coordinates for the left edges of an object."
    ScaleLeft = m_ScaleLeft
End Property

Public Property Let ScaleLeft(ByVal New_ScaleLeft As Single)
    m_ScaleLeft = New_ScaleLeft
    PropertyChanged "ScaleLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_Description = "Returns/sets the number of units for the vertical measurement of an object's interior."
    ScaleHeight = m_ScaleHeight
End Property

Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
    m_ScaleHeight = New_ScaleHeight
    PropertyChanged "ScaleHeight"
End Property

'The Underscore following "Scale" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Scale_(Optional X1 As Variant, Optional Y1 As Variant, Optional X2 As Variant, Optional Y2 As Variant)
    
End Sub


Public Sub ablend()
If secX > 0 Then Exit Sub

GetWindowRect UserControl.Parent.hwnd, Outer
DeskHDc = GetDC(0)
DoEvents
AlphaBlending DeskHDc, (UserControl.Parent.Left / Screen.TwipsPerPixelX) + 32 _
, (UserControl.Parent.Top + UserControl.Parent.height) / Screen.TwipsPerPixelY _
, UserControl.Parent.width / Screen.TwipsPerPixelX _
, 32 _
, picdc.hdc, 0 _
, 0 _
, 1, 1, 130
AlfB1.Left = (UserControl.Parent.Left / Screen.TwipsPerPixelX) + 32
AlfB1.Top = (UserControl.Parent.Top + UserControl.Parent.height) / Screen.TwipsPerPixelY
AlfB1.Right = (UserControl.Parent.Left + UserControl.Parent.width) / Screen.TwipsPerPixelX
AlfB1.Bottom = AlfB1.Top + 32

AlfB2.Left = Outer.Right
AlfB2.Top = (UserControl.Parent.Top / Screen.TwipsPerPixelY) + 32
AlfB2.Right = AlfB2.Left + 32
AlfB2.Bottom = ((UserControl.Parent.Top + UserControl.Parent.height) / Screen.TwipsPerPixelY) + 32
DoEvents
AlphaBlending DeskHDc, Outer.Right _
, (UserControl.Parent.Top / Screen.TwipsPerPixelY) + 32 _
, 32 _
, (UserControl.Parent.height / Screen.TwipsPerPixelY) - 32 _
, picdc.hdc, 0 _
, 0 _
, 1, 1, 130

DoEvents
ReleaseDC 0&, DeskHDc
DoEvents
End Sub




Private Sub Timer1_Timer()
DoEvents

If (xx = UserControl.Parent.Left) And (yy = UserControl.Parent.Top) Then
Exit Sub
Else
If TimeX = True Then Exit Sub
TimeX = True

RedrawWindow 0&, AlfB1, 0&, 135
DoEvents
RedrawWindow 0&, AlfB2, 0&, 135
DoEvents

DoEvents
DoEvents

xx = UserControl.Parent.Left
yy = UserControl.Parent.Top
Call ablend
End If
TimeX = False
End Sub


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
    m_AutoRedraw = m_def_AutoRedraw
    m_hDC = m_def_hDC
    m_hWnd = m_def_hWnd
    Set m_Picture = LoadPicture("")
    m_ScaleWidth = m_def_ScaleWidth
    m_ScaleTop = m_def_ScaleTop
    m_ScaleMode = m_def_ScaleMode
    m_ScaleLeft = m_def_ScaleLeft
    m_ScaleHeight = m_def_ScaleHeight
    m_cthru = m_def_cthru
 m_width = m_def_width
 m_height = m_def_height

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_AutoRedraw = PropBag.ReadProperty("AutoRedraw", m_def_AutoRedraw)
    m_hDC = PropBag.ReadProperty("hDC", m_def_hDC)
    m_hWnd = PropBag.ReadProperty("hWnd", m_def_hWnd)
    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
     m_ScaleTop = PropBag.ReadProperty("ScaleTop", m_def_ScaleTop)
    m_ScaleMode = PropBag.ReadProperty("ScaleMode", m_def_ScaleMode)
    m_ScaleLeft = PropBag.ReadProperty("ScaleLeft", m_def_ScaleLeft)
   m_width = PropBag.ReadProperty("Width", m_def_width)
    m_height = PropBag.ReadProperty("Height", m_def_height)
   m_ScaleWidth = PropBag.ReadProperty("ScaleWidth", m_def_ScaleWidth)
    m_ScaleHeight = PropBag.ReadProperty("ScaleHeight", m_def_ScaleHeight)
    m_cthru = PropBag.ReadProperty("cthru", m_def_cthru)
End Sub

Private Sub UserControl_Resize()
picdc.width = UserControl.width + 10
picdc.height = UserControl.height + 10

End Sub

Private Sub UserControl_Terminate()
On Error Resume Next
Timer1.Enabled = False
tRect.Top = 0
tRect.Left = 0
tRect.Right = Screen.width / Screen.TwipsPerPixelX
tRect.Bottom = Screen.height / Screen.TwipsPerPixelY
RedrawWindow 0&, tRect, 0&, 135
DoEvents
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("AutoRedraw", m_AutoRedraw, m_def_AutoRedraw)
    Call PropBag.WriteProperty("hDC", m_hDC, m_def_hDC)
    Call PropBag.WriteProperty("hWnd", m_hWnd, m_def_hWnd)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)

    Call PropBag.WriteProperty("ScaleTop", m_ScaleTop, m_def_ScaleTop)
    Call PropBag.WriteProperty("ScaleMode", m_ScaleMode, m_def_ScaleMode)
    Call PropBag.WriteProperty("ScaleLeft", m_ScaleLeft, m_def_ScaleLeft)
    Call PropBag.WriteProperty("Width", m_width, m_def_width)
    Call PropBag.WriteProperty("Height", m_height, m_def_height)
     Call PropBag.WriteProperty("ScaleWidth", m_ScaleWidth, m_def_ScaleWidth)
    Call PropBag.WriteProperty("ScaleHeight", m_ScaleHeight, m_def_ScaleHeight)
    Call PropBag.WriteProperty("cthru", m_cthru, m_def_cthru)
End Sub

Public Property Get height() As Long
    height = m_height
End Property

Public Property Let height(ByVal New_Height As Long)
    m_height = New_Height
    PropertyChanged "Height"
End Property
Public Property Get width() As Long
    width = m_width
End Property

Public Property Let width(ByVal New_width As Long)
    m_width = New_width
    PropertyChanged "width"
End Property
Public Property Get cthru() As Integer
    cthru = m_cthru
End Property

Public Property Let cthru(ByVal New_cthru As Integer)
    m_cthru = New_cthru
    PropertyChanged "cthru"
End Property



