VERSION 5.00
Begin VB.UserControl nPolyButton 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   735
   DrawStyle       =   6  'Inside Solid
   DrawWidth       =   4
   PropertyPages   =   "nPolyButton.ctx":0000
   ScaleHeight     =   780
   ScaleWidth      =   735
End
Attribute VB_Name = "nPolyButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit


Private Type POINTAPI
    x As Long
    y As Long
End Type

Public Enum POLYDRAW
    FocusRect = 1
    Ordinary
    MouseDown
End Enum

'API Declares for setting the button shape as round
'these functions cut the usercontrol and give it a round shape
'so that only the circle is opaque, the rest of the region is transparent


' This code assumes that your form's scalemode is set to vbTwips.
' API functions for creating and setting regions.

Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal x3 As Long, ByVal y3 As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long


' Constants used when combining regions.
Private Const RGN_AND = 1
Private Const RGN_COPY = 5
Private Const RGN_DIFF = 4
Private Const RGN_MAX = RGN_COPY
Private Const RGN_MIN = RGN_AND
Private Const RGN_OR = 2
Private Const RGN_XOR = 3

Private Const ALTERNATE = 1
Private Const WINDING = 2

Private cFace As Long, cLight As Long, cHighLight As Long, cShadow As Long, cDarkShadow As Long, cText As Long

Private Const PR_COLOR_BTNFACE = 15
Private Const PR_COLOR_BTNSHADOW = 16
Private Const PR_COLOR_BTNTEXT = 18
Private Const PR_COLOR_BTNHIGHLIGHT = 20
Private Const PR_COLOR_BTNDKSHADOW = 21
Private Const PR_COLOR_BTNLIGHT = 22

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long


'storage for the first set of points for lines
Private paPoints(32) As POINTAPI


'Default Property Values:
Const m_def_ForeColor = 0
Const m_def_Caption = "polyBtn"
'Const m_def_Radius = 0
'Property Variables:
Dim m_ForeColor As OLE_COLOR
Dim m_Caption As String
'Dim m_Radius As Integer
Dim lSides As Long ' the number of sides on the form
Dim lRotation As Long ' how many degrees from parallel to top to rotate control

Dim test As Boolean

'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp

Private Sub SetColors()
'this function is taken form another upload in PSC,
'the Gurhan Button by 'Gurhan KARTAL
'Thanks Gurhan

    cFace = GetSysColor(PR_COLOR_BTNFACE)
    cShadow = GetSysColor(PR_COLOR_BTNSHADOW)
    cLight = GetSysColor(PR_COLOR_BTNLIGHT)
    cDarkShadow = GetSysColor(PR_COLOR_BTNDKSHADOW)
    cHighLight = GetSysColor(PR_COLOR_BTNHIGHLIGHT)
    cText = GetSysColor(PR_COLOR_BTNTEXT)

End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
    draw FocusRect
End Sub

Private Sub UserControl_GotFocus()
    draw FocusRect
End Sub

Private Sub UserControl_Initialize()
    ReCalcPoints
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    Select Case KeyCode
      Case 32
        Call UserControl_MouseDown(0, 0, 0, 0)
      Case Else
    End Select
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyUp(KeyCode, Shift)
    Select Case KeyCode
      Case 32
        Call UserControl_MouseUp(0, 0, 0, 0)
        Call UserControl_Click
      Case Else
    End Select

End Sub

Private Sub UserControl_LostFocus()
    Cls
    UserControl_Resize
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
    draw MouseDown
    Select Case Button
        Case vbLeftButton
        Case vbRightButton
        Case vbMiddleButton
        Case Else
    End Select
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
    draw Ordinary
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000000)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    
    lSides = PropBag.ReadProperty("Sides", 3)
    lRotation = PropBag.ReadProperty("Rotation", 0)
    
    SetColors
    draw Ordinary
    test = True
End Sub

Private Function draw(Arg As POLYDRAW, Optional bfocusRectMDwn As Boolean = False)
'this function draws the button. it uses the circle method to draw
'========================================================================================
'written by Praveen Menon
'12th Mrach 2002
'Last Modified Date
'12th March 2002

'========================================================================================
'arguments
'========================================================================================
    'strArg ==> used to determine which state of the button is to be drawn
    'three arguments can be there
        '   1) FocusRect ===> draws the focus rect of the button
        '   2) Ordinary ===> draws the ordinary or the mouseup state of the button
        '   3) MouseDown ===> draws the mousedown state or the keydown state of thebutton
    'bfocusRectMDwn ===> used to determine whether the focus rect is for
        '                the mousedown state or the mouseup state
'========================================================================================

  'Dim rad1 As Integer
  'Dim rad2 As Integer

    Select Case Arg
    
        Case FocusRect
            UserControl.DrawStyle = 2
            DrawWidth = 1
            If Not bfocusRectMDwn Then
                DrawPoly -30, 0, vbHighlight
            Else
                DrawPoly -5, 0, vbHighlight
            End If
            UserControl.DrawStyle = 6
            DrawWidth = 4
     
        Case Ordinary
            Cls
            DrawPoly 1, 7, vbWhite
            DrawPoly -50, 0, cDarkShadow
            DrawPoly -20, 0, cShadow
            UserControl.ForeColor = m_ForeColor
            TextOut UserControl.hdc, (ScaleWidth - UserControl.TextWidth(m_Caption)) / 2 / Screen.TwipsPerPixelX, (ScaleHeight - UserControl.TextHeight(m_Caption)) / 2 / Screen.TwipsPerPixelY, m_Caption, Len(m_Caption)
      
        Case MouseDown
            Cls
            DrawPoly 1, 10, cShadow
            DrawWidth = 4
            UserControl.ForeColor = m_ForeColor
            TextOut UserControl.hdc, (ScaleWidth - UserControl.TextWidth(m_Caption)) / 2 / Screen.TwipsPerPixelX + 2, (ScaleHeight - UserControl.TextHeight(m_Caption)) / 2 / Screen.TwipsPerPixelY + 2, m_Caption, Len(m_Caption)
            draw FocusRect, True
      
      Case Else
            MsgBox "oops!"
    End Select

End Function

Private Sub UserControl_Resize()
Dim lrgn As Long
Dim i As Long

    ReCalcPoints
    Cls
    If Width > Height Then Height = Width Else Width = Height
    lrgn = CreatePolygonRgn(paPoints(0), lSides, WINDING)
    
    If IsNull(lrgn) Then
        MsgBox "no lock on region"
    End If
    
    i = SetWindowRgn(UserControl.hWnd, lrgn, True)
    
    If i = 0 Then MsgBox "error setting windows region"
    draw Ordinary
    '?m_Radius = Height

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    draw "Ordinary"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "General"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = "General"
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
End Property

Public Property Get Sides() As Integer
Attribute Sides.VB_ProcData.VB_Invoke_Property = "General"
    Sides = lSides
End Property
Public Property Let Sides(ByVal New_sides As Integer)
    lSides = New_sides
    If lSides < 3 Then
        lSides = 3
    End If
    PropertyChanged "Sides"
    UserControl_Resize
End Property
Public Property Get Rotation() As Integer
Attribute Rotation.VB_ProcData.VB_Invoke_Property = "General"
    Rotation = lRotation
End Property
Public Property Let Rotation(ByVal New_rotation As Integer)
    lRotation = New_rotation
    PropertyChanged "Rotation"
    UserControl_Resize
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

    Set UserControl.Font = Ambient.Font
    m_Caption = m_def_Caption
    lRotation = 0
    lSides = 3

    m_ForeColor = m_def_ForeColor
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000000)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Rotation", lRotation, 0)
    Call PropBag.WriteProperty("Sides", lSides, 3)
    
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    draw "Ordinary"
End Property

Private Sub DrawPoly(Xoffset As Long, Yoffset As Long, ByVal Color As ColorConstants)
Dim i As Long
Dim x0 As Long, y0 As Long, x1 As Long, y1 As Long
Dim rads(2) As Double

    For i = 0 To lSides - 1
        rads(1) = (((360 / lSides) * i) + lRotation) * (3.14159265358979 / 180) 'calculate the angle from the center of the control to the point
        rads(2) = (((360 / lSides) * (i + 1)) + lRotation) * (3.14159265358979 / 180) 'calculate the angle from the center of the control to the point
        x0 = (Width / (2)) * (1 + Cos(rads(1))) + Xoffset        '  form of       scalar * (offset + vector)     ''' offset = dist from 0,0 , i.e. the center of the control
        y0 = (Height / (2)) * (1 + Sin(rads(1))) + Yoffset
        x1 = (Width / (2)) * (1 + Cos(rads(2))) + Xoffset
        y1 = (Height / (2)) * (1 + Sin(rads(2))) + Yoffset
        UserControl.Line (x0, y0)-(x1, y1), Color 'vbGreen
    Next
End Sub

Private Sub ReCalcPoints()
Dim i As Long
Dim rads As Double
    If lSides < 3 Then
        lSides = 3 ' prevent invalid sizes
    
        If test = False Then
                Exit Sub
        End If
    test = test
    End If
    
    If Width < Height Then Height = Width Else Width = Height ' make it square

    For i = 0 To lSides - 1
        rads = (((360 / lSides) * i) + lRotation) * (3.14159265358979 / 180) 'calculate the angle from the center of the control to the point
        paPoints(i).x = (UserControl.Width / (2 * Screen.TwipsPerPixelX)) * _
                 (1 + Cos(rads))
        
        paPoints(i).y = (UserControl.Height / (2 * Screen.TwipsPerPixelY)) * _
                 (1 + Sin(rads))
    Next
        paPoints(lSides).x = paPoints(0).x
        paPoints(lSides).y = paPoints(0).y
End Sub

