VERSION 5.00
Begin VB.UserControl NeuralProcessor 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ForeColor       =   &H00FFFFFF&
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.VScrollBar vS 
      Height          =   2295
      Index           =   2
      LargeChange     =   10
      Left            =   0
      Max             =   360
      Min             =   -360
      SmallChange     =   2
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.VScrollBar vS 
      Height          =   2295
      Index           =   1
      LargeChange     =   10
      Left            =   120
      Max             =   360
      Min             =   -360
      SmallChange     =   2
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.VScrollBar vS 
      Height          =   2295
      Index           =   0
      LargeChange     =   10
      Left            =   240
      Max             =   360
      Min             =   -360
      SmallChange     =   2
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "NeuralProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Public Type typNeuron
Hit As Boolean
Value(0 To 2) As Integer
End Type

Private NeuralNet() As typNeuron

Const m_def_nnWidth = 1
Const m_def_nnHeight = 1
Const m_def_nnDepth = 1

Dim m_nnWidth As Long
Dim m_nnHeight As Long
Dim m_nnDepth As Long

Private vStopNet As Boolean

Private dX, dY, LastVS(0 To 2)

Event Click()
Event DblClick()
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Sub About()
Attribute About.VB_UserMemId = -552
MsgBox "This NeuralNet was created by: Cory J. Geesaman", vbInformation, "Created By"
MsgBox "This was created because i got bored and had a passing thought of making something like this.", vbInformation, "Why This Was Made"
MsgBox "The Background for this control was made by Richard Gardener", vbInformation, "Background"
MsgBox "If you have any ideas for implementing this email: cory@geesaman.com, I am very-much open to suggestions for implementation into computer AI for charactors in my game.", vbInformation, "Got Ideas For The NeuralNet"
MsgBox "Visit: http://www.naven.net/ for information on NAVEN, a very neat game I am making.", vbInformation, "Visit My Site"
End Sub

Private Sub nDrawLine(X1 As Long, X2 As Long, Y1 As Long, Y2 As Long, Color As Long, Optional Width As Long)
If Width < 1 Then Width = 1
Dim p_hPen As Long, p_old_hPen As Long, i_Point As POINTAPI
p_hPen = CreatePen(0, CLng(1), Color)
p_old_hPen = SelectObject(UserControl.hdc, p_hPen)
MoveToEx UserControl.hdc, X1, Y1, i_Point
LineTo UserControl.hdc, X2, Y2
DeleteObject p_hPen
DeleteObject p_old_hPen
End Sub

Public Function GetNNValue(X, Y, Z, V) As Integer
GetNNValue = NeuralNet(X, Y, Z).Value(V)
End Function

Public Sub SetNNValue(X, Y, Z, V, Value)
NeuralNet(X, Y, Z).Value(V) = Value
End Sub

Private Sub Draw3dPixel(ByVal ivertX As Double, ByVal ivertY As Double, ByVal ivertZ As Double, Color As Long)
Dim cx As Single, cy As Single
Dim sX As Single, sY As Single, sZ As Single
Dim tempx As Single, tempy As Single, tempz As Single


  cx = UserControl.ScaleWidth / 2
  cy = UserControl.ScaleHeight / 2
  s = 361
  
  '// Camera Coordinates
  If vS(0) = 0 Then sX = 1 Else sX = vS(0)
  If vS(1) = 0 Then sY = 1 Else sY = vS(1)
  If vS(2) = 0 Then sZ = 1 Else sZ = vS(2)
  
 
  tempx = ivertX + sX '// transform origin
  tempy = ivertY + sY '// to screen center
  tempz = ivertZ + sZ '// to screen center
    
      
  PERS = s / (s + tempz) '//Get perspective factor

  View_Plot_X = cx + tempx * PERS
  View_Plot_Y = cy - tempy * PERS
  
  DeleteObject SetPixel(UserControl.hdc, View_Plot_X, View_Plot_Y, Color)
End Sub

Private Sub Draw3dLine(ByVal ivertX As Double, ByVal ivertY As Double, ByVal ivertZ As Double, ByVal ivertX2 As Double, ByVal ivertY2 As Double, ByVal ivertZ2 As Double, Color As Long)
Dim cx As Single, cy As Single
Dim sX As Single, sY As Single, sZ As Single
Dim tempx As Single, tempy As Single, tempz As Single
Dim tempx2 As Single, tempy2 As Single, tempz2 As Single
Dim View_Plot_X As Long, View_Plot_X2 As Long, View_Plot_Y As Long, View_Plot_Y2 As Long


  cx = UserControl.ScaleWidth / 2
  cy = UserControl.ScaleHeight / 2
  s = 361
  
  '// Camera Coordinates
  If vS(0) = 0 Then sX = 1 Else sX = vS(0)
  If vS(1) = 0 Then sY = 1 Else sY = vS(1)
  If vS(2) = 0 Then sZ = 1 Else sZ = vS(2)
  
 
  tempx = ivertX + sX '// transform origin
  tempy = ivertY + sY '// to screen center
  tempz = ivertZ + sZ '// to screen center
  tempx2 = ivertX2 + sX '// transform origin
  tempy2 = ivertY2 + sY '// to screen center
  tempz2 = ivertZ2 + sZ '// to screen center
    
      
  PERS = s / (s + tempz) '//Get perspective factor
  PERS2 = s / (s + tempz2) '//Get perspective factor

  View_Plot_X = cx + tempx * PERS
  View_Plot_Y = cy - tempy * PERS
  View_Plot_X2 = cx + tempx2 * PERS2
  View_Plot_Y2 = cy - tempy2 * PERS2
  
  nDrawLine View_Plot_X, View_Plot_X2, View_Plot_Y, View_Plot_Y2, Color, 1
End Sub

Public Function InitNet(Width As Long, Height As Long, Depth As Long) As Boolean
On Error GoTo ErrH
Dim X As Long, Y As Long, Z As Long, Last As Long
m_nnWidth = Width
m_nnHeight = Height
m_nnDepth = Depth
ReDim NeuralNet(0 To Width - 1, 0 To Height - 1, 0 To Depth - 1) As typNeuron
Last = 100
X = 0
Do
Y = 0
Do
Z = 0
Do
Last = Rand(255, Last)
NeuralNet(X, Y, Z).Value(0) = Last
Last = Rand(255, Last)
NeuralNet(X, Y, Z).Value(1) = Last
Last = Rand(255, Last)
NeuralNet(X, Y, Z).Value(2) = Last
Z = Z + 1
Loop Until Z >= Depth
Y = Y + 1
Loop Until Y >= Height
X = X + 1
Loop Until X >= Width
InitNet = True
Exit Function
ErrH:
InitNet = False
Exit Function
End Function

Public Sub StartNet()
vStopNet = False
vS(0).Value = 1
vS(1).Value = 1
vS(2).Value = 1
Dim a As Variant, b As Integer
a = Time
a = Split(a, " ")
a = Split(a(0), ":")
b = a(2)
i = 0
Do
a = Time
a = Split(a, " ")
a = Split(a(0), ":")
If b <> a(2) Then
Label1.Caption = i & " FPS"
i = 0
b = a(2)
End If
i = i + 1
DoEvents
DrawNet
DoEvents
Loop Until vStopNet = True
End Sub

Public Sub StopNet()
vStopNet = True
End Sub

Public Sub DrawSides()
Draw3dLine 0, 0, 0, 0, 0, m_nnDepth - 1, vbWhite
Draw3dLine 0, 0, 0, 0, m_nnHeight - 1, 0, vbWhite
Draw3dLine 0, 0, 0, m_nnWidth - 1, 0, 0, vbWhite
Draw3dLine 0, 0, m_nnDepth - 1, 0, m_nnHeight - 1, m_nnDepth - 1, vbWhite
Draw3dLine 0, 0, m_nnDepth - 1, m_nnWidth - 1, 0, m_nnDepth - 1, vbWhite
Draw3dLine 0, m_nnHeight - 1, 0, 0, m_nnHeight - 1, m_nnDepth - 1, vbWhite
Draw3dLine 0, m_nnHeight - 1, m_nnDepth - 1, m_nnWidth - 1, m_nnHeight - 1, m_nnDepth - 1, vbWhite
Draw3dLine m_nnWidth - 1, m_nnHeight - 1, m_nnDepth - 1, m_nnWidth - 1, 0, m_nnDepth - 1, vbWhite
Draw3dLine m_nnWidth - 1, m_nnHeight - 1, m_nnDepth - 1, m_nnWidth - 1, m_nnHeight - 1, 0, vbWhite
Draw3dLine m_nnWidth - 1, m_nnHeight - 1, 0, 0, m_nnHeight - 1, 0, vbWhite
Draw3dLine m_nnWidth - 1, 0, 0, m_nnWidth - 1, m_nnHeight - 1, 0, vbWhite
Draw3dLine m_nnWidth - 1, 0, 0, m_nnWidth - 1, 0, m_nnDepth - 1, vbWhite
End Sub

Public Sub DrawNet()
UserControl.Cls
Dim i As Long, Last As Long, i2 As Long, Last2 As Long, uD As Boolean, nuD As Boolean
uD = NeuralNet(0, 0, 0).Hit
nuD = Not uD
i = 0
Last = 255
Do
'DrawSides
UserControl.Refresh
DoEvents
Last = Rand(m_nnWidth - 1, Last)
If NeuralNet(Last, 0, 0).Hit = uD Then
i = i + 1
'''''''''''''''''''''''''''''''''''''''''''''''''''
i2 = 0
Last2 = 255
NeuronStepY Last, Last2, i2, uD, nuD
'''''''''''''''''''''''''''''''''''''''''''''''''''
NeuralNet(Last, 0, 0).Hit = nuD
End If
Loop Until i >= m_nnWidth

UserControl.Refresh
End Sub

Private Sub NeuronStepY(Last As Long, Last2 As Long, i2 As Long, uD As Boolean, nuD As Boolean)
Dim i3 As Long, Last3 As Long
Do
Last2 = Rand(m_nnHeight - 1, Last2)
If NeuralNet(Last, Last2, 0).Hit = uD Then
i2 = i2 + 1
''''''''''''''''''''''''''''
i3 = 0
Last3 = 255
NeuronStepZ Last, Last2, Last3, i3, uD, nuD
''''''''''''''''''''''''''''
NeuralNet(Last, Last2, 0).Hit = nuD
End If
Loop Until i2 >= m_nnHeight
End Sub
Private Sub NeuronStepZ(Last As Long, Last2 As Long, Last3 As Long, i3 As Long, uD As Boolean, nuD As Boolean)
Do
Last3 = Rand(m_nnDepth - 1, Last3)
If NeuralNet(Last, Last2, Last3).Hit = uD Then
i3 = i3 + 1
NeuralNet(Last, Last2, Last3).Value(0) = NeuralNet(Last, Last2, Last3).Value(0) + GetNeuronStep(Last, Last2, Last3, 0)
NeuralNet(Last, Last2, Last3).Value(1) = NeuralNet(Last, Last2, Last3).Value(1) + GetNeuronStep(Last, Last2, Last3, 1)
NeuralNet(Last, Last2, Last3).Value(2) = NeuralNet(Last, Last2, Last3).Value(2) + GetNeuronStep(Last, Last2, Last3, 2)
If (Last = 0 Or Last2 = 0 Or Last3 = 0 Or Last = m_nnWidth - 1 Or Last2 = m_nnHeight - 1 Or Last3 = m_nnDepth - 1) _
And CombinedEdges(Last, Last2, Last3) Then
Draw3dPixel Last, Last2, Last3, RGB(NeuralNet(Last, Last2, Last3).Value(0), NeuralNet(Last, Last2, Last3).Value(1), NeuralNet(Last, Last2, Last3).Value(2))
End If
NeuralNet(Last, Last2, Last3).Hit = nuD
End If
Loop Until i3 >= m_nnDepth
End Sub

Private Function CombinedEdges(Last As Long, Last2 As Long, Last3 As Long) As Boolean
i = 0
If Last = 0 Then i = i + 1
If Last = m_nnWidth - 1 Then i = i + 1
If Last2 = 0 Then i = i + 1
If Last2 = m_nnHeight - 1 Then i = i + 1
If Last3 = 0 Then i = i + 1
If Last3 = m_nnDepth - 1 Then i = i + 1
If i > 1 Then CombinedEdges = True Else CombinedEdges = True
End Function

Private Function GetNeuronStep(X As Long, Y As Long, Z As Long, Value As Byte) As Integer
On Error Resume Next
Dim MinX, MaxX, MinY, MaxY, MinZ, MaxZ, vX, vY, vZ, cV
If X > 0 Then MinX = X - 1 Else MinX = X
If X < m_nnWidth - 1 Then MaxX = X + 1 Else MaxX = X
If Y > 0 Then MinY = Y - 1 Else MinY = Y
If Y < m_nnHeight - 1 Then MaxY = Y + 1 Else MaxY = Y
If Z > 0 Then MinZ = Z - 1 Else MinZ = Z
If Z < m_nnDepth - 1 Then MaxZ = Z + 1 Else MaxZ = Z
'MinX = X - 1
'MinY = Y - 1
'MinZ = Z - 1
'MaxX = X + 1
'MaxY = Y + 1
'MaxZ = Z + 1
cV = 0
vX = MinX
Do
vY = MinY
Do
vZ = MinZ
Do
If NeuralNet(vX, vY, vZ).Value(Value) > NeuralNet(X, Y, Z).Value(Value) Then
cV = cV + 1
ElseIf NeuralNet(vX, vY, vZ).Value(Value) < NeuralNet(X, Y, Z).Value(Value) Then
cV = cV - 1
End If
'cV = cV + GetChangeNumber(vX, vY, vZ, X, Y, Z, Value)
'cV = cV + GetChangeNumber(vX, vY, Z + 1, X, Y, Z, Value)
'cV = cV + GetChangeNumber(vX, vY, Z - 1, X, Y, Z, Value)
vZ = vZ + 1
Loop Until vZ > MaxZ
vY = vY + 1
Loop Until vY > MaxY
vX = vX + 1
Loop Until vX > MaxX
GetNeuronStep = cV
End Function

Public Function Rand(Max As Long, Optional Last) As Long
If Last < 1 Then Last = 100
If Max < 1 Then
Rand = 0
Exit Function
End If
a = Rnd(Last)
b = Mid(a, InStr(1, a, ".", vbTextCompare) + 1, Len(Str(Max)) - 1)
If b > Max Then
b = b - ((b \ Max) * Max)
End If
If b < 1 Then b = 0
Rand = b
End Function

Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
dX = X
dY = Y
LastVS(0) = vS(0).Value
LastVS(1) = vS(1).Value
LastVS(2) = vS(2).Value
End If
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
tx = LastVS(0) + (dX - X)
ty = LastVS(1) + (dY - Y)
If tx = 0 Then tx = 1
If ty = 0 Then ty = 1
If tx < -360 Then tx = -360
If ty < -360 Then ty = -360
If tx > 360 Then tx = 360
If ty > 360 Then ty = 360
vS(0).Value = -tx
vS(1).Value = ty
End If
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
End Property

Public Property Get nnWidth() As Long
    nnWidth = m_nnWidth
End Property

Public Property Let nnWidth(ByVal New_nnWidth As Long)
    m_nnWidth = New_nnWidth
    PropertyChanged "nnWidth"
End Property

Public Property Get nnHeight() As Long
    nnHeight = m_nnHeight
End Property

Public Property Let nnHeight(ByVal New_nnHeight As Long)
    m_nnHeight = New_nnHeight
    PropertyChanged "nnHeight"
End Property

Public Property Get nnDepth() As Long
    nnDepth = m_nnDepth
End Property

Public Property Let nnDepth(ByVal New_nnDepth As Long)
    m_nnDepth = New_nnDepth
    PropertyChanged "nnDepth"
End Property

Private Sub UserControl_InitProperties()
    m_nnWidth = m_def_nnWidth
    m_nnHeight = m_def_nnHeight
    m_nnDepth = m_def_nnDepth
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    m_nnWidth = PropBag.ReadProperty("nnWidth", m_def_nnWidth)
    m_nnHeight = PropBag.ReadProperty("nnHeight", m_def_nnHeight)
    m_nnDepth = PropBag.ReadProperty("nnDepth", m_def_nnDepth)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("nnWidth", m_nnWidth, m_def_nnWidth)
    Call PropBag.WriteProperty("nnHeight", m_nnHeight, m_def_nnHeight)
    Call PropBag.WriteProperty("nnDepth", m_nnDepth, m_def_nnDepth)
End Sub

