VERSION 5.00
Begin VB.UserControl CoolBtn 
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   945
   ScaleHeight     =   495
   ScaleWidth      =   945
   Begin VB.CommandButton TabStop 
      Caption         =   "Hidden"
      Height          =   420
      Left            =   -10000
      TabIndex        =   1
      Top             =   0
      Width           =   825
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      BorderStyle     =   3  'Dot
      Height          =   195
      Left            =   45
      Top             =   45
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Btn"
      Height          =   195
      Left            =   165
      TabIndex        =   0
      Top             =   45
      Width           =   240
   End
End
Attribute VB_Name = "CoolBtn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'Sub Gradient(gFrom As Integer, gTo As Integer, Optional rHue As Integer, Optional gHue As Integer, Optional bHue As Integer)
'e = gFrom
'd = e - gTo
'inc = d / ScaleY(UserControl.Height, vbTwips, vbPixels)
'For i = 0 To ScaleY(UserControl.Height, vbTwips, vbPixels)
'  e = e - inc
'  UserControl.Line (0, ScaleY(i, vbPixels, vbTwips))-(UserControl.Width, ScaleY(i, vbPixels, vbTwips)), RGB(CInt(e) + rHue, CInt(e) + gHue, CInt(e) + bHue)
'Next i
'End Sub

Dim Up As Boolean

'Default Property Values:
Const m_def_GradientFrom = 230
Const m_def_GradientTo = 190
Const m_def_RedHue = 0
Const m_def_GreenHue = 0
Const m_def_BlueHue = 0
Const m_def_Caption = "Button"
'Property Variables:
Dim m_GradientFrom As Integer
Dim m_GradientTo As Integer
Dim m_RedHue As Variant
Dim m_GreenHue As Integer
Dim m_BlueHue As Integer
Dim m_Caption As Variant
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."




Sub CenterCaption()
w1 = Label1.Width
h1 = Label1.Height
w2 = UserControl.Width
h2 = UserControl.Height

Label1.Left = (w2 / 2) - (w1 / 2)
Label1.Top = (h2 / 2) - (h1 / 2)
Label1.Caption = m_Caption
End Sub


Sub Down3D()
pixWidth = ScaleX(UserControl.Width, vbTwips, vbPixels)
pixHeight = ScaleY(UserControl.Height, vbTwips, vbPixels)
UserControl.Line (0, 0)-(ScaleX(pixWidth - 1, vbPixels, vbTwips), ScaleY(pixHeight - 1, vbPixels, vbTwips)), RGB(76 + m_RedHue, 76 + m_GreenHue, 76 + m_BlueHue), B

UserControl.Line (ScaleX(2, vbPixels, vbTwips), ScaleY(pixHeight - 2, vbPixels, vbTwips))-(ScaleX(pixWidth - 1, vbPixels, vbTwips), ScaleY(pixHeight - 2, vbPixels, vbTwips)), vbWhite
UserControl.Line (ScaleX(pixWidth - 2, vbPixels, vbTwips), ScaleY(pixHeight - 2, vbPixels, vbTwips))-(ScaleX(pixWidth - 2, vbPixels, vbTwips), ScaleY(2, vbPixels, vbTwips)), vbWhite

UserControl.Line (ScaleX(1, vbPixels, vbTwips), ScaleY(2, vbPixels, vbTwips))-(ScaleX(1, vbPixels, vbTwips), ScaleY(pixHeight - 2, vbPixels, vbTwips)), RGB(173 + m_RedHue, 173 + m_GreenHue, 173 + m_BlueHue)
UserControl.Line (ScaleX(1, vbPixels, vbTwips), ScaleY(1, vbPixels, vbTwips))-(ScaleX(pixWidth - 2, vbPixels, vbTwips), ScaleY(1, vbPixels, vbTwips)), RGB(173 + m_RedHue, 173 + m_GreenHue, 173 + m_BlueHue)

End Sub

Sub Draw3D()
  If Up Then
    Up3D
  Else
    Down3D
  End If
End Sub

Sub Up3D()
pixWidth = ScaleX(UserControl.Width, vbTwips, vbPixels)
pixHeight = ScaleY(UserControl.Height, vbTwips, vbPixels)
UserControl.Line (0, 0)-(ScaleX(pixWidth - 1, vbPixels, vbTwips), ScaleY(pixHeight - 1, vbPixels, vbTwips)), RGB(76 + m_RedHue, 76 + m_GreenHue, 76 + m_BlueHue), B
UserControl.Line (ScaleX(2, vbPixels, vbTwips), ScaleY(pixHeight - 2, vbPixels, vbTwips))-(ScaleX(pixWidth - 1, vbPixels, vbTwips), ScaleY(pixHeight - 2, vbPixels, vbTwips)), RGB(173 + m_RedHue, 173 + m_GreenHue, 173 + m_BlueHue)
UserControl.Line (ScaleX(pixWidth - 2, vbPixels, vbTwips), ScaleY(pixHeight - 2, vbPixels, vbTwips))-(ScaleX(pixWidth - 2, vbPixels, vbTwips), ScaleY(2, vbPixels, vbTwips)), RGB(173 + m_RedHue, 173 + m_GreenHue, 173 + m_BlueHue)
UserControl.Line (ScaleX(1, vbPixels, vbTwips), ScaleY(2, vbPixels, vbTwips))-(ScaleX(1, vbPixels, vbTwips), ScaleY(pixHeight - 2, vbPixels, vbTwips)), vbWhite
UserControl.Line (ScaleX(1, vbPixels, vbTwips), ScaleY(1, vbPixels, vbTwips))-(ScaleX(pixWidth - 2, vbPixels, vbTwips), ScaleY(1, vbPixels, vbTwips)), vbWhite

End Sub


Private Sub Label1_Click()
  RaiseEvent Click
End Sub

Private Sub Label1_DblClick()
  RaiseEvent DblClick
End Sub


Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Up = False
  Draw3D
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Up = True
  Draw3D
End Sub


Private Sub TabStop_Click()
  RaiseEvent Click
End Sub

Private Sub TabStop_GotFocus()
  Shape1.Visible = True
  UserControl_MouseDown True, False, 0, 0
End Sub


Private Sub TabStop_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = Asc(" ") Then
  Up = False
  Draw3D
End If
End Sub


Private Sub TabStop_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Up3D
    RaiseEvent Click
    Down3D
  End If
    
End Sub

Private Sub TabStop_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = Asc(" ") Then
  Up = True
  Draw3D
End If
End Sub


Private Sub TabStop_LostFocus()
  Shape1.Visible = False
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Up = False
  Draw3D
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Up = True
  Draw3D
End Sub


Private Sub UserControl_Paint()
Gradient m_GradientFrom, m_GradientTo, m_RedHue, m_GreenHue, m_BlueHue
Draw3D

End Sub

Private Function Gradient(gFrom As Integer, gTo As Integer, Optional rHue, Optional gHue, Optional bHue) As Variant
e = gFrom
d = e - gTo
inc = d / ScaleY(UserControl.Height, vbTwips, vbPixels)
For i = 0 To ScaleY(UserControl.Height, vbTwips, vbPixels)
  e = e - inc
  UserControl.Line (0, ScaleY(i, vbPixels, vbTwips))-(UserControl.Width, ScaleY(i, vbPixels, vbTwips)), RGB(CInt(e) + rHue, CInt(e) + gHue, CInt(e) + bHue)
Next i
End Function


Public Property Get GradientFrom() As Integer
  GradientFrom = m_GradientFrom
End Property

Public Property Let GradientFrom(ByVal New_GradientFrom As Integer)
  m_GradientFrom = New_GradientFrom
  PropertyChanged "GradientFrom"
End Property

Public Property Get GradientTo() As Integer
  GradientTo = m_GradientTo
End Property

Public Property Let GradientTo(ByVal New_GradientTo As Integer)
  m_GradientTo = New_GradientTo
  PropertyChanged "GradientTo"
End Property

Public Property Get RedHue() As Variant
  RedHue = m_RedHue
End Property

Public Property Let RedHue(ByVal New_RedHue As Variant)
  m_RedHue = New_RedHue
  PropertyChanged "RedHue"
End Property

Public Property Get GreenHue() As Integer
  GreenHue = m_GreenHue
End Property

Public Property Let GreenHue(ByVal New_GreenHue As Integer)
  m_GreenHue = New_GreenHue
  PropertyChanged "GreenHue"
End Property

Public Property Get BlueHue() As Integer
  BlueHue = m_BlueHue
End Property

Public Property Let BlueHue(ByVal New_BlueHue As Integer)
  m_BlueHue = New_BlueHue
  PropertyChanged "BlueHue"
End Property

Private Sub UserControl_Click()
  RaiseEvent Click
  Up = True
End Sub

Private Sub UserControl_DblClick()
  RaiseEvent DblClick
End Sub

Public Property Get Caption() As Variant
Attribute Caption.VB_Description = "Sets the caption of the button."
  Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As Variant)
  m_Caption = New_Caption
  PropertyChanged "Caption"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  m_GradientFrom = m_def_GradientFrom
  m_GradientTo = m_def_GradientTo
  m_RedHue = m_def_RedHue
  m_GreenHue = m_def_GreenHue
  m_BlueHue = m_def_BlueHue
  m_Caption = m_def_Caption
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  m_GradientFrom = PropBag.ReadProperty("GradientFrom", m_def_GradientFrom)
  m_GradientTo = PropBag.ReadProperty("GradientTo", m_def_GradientTo)
  m_RedHue = PropBag.ReadProperty("RedHue", m_def_RedHue)
  m_GreenHue = PropBag.ReadProperty("GreenHue", m_def_GreenHue)
  m_BlueHue = PropBag.ReadProperty("BlueHue", m_def_BlueHue)
  m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
  
End Sub

Private Sub UserControl_Resize()
Gradient m_GradientFrom, m_GradientTo, m_RedHue, m_GreenHue, m_BlueHue
Draw3D
CenterCaption
Shape1.Top = ScaleY(2, vbPixels, vbTwips)
Shape1.Left = ScaleX(2, vbPixels, vbTwips)
Shape1.Height = UserControl.Height - ScaleY(4, vbPixels, vbTwips)
Shape1.Width = UserControl.Width - ScaleX(4, vbPixels, vbTwips)
End Sub

Private Sub UserControl_Show()
CenterCaption
Up = True
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Call PropBag.WriteProperty("GradientFrom", m_GradientFrom, m_def_GradientFrom)
  Call PropBag.WriteProperty("GradientTo", m_GradientTo, m_def_GradientTo)
  Call PropBag.WriteProperty("RedHue", m_RedHue, m_def_RedHue)
  Call PropBag.WriteProperty("GreenHue", m_GreenHue, m_def_GreenHue)
  Call PropBag.WriteProperty("BlueHue", m_BlueHue, m_def_BlueHue)
  Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
  Label1.Caption = m_Caption
  CenterCaption
  UserControl_Paint
  
End Sub

