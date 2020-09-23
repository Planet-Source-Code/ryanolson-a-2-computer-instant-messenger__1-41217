VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form serverForm 
   Caption         =   "Form2"
   ClientHeight    =   2910
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   2910
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   610
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   4680
      TabIndex        =   3
      Top             =   2295
      Width           =   4680
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   0
         Width           =   3615
      End
      Begin Project1.CoolBtn sendBt 
         Height          =   615
         Left            =   3600
         TabIndex        =   1
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         Caption         =   "Send"
      End
   End
   Begin RichTextLib.RichTextBox rt 
      Height          =   1455
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2566
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"serverForm.frx":0000
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   240
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnu1 
      Caption         =   "All the Stuff"
      Begin VB.Menu mnuRmvLn 
         Caption         =   "Remove All Extra Line Breaks"
      End
      Begin VB.Menu mnuRmvCn 
         Caption         =   "Remove Connection Messages"
      End
      Begin VB.Menu mnuClr 
         Caption         =   "Clear Window"
      End
   End
   Begin VB.Menu mnu2 
      Caption         =   "Connection"
      Begin VB.Menu mnuCon 
         Caption         =   "Connect"
      End
      Begin VB.Menu mnuDis 
         Caption         =   "Disconnect"
      End
   End
End
Attribute VB_Name = "serverForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
ws.Close
ws.LocalPort = 521
ws.Listen
rt.SelColor = vbGreen
rt.SelBold = True
rt.SelText = vbCrLf & "-----Connecting"
rt.SelBold = False
rt.SelColor = vbBlack
rt.SelStart = Len(rt.text)

End Sub

Private Sub Form_Resize()
On Error Resume Next
rt.Width = Me.ScaleWidth
rt.Height = Me.ScaleHeight - Picture1.Height
End Sub

Private Sub mnuClr_Click()
rt.text = ""
End Sub

Private Sub mnuCon_Click()
Form_Load
End Sub

Private Sub mnuDis_Click()
ws_Close
End Sub

Private Sub mnuRmvCn_Click()
Dim quiter As Integer
quiter = 0
Do Until quiter = 1
DoEvents
rt.Find "-----Connected" + vbCrLf
If rt.SelText = "" Then
quiter = 1
Exit Do
End If
rt.SelText = ""
Loop
quiter = 0
Do Until quiter = 1
DoEvents
rt.Find "-----Connecting" + vbCrLf
If rt.SelText = "" Then
quiter = 1
Exit Do
End If
rt.SelText = ""
Loop
quiter = 0
Do Until quiter = 1
DoEvents
rt.Find "-----Not Connected" + vbCrLf
If rt.SelText = "" Then
quiter = 1
Exit Do
End If
rt.SelText = ""
Loop
quiter = 0
End Sub

Private Sub mnuRmvLn_Click()
Dim quiter As Integer
quiter = 0
Do Until quiter = 1
DoEvents
rt.Find vbCrLf + vbCrLf
If rt.SelText = "" Then
quiter = 1
Exit Do
End If
rt.SelText = vbCr
Loop
quiter = 0
Do Until quiter = 1
DoEvents
rt.Find ": " + vbCrLf
If rt.SelText = "" Then
quiter = 1
Exit Do
End If
rt.SelText = ": "
Loop
End Sub

Private Sub Picture1_Resize()
On Error Resume Next
sendBt.Left = Me.ScaleWidth - sendBt.Width
Text1.Width = Me.ScaleWidth - sendBt.Width

End Sub

Private Sub Picture2_resize()
Command1.Width = Me.ScaleWidth
End Sub

Private Sub sendBt_Click()
On Error Resume Next
If Text1.text = "" Then
MsgBox "no text", vbExclamation
Exit Sub
End If
ws.SendData Text1.text
rt.SelStart = Len(rt.text)
rt.SelColor = vbRed
rt.SelFontName = "arial"
rt.SelText = vbCrLf & "Ryan: "
rt.SelFontName = "verdana"
rt.SelColor = vbBlack
rt.SelText = Text1.text
Text1.text = ""
Text1.SetFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
sendBt_Click
Text1.text = ""
Text1.SelStart = 0
End If
If Text1.text = vbCrLf Then Text1.text = ""

End Sub

Private Sub ws_Close()
ws.Close
rt.SelColor = vbGreen
rt.SelBold = True
rt.SelText = vbCrLf & "-----Not Connected"
rt.SelBold = False
rt.SelColor = vbBlack
rt.SelStart = Len(rt.text)
Form_Load
End Sub

Private Sub ws_ConnectionRequest(ByVal requestID As Long)
If ws.State <> sckClosed Then ws.Close
ws.Accept requestID
rt.SelColor = vbGreen
rt.SelBold = True
rt.SelText = vbCrLf & "-----Connected"
rt.SelBold = False
rt.SelColor = vbBlack
rt.SelStart = Len(rt.text)
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
Dim text As String

ws.GetData text

rt.SelStart = Len(rt.text)
rt.SelColor = vbBlue
rt.SelFontName = "arial"
rt.SelText = vbCrLf & "Aaron: "
rt.SelFontName = "courier new"
rt.SelColor = vbBlack
rt.SelText = text

End Sub
