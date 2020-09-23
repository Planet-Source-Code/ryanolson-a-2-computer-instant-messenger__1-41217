VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   4680
      TabIndex        =   2
      Top             =   2580
      Width           =   4680
      Begin Project2.CoolBtn sendbt 
         Height          =   600
         Left            =   3600
         TabIndex        =   1
         Top             =   10
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1058
         Caption         =   "Send"
      End
      Begin VB.TextBox Text1 
         Height          =   600
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   10
         Width           =   3615
      End
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   1080
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rt 
      Height          =   735
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0000
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p2dn As Integer



Private Sub Form_Load()
p2dn = 0
Do Until ws.State <> sckClosed
DoEvents
ws.Close
ws.Connect "server ip here", 521
Loop

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
sendbt.Left = Me.ScaleWidth - sendbt.Width
Text1.Width = Me.ScaleWidth - sendbt.Width

End Sub

Private Sub sendbt_Click()
If Text1.text = "" Then
MsgBox "no text", vbExclamation
Exit Sub
End If
ws.SendData Text1.text
rt.SelStart = Len(rt.text)
rt.SelColor = vbBlue
rt.SelFontName = "arial"
rt.SelText = vbCrLf & "Aaron: "
rt.SelFontName = "courier new"
rt.SelColor = vbBlack
rt.SelText = Text1.text
Text1.text = ""
Text1.SetFocus
End Sub




Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
sendbt_Click
Text1.text = ""
Text1.SelStart = 0
End If
If Text1.text = vbCrLf Then Text1.text = ""

End Sub




Private Sub ws_Close()
Dim msg
ws.Close
rt.SelColor = vbGreen
rt.SelBold = True
rt.SelText = vbCrLf & "-----Not Connected"
rt.SelBold = False
rt.SelColor = vbBlack
rt.SelStart = Len(rt.text)
msg = MsgBox("Do you want to try to reconnect?", vbYesNo)
If msg = vbYes Then Form_Load
End Sub

Private Sub ws_Connect()
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
rt.SelColor = vbRed
rt.SelFontName = "arial"
rt.SelText = vbCrLf & "Ryan: "
rt.SelColor = vbBlack
rt.SelFontName = "verdana"
rt.SelText = text

End Sub

