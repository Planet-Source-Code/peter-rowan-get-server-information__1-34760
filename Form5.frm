VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "All Info"
   ClientHeight    =   3195
   ClientLeft      =   5235
   ClientTop       =   4965
   ClientWidth     =   4680
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.Timer Timer4 
      Left            =   3480
      Top             =   360
   End
   Begin VB.Timer Timer3 
      Left            =   360
      Top             =   2760
   End
   Begin VB.Timer Timer2 
      Left            =   240
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   1080
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2400
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1680
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Send"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "192.168.0.3"
      Top             =   120
      Width           =   2535
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4200
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "SendMail"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ssh"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FTP"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   960
      Width           =   375
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text2.Text = "" Then
'Winsock1.Close
Call FTP
Else
'Winsock1.Close

Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Call FTP

End If


End Sub

Private Sub FTP()
If Winsock1.State = sckConnected Then


GoTo start
Else
Winsock1.Close
GoTo start
End If


start:
Timer1.Interval = 3000
Winsock1.RemoteHost = Text1.Text
Winsock1.RemotePort = 21
Winsock1.Connect
End Sub

Private Sub SSH()
Timer2.Interval = 3000
Winsock1.Close
Winsock1.RemoteHost = Text1.Text
Winsock1.RemotePort = 22
Winsock1.Connect
End Sub

Private Sub MAIL()
Timer3.Interval = 3000
Winsock1.Close
Winsock1.RemoteHost = Text1.Text
Winsock1.RemotePort = 25
Winsock1.Connect

End Sub

Private Sub Form_Load()
Text1.Text = Form1.Text1.Text

End Sub

Private Sub Form_Terminate()
Unload Me


End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me


End Sub

Private Sub Text2_Change()
Call SSH
End Sub

Private Sub Text3_Change()
Call MAIL

End Sub

Private Sub Timer1_Timer()
Winsock1.Close
Text2.Text = "Server Temporally Down, Or Port Not Open"

Alert = MsgBox("Time out" & vbNewLine & "Maybe the server isn't running this service?", vbExclamation, "Error")
Timer1.Enabled = False
Call SSH



End Sub

Private Sub Timer2_Timer()
Winsock1.Close
Text3.Text = "Server Temporally Down, Or Port Not Open"
Alert = MsgBox("Time out" & vbNewLine & "Maybe the server isn't running this service?", vbExclamation, "Error")

Timer2.Enabled = False
Call MAIL

End Sub

Private Sub Timer3_Timer()
Winsock1.Close
Text4.Text = "Server Temporally Down, Or Port Not Open"
Alert = MsgBox("Time out" & vbNewLine & "Maybe the server isn't running this service?", vbExclamation, "Error")
Timer3.Enabled = False
Call MAIL
End Sub

Private Sub Timer4_Timer()
'Winsock1.Close
'Winsock1.Close
Timer4.Enabled = False

End Sub

Private Sub Winsock1_Connect()
'winsock connect

Dim Query As String
Form2.Caption = "Connected"


'Query = "GET / HTTP/1.0" & vbCrLf & vbCrLf

'Query = "242424gd" & vbCrLf & vbCrLf

Query = "" & vbCrLf





Winsock1.SendData Query


End Sub




Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'winsock arrival


On Error Resume Next
Dim WebData As String
Winsock1.GetData WebData, vbString


If Text2.Text = "" Then
Text2.Text = Text2.Text + WebData
Timer1.Enabled = False
Else

If Text3.Text = "" Then
Text3.Text = Text3.Text + WebData
Timer2.Enabled = False
Else

If Text4.Text = "" Then
Text4.Text = Text4.Text + WebData
Timer3.Enabled = False
Timer4.Interval = 1500



End If
End If
End If
End Sub



