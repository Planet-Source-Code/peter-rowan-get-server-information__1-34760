VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FTP Info"
   ClientHeight    =   3195
   ClientLeft      =   3420
   ClientTop       =   3135
   ClientWidth     =   4680
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.Timer Timer2 
      Left            =   3720
      Top             =   240
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
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "192.168.0.3"
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   4455
   End
   Begin VB.Timer Timer1 
      Left            =   3120
      Top             =   120
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4200
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = True
Timer1.Interval = 3000
Winsock1.RemoteHost = Text1.Text
Winsock1.RemotePort = 21
Winsock1.Connect


End Sub

Private Sub Form_Load()
Text1.Text = Form1.Text1.Text
End Sub

Private Sub Timer1_Timer()
Winsock1.Close
Text2.Text = "Server Temporally Down, Or Port Not Open"

Alert = MsgBox("Time out" & vbNewLine & "Maybe the server isn't running this service?", vbExclamation, "Error")
Timer1.Enabled = False

End Sub

Private Sub Timer2_Timer()
Winsock1.Close

Timer2.Enabled = False
End Sub

Private Sub Winsock1_Connect()
'winsock connect

Dim Query As String
Form2.Caption = "Connected"


'Query = "GET / HTTP/1.0" & vbCrLf & vbCrLf

'Query = "242424gd" & vbCrLf & vbCrLf

Query = vbCrLf



Winsock1.SendData Query


End Sub




Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'winsock arrival


On Error Resume Next
Dim WebData As String
Winsock1.GetData WebData, vbString
Text2.Text = Text2.Text + WebData
Timer1.Enabled = False
Timer2.Enabled = True
Timer2.Interval = 1500
End Sub

