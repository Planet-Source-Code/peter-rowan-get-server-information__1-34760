VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get Info"
   ClientHeight    =   7620
   ClientLeft      =   1860
   ClientTop       =   825
   ClientWidth     =   7305
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   7305
   Begin VB.CommandButton Command7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "EXIT"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7200
      Width           =   3015
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Convert code to HTML"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7200
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Get all info"
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Get SendMail info"
      Height          =   495
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Get FTP info"
      Height          =   495
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Get Ssh Info"
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Text            =   "Get HTML Code and Server Info"
      Top             =   840
      Width           =   3135
   End
   Begin VB.Timer Timer1 
      Left            =   3480
      Top             =   120
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   5775
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1320
      Width           =   7095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "192.168.0.3"
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Query"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3600
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub Command1_Click()
On Error Resume Next

'connect button

If Text2.Text = "" Then

Winsock1.RemoteHost = Text1.Text
Winsock1.RemotePort = 80
Winsock1.Connect

Timer1.Interval = 10000

Else

Text2.Text = ""

Winsock1.RemoteHost = Text1.Text
Winsock1.RemotePort = 80
Winsock1.Connect

Timer1.Interval = 3000
End If



End Sub



Private Sub Command2_Click()

'test button

Form2.Show

End Sub

Private Sub Command3_Click()
Form3.Show

End Sub

Private Sub Command4_Click()
Form4.Show

End Sub

Private Sub Command5_Click()
Form5.Show

End Sub

Private Sub Command6_Click()

Open App.Path & "\" & "index.html" For Output As #1
Print #1, Text2.Text
Close #1

msg = MsgBox(App.Path & "\" & "index.html" & vbNewLine & "was created", vbInformation, "Get Info")



End Sub

Private Sub Command7_Click()
sure = MsgBox("Are you sure??", vbYesNo, "Get Info")

Select Case sure
Case 6
    End
    
Case 7
    
End Select

End Sub

Private Sub Form_Load()
'form load

Combo1.AddItem "Get HTML Code and Server Info"
Combo1.AddItem "Get More Info"

End Sub

Private Sub Timer1_Timer()
'timer stuff

Winsock1.Close
Form1.Caption = "Dissconected"
End Sub

Private Sub Winsock1_Close()
'winsock close

Form1.Caption = "Dissconected"
End Sub

Private Sub Winsock1_Connect()
'winsock connect

Dim Query As String
Form1.Caption = "Connected"

If Combo1.Text = "Get HTML Code and Server Info" Then
Query = "GET / HTTP/1.0" & vbCrLf & vbCrLf

ElseIf Combo1.Text = "Get More Info" Then
Query = "242424gd" & vbCrLf & vbCrLf
End If


Winsock1.SendData Query


End Sub




Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'winsock arrival


On Error Resume Next
Dim WebData As String
Winsock1.GetData WebData, vbString
Text2.Text = Text2.Text + WebData

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error GoTo Err

Err:

MsgBox "Host Not Found"

End Sub


