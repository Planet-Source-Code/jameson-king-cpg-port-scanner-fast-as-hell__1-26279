VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form PS_Test 
   Caption         =   "Port scanner FAST"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   3795
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5070
      Top             =   3960
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   420
      TabIndex        =   2
      Top             =   1140
      Width           =   2985
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      Height          =   585
      Left            =   2130
      TabIndex        =   1
      Top             =   300
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Up Sockets"
      Height          =   585
      Left            =   420
      TabIndex        =   0
      Top             =   300
      Width           =   1695
   End
   Begin MSWinsockLib.Winsock Wsk 
      Index           =   0
      Left            =   4080
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Height          =   285
      Left            =   450
      TabIndex        =   6
      Top             =   3600
      Width           =   2895
   End
   Begin VB.Label Label2 
      Height          =   285
      Left            =   450
      TabIndex        =   5
      Top             =   3930
      Width           =   2895
   End
   Begin VB.Label Label3 
      Height          =   285
      Left            =   450
      TabIndex        =   4
      Top             =   4260
      Width           =   2895
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Open Ports"
      Height          =   195
      Left            =   420
      TabIndex        =   3
      Top             =   870
      Width           =   795
   End
End
Attribute VB_Name = "PS_Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Max_Con
Public Max_Port
Public Last_Checked
Public url As String
Private Is_ready(33999) As Ready
Private Type Ready
    Value As Integer
End Type




Private Sub Command1_Click()
blah = InputBox("How Many Sockets? (500 is good)")
Max_Con = blah
For x = 1 To blah
    Load Wsk(x)
Next x
End Sub

Private Sub Command2_Click()
url = InputBox("URL to SCAN", , "localhost")
Me.Caption = url
Max_Port = InputBox("Maximum port to scan")
Call Check
End Sub

Private Sub Timer1_Timer()
Label3.Caption = Val(Label3.Caption) + 1 & "     Second/s"
End Sub

Private Sub Wsk_Connect(Index As Integer)
Label1.Caption = "Last Checked In: " & Wsk(Index).RemotePort
List1.AddItem Index, 0
Wsk(Index).Close: Is_ready(Index).Value = 0 'set socket ready
End Sub
Private Sub Wsk_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Label1.Caption = "Last Checked In: " & Wsk(Index).RemotePort
Wsk(Index).Close: Is_ready(Index).Value = 0 'set socket ready
End Sub
Sub Check()
Strt:
Timer1.Enabled = True
DoEvents
For x = 1 To Max_Con
DoEvents
'Wsk(x).Close This is Bad!!
If Is_ready(x).Value <> 0 Then GoTo 20: ' Socket Isnt done checking skip
    Wsk(x).Connect url, Last_Checked + 1
   
    Is_ready(x).Value = 1 'Set socket not ready
    Last_Checked = Last_Checked + 1
    Label2.Caption = "Next To Be Scanned: " & Last_Checked
20:
Next x
If Val(Last_Checked) >= Val(Max_Port) Then
    Timer1.Enabled = False
    For y = 1 To Max_Con
    'Clean Up
        Is_ready(y).Value = 0
        Unload Wsk(y)
    Next y
    MsgBox ("Stoped")
    Exit Sub
Else:
    GoTo Strt:
End If
End Sub


