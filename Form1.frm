VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "smsSender"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox comModem 
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Example of how to send 5 times"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Sent it !!"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtNumber 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Text            =   "07971806112"
      Top             =   840
      Width           =   4215
   End
   Begin VB.TextBox txtMessage 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Text            =   "This is a test message to your mobile"
      Top             =   1320
      Width           =   5295
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5880
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      NullDiscard     =   -1  'True
      RThreshold      =   1
      RTSEnable       =   -1  'True
      SThreshold      =   1
      InputMode       =   1
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Modem Comm Port:"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "SMS Message"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Enter Mobile Number:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSend_Click()
Dim cTap As cTapProtocol
Set cTap = New cTapProtocol

'1. Setup the class with the commport of your modem and the mscomm object
cTap.Init comModem.ListIndex + 1, Me.MSComm1

'2. Connect to a valid TAP centre im using vodavone UK bu you can use any !
If Not cTap.Connected Then
  'cTap.Connect "07785 499993"
  cTap.Connect "9,07860 980480"
End If

'3. Logon to the tap centre if connected
If cTap.Connected Then
  cTap.Logon
End If

If cTap.Connected And cTap.LoggedOn Then
  cTap.SendMessage Me.txtNumber, Me.txtMessage
  If cTap.LastStatus = SENTOK Then
    MsgBox "Message Sent OK"
  Else
    MsgBox "Message Failed"
  End If
End If

If cTap.Connected Then
  cTap.Disconnect
Else
  MsgBox "No response from modem"
End If
  



Set cTap = Nothing
End Sub


Private Sub Command1_Click()
Dim cTap As cTapProtocol
Set cTap = New cTapProtocol

cTap.Init comModem.ListIndex + 1, Me.MSComm1
For x = 1 To 15
  If Not cTap.Connected Then
    'cTap.Connect "07785 499993"
    cTap.Connect "9,07860 980480"
  End If
  
  If cTap.Connected Then
    If Not cTap.LoggedOn Then
      cTap.Logon
    End If
    If cTap.LoggedOn And cTap.OkToSend Then
      cTap.SendMessage txtNumber, txtMessage
      If cTap.LastStatus = QUOITAFULL Then
        cTap.Disconnect
      End If
    End If
  End If
Next x

If cTap.Connected Then
  cTap.Disconnect
End If

Set cTap = Nothing

End Sub

Private Sub Form_Load()
comModem.Clear
comModem.AddItem "Com 1"
comModem.AddItem "Com 2"
comModem.AddItem "Com 3"
comModem.AddItem "Com 4"
comModem.ListIndex = 0
End Sub

