VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTapProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=============================================================================
' TAP Protocol Class for vb
' Written by Mike Ball, Staffordshire, England
' Mikeb@ibs-systems.co.uk
' If you use this class then please mention me in your credits
' Thanks and enjoy
' Please also vote for me :-)
'
' Tap centre numbers can be obtained from http://www.notepager.net/tap-phone-numbers.htm
'
'




Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)




Private WithEvents m_Comm As MSComm
Attribute m_Comm.VB_VarHelpID = -1
Private m_Buffer As String

Public Enum enumStatus
  SENTOK
  BADCHECKSUM
  MESSAGENOTSUPPORTED
  QUOITAFULL
End Enum

Private Type RECIEVE_BUFF
  strASCII() As String
  lASCII() As Integer
  lLength As Long
End Type

Private Const SOH = 1
Private Const STX = 2
Private Const ETX = 3
Private Const EOT = 4
Private Const ENQ = 5
Private Const ACK = 6
Private Const CR = 13
Private Const NAK = 21
Private Const ESC = 27
Private Const RS = 30
Private Const US = 31

Private m_bConnected As Boolean
Private m_bLoggedOn As Boolean
Private m_bOkToSend As Boolean
Private m_LastReply As enumStatus

Public Function Init(lPort As Long, objComm As MSComm) As Boolean
Dim szSetting As String
  Set m_Comm = objComm
  If m_Comm.PortOpen Then
    m_Comm.PortOpen = False
  End If
  m_Comm.CommPort = lPort
  m_Comm.Settings = "9600,N,8,1"
  m_Comm.PortOpen = True
End Function



Public Function Connect(szNumber As String) As Boolean
Dim x As Long
Dim bOK As Boolean
Dim szReply As String

  SendData "ATZ", True, False
  WaitReply 10, Chr(CR)
  SendData "ATH", True, False
  WaitReply 10, Chr(CR)
  SendData "ATDT " & szNumber, True, False
  m_Buffer = ""
  szReply = WaitReply(100, "CONNECT")  ' Wait  to connect
  bOK = False
  If Len(szReply) > 0 Then
    For x = 1 To 5
      SendData vbCr, False, False
      szReply = WaitReply(20, Chr(CR))
      If InStr(1, szReply, "ID=", vbTextCompare) Then
        bOK = True
        m_bConnected = True
        m_bLoggedOn = False
        Exit For
      End If
    Next x
  End If
  
  Connect = bOK
  
End Function

Public Sub Disconnect()
 
  SendData Chr(CR) & Chr(ESC) & Chr(EOT), True, False
  WaitReply 10, Chr(CR)
  SendData "ATH", True, False
  WaitReply 10, Chr(CR)
  SendData "ATH", True, False
  WaitReply 10, Chr(CR)
  
  m_bConnected = False
  m_bLoggedOn = False
  
End Sub

Public Function Logon() As Boolean
Dim szReply As String
Dim bOK As Boolean
Dim x As Long

bOK = False

If Not m_bConnected Then
  Debug.Assert False
  bOK = False
Else
  SendData Chr(ESC) & "PG1", True, False
  szReply = ""
  For x = 1 To 10 ' Get the [p string
    szReply = WaitReply(20, Chr(13))
    If InStr(1, szReply, "[p", vbTextCompare) Then
      bOK = True
      m_LastReply = SENTOK
      Exit For
    End If
  Next x
End If

m_bOkToSend = bOK
m_bLoggedOn = bOK
Logon = bOK

End Function

Private Sub Class_Initialize()
  m_bLoggedOn = False
  m_bConnected = False
  m_bOkToSend = False
End Sub

Private Sub Class_Terminate()
If m_Comm.PortOpen Then
  m_Comm.PortOpen = False
End If
Set m_Comm = Nothing
End Sub

Public Property Get LastStatus() As enumStatus
  LastStatus = m_LastReply
End Property

Public Property Get OkToSend() As Boolean
  OkToSend = m_bOkToSend
End Property

Public Property Get LoggedOn() As Boolean
  LoggedOn = m_bLoggedOn
End Property

Public Property Get Connected() As Boolean
  Connected = m_bConnected
End Property

Private Sub m_Comm_OnComm()
Dim Buffer As Variant
    Select Case m_Comm.CommEvent
        ' Event messages.
        Case comEvReceive
            Buffer = m_Comm.Input
            m_Buffer = m_Buffer & StrConv(Buffer, vbUnicode)
            
    End Select
End Sub


Private Function SendData(szdata As Variant, bWantCr As Boolean, bWantChecksum As Boolean) As Boolean
Dim szTemp As String
Dim B As String
Dim Check1 As String
Dim Check2 As String
Dim Check3 As String
Dim CheckSum As String
Dim j As Integer
Dim Sum As Integer
Dim D As Integer
Dim d1 As Integer
Dim d2 As Integer
Dim d3 As Integer

szTemp = szdata
If bWantCr Then
  szTemp = szTemp & vbCrLf
End If

If bWantChecksum Then
  j = 1
  Sum = 0
  Do Until j > Len(szTemp)
    B$ = Mid(szTemp, j, 1)
    D = Asc(B)
    Sum = Sum + D
    j = j + 1
   Loop
  'Create the three characters to be transmitted to represent this checksum.
  d3 = 48 + Sum - Int(Sum / 16) * 16
  Sum = Int(Sum / 16)
  d2 = 48 + Sum - Int(Sum / 16) * 16
  Sum = Int(Sum / 16)
  d1 = 48 + Sum - Int(Sum / 16) * 16
  Check1 = Chr$(d1)
  Check2 = Chr$(d2)
  Check3 = Chr$(d3)
  CheckSum = Check1 & Check2 & Check3
  szTemp = szTemp & CheckSum
End If


Debug.Print Time & " - SEND DATA: " & szTemp
  m_Buffer = "" ' reset buffer first
  m_Comm.Output = szTemp
  DoEvents
  Sleep 500
  DoEvents
End Function


Private Function GetBuffer(szPacket As String) As RECIEVE_BUFF
Dim szBuff As RECIEVE_BUFF
Dim lLen As Long
Dim x As Long
Dim szCh As String

lLen = Len(szPacket)
ReDim szBuff.lASCII(lLen)
ReDim szBuff.strASCII(lLen)

szBuff.lLength = lLen
For x = 1 To lLen
  szCh = Mid(szPacket, x, 1)
  szBuff.strASCII(x) = szCh
  szBuff.lASCII(x) = Asc(szCh)
Next x
GetBuffer = szBuff

End Function


Private Function WaitReply(lDelay As Long, WaitString As String) As Variant
Dim szTemp As Variant
Dim Buffer As Variant
Dim x As Long
Dim bOK As Boolean

DoEvents
If WaitString = "" Then
  bOK = True
  For x = 1 To lDelay
    Sleep 1000
    DoEvents
  Next x
Else
  bOK = False
  For x = 1 To lDelay
    DoEvents
    If InStr(m_Buffer, WaitString) Then
      bOK = True
      Exit For
      
    Else
      DoEvents
      Sleep 500
      DoEvents
    End If
  Next x
End If

If Len(m_Buffer) > 0 Then
  DoEvents
  Debug.Print Time & " - GOT DATA: " & m_Buffer
End If

If bOK = False Then
  Debug.Print Time & " - Timeout on waitreply for " & WaitString
End If

WaitReply = m_Buffer


End Function

Public Sub SendMessage(szNumber As String, szMessage As String)
Dim szPacket As String
Dim szReply As String
Dim bStop As Boolean
Dim lRetry As Long
Dim lReply As Long
Dim szBuff As RECIEVE_BUFF

szPacket = Chr(STX) & szNumber & Chr(CR) & szMessage & Chr(CR) & Chr(ETX)

lRetry = 1

Do While bStop = False
  SendData szPacket, False, True
  SendData "", True, False
  szReply = WaitReply(30, Chr(CR))
  szBuff = GetBuffer(szReply)
       
  '<ACK> <CR> = OK, send next block.
  '<NAK> <CR> = Checksum error, send latest block again.
  '<RS> <CR> = Abandon current transaction and go to next. RS may occur when the checksum is OK, but the current transaction violates a system rule. At the option of the system, it may occur in other cases.
  '<ESC> <EOT> <CR> = Begin disconnect.
       
  If szBuff.lLength = 0 Then
    m_bOkToSend = False
    bStop = True
    Debug.Print "NO REPLY ??"
  Else
    lReply = szBuff.lASCII(szBuff.lLength - 1)
    Select Case lReply
    Case ACK
      Debug.Print "OK ACK"
      m_bOkToSend = True
      bStop = True
      m_LastReply = SENTOK
    Case NAK
      Debug.Print "CHECKSUM ?? NAK"
      m_bOkToSend = False
      m_LastReply = BADCHECKSUM

    Case RS
      Debug.Print "SYSTEM DOES NOT ACCECPT THIS TYPE OF MESSAGE RS"
      m_bOkToSend = False
      bStop = True
      m_LastReply = MESSAGENOTSUPPORTED
      
    Case EOT
      Debug.Print "NO MORE PLEASE DISCONNECT EOT"
      m_bOkToSend = False
      bStop = True
      m_LastReply = QUOITAFULL
      
    Case Else
      m_bOkToSend = False
      bStop = True
      m_LastReply = QUOITAFULL
    End Select
  End If
  lRetry = lRetry + 1
  If lRetry > 4 Then
    bStop = True
    m_bOkToSend = False
  End If
Loop

End Sub
