VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Yahoo Messenger Chat Viewer"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4343
      TabIndex        =   4
      Top             =   4560
      Width           =   1215
   End
   Begin VB.ComboBox cbadapter 
      Height          =   315
      Left            =   383
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   3960
      Width           =   3375
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   383
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   3255
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   5175
   End
   Begin VB.Label lblWhy 
      Caption         =   "0"
      Height          =   255
      Left            =   6480
      TabIndex        =   15
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label lblWho 
      Caption         =   "0"
      Height          =   255
      Left            =   6480
      TabIndex        =   14
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label lblFrom 
      Caption         =   "0"
      Height          =   255
      Left            =   6480
      TabIndex        =   13
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblFor 
      Caption         =   "0"
      Height          =   255
      Left            =   6480
      TabIndex        =   12
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblAny 
      Caption         =   "0"
      Height          =   255
      Left            =   6480
      TabIndex        =   11
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "why"
      Height          =   255
      Left            =   5880
      TabIndex        =   10
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "who"
      Height          =   255
      Left            =   5880
      TabIndex        =   9
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "from"
      Height          =   255
      Left            =   5880
      TabIndex        =   8
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "for"
      Height          =   255
      Left            =   5880
      TabIndex        =   7
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "any :"
      Height          =   255
      Left            =   5880
      TabIndex        =   6
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Word Count"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'packet sniffer example program by ruleworld
'email: ruleworld@gmail.com

Dim go As Boolean

Private Sub cmdExit_Click()
Unload Me
End
End Sub

Private Sub cmdStart_Click()
vpSetCurrentAdapter cbadapter.ListIndex ' Set up the current working adapter
cmdStart.Enabled = False
cmdStop.Enabled = True
capture ' in this method we capture packet and analyze packet
End Sub

Private Sub cmdStop_Click()
go = False 'stop the loop
cmdStart.Enabled = True
cmdStop.Enabled = False
End Sub

Private Sub Form_Load()

Dim numadapters As Long
Dim i As Integer
Dim name As String
Dim desc As String
'Dim adinf As AdINFO ' Adapter info to use with vpGetAdpaterInfo

numadapters = VBPcapInit ' Start VBPCAP engine

For i = 0 To numadapters - 1
    vpGetAdapterInfoVB5 i, name, desc   'Enumerate and show adapters
    'vpGetAdapterInfo i, adinf 'Uncomment this and comment the line above
    cbadapter.AddItem desc
    'cbadapter.AddItem adinf.Description 'Same as above
Next

cbadapter.ListIndex = 0

cmdStart.Enabled = True
cmdStop.Enabled = False

'initialize word count variables
anyNum = 0
forNum = 0
fromNum = 0
whoNum = 0
whyNum = 0

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
VBPcapTerminate ' Stop VBPCAP engine
End Sub

Sub capture()

Dim cbuff(1600) As Byte ' Declaring Byte array. a packet normally won't exceed 1500 bytes. 1600 is a safe value
Dim hd As PacketHeader ' This is the header returned for each packet
Dim rval As Long ' retvalue

' the following three classes are for decoding packet header, used for filtering packets
Dim eth As New clsEthernet
Dim iph As New clsIP
Dim tcph As New clsTCP

Dim pos, n, evaluate  As Integer
Dim str1, chatText, chatNick As String


vpBegin 20   ' 20 msec timeout
go = True

Do While go
    rval = vpCapture(cbuff(), hd) 'vbpcap packet capture function
        
    If rval > 0 Then ' rval > 0 means we received a packet
            
        evaluate = 1 'if evaluate = 1 then we will analyze the packet
        eth.Decode cbuff 'decode ethernet header
        If (eth.TypeCode = &H800) Then ' IPPROTOCOL
                        
            iph.Decode cbuff 'decode ip header
            
            'if src/dest ip matches our ip "xxx.xxx.xxx.xxx" then we will evaluate
            'If (StrComp(iph.DestIP, "xxx.xxx.xxx.xxx", vbTextCompare) = 0) Then evaluate = 1
            'If (StrComp(iph.SrcIP, "xxx.xxx.xxx.xxx", vbTextCompare) = 0) Then evaluate = 1
                    
            If iph.Protocol <> 6 Then ' tcp = 6
                evaluate = 0 ' if not tcp then we will not evaluate packet
            Else
                tcph.Decode cbuff, (iph.TotalLength * 4) 'decode tcp header
                n = tcph.srcPort
                If (n < 1024) Then evaluate = 0 'src port must be 1024 or higher
                n = tcph.dstPort
                If (n < 1024) Then evaluate = 0 'dest port must be 1024 or higher
            End If
        End If
    
        ' pick the first four characters of packet data and see if matches "YMSG"
        str1 = ""
        str1 = str1 & Chr(cbuff(54))
        str1 = str1 & Chr(cbuff(55))
        str1 = str1 & Chr(cbuff(56))
        str1 = str1 & Chr(cbuff(57))
                        
        If evaluate = 1 And (StrComp(str1, "YMSG", vbTextCompare) = 0) Then
            
            chatNick = getNick(cbuff) 'get the chat sender's nick from packet
            chatText = getChat(cbuff) 'get that chat text from packet
        
            If Len(chatText) <> 0 Then
                ' the chat text contains html formatting code btn <> so we will get the text after last ">"
                pos = InStrRev(chatText, ">")
                chatText = Mid(chatText, pos + 1)
            End If
            
            If Len(chatText) <> 0 Then
                wordCount chatText ' look for some specific words in chat text
                Text1.Text = Text1.Text & chatNick & ": " & chatText & vbCrLf
                Text1.SelStart = Len(Text1.Text)
            End If
        End If
        
    End If

    DoEvents 'this method handles control to windows to update the window. useful if the loop is too busy. keep it.
Loop

vpEnd

End Sub

