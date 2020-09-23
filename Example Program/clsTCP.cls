VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Variabile o variabili locali per la memorizzazione dei valori delle proprietà.
Private mvarsrcPort As Long 'Copia locale.
Private mvardstPort As Long 'Copia locale.
Private mvarseqNumber As Long 'Copia locale.
Private mvarackNumber As Long 'Copia locale.
Private mvarHLEN As Byte
Private mvardataOffset As Byte 'Copia locale.
Private mvarReserved As Byte 'Copia locale.
Private mvarFLAGS As Byte 'Copia locale.
Private mvarwindow As Integer 'Copia locale.
Private mvarHD_Checksum As Integer 'Copia locale.
Private mvarurgentPointer As Integer 'Copia locale.
Private mvarOptPad As Long 'Copia locale.
Private mvarsFLAGS  As String
Public PROTID As Long
Private Const START_OFFSET = 34




Public Property Get OptPad() As Long
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.OptPad
    OptPad = mvarOptPad
End Property





Public Property Get urgentPointer() As Integer
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.urgentPointer
    urgentPointer = mvarurgentPointer
End Property






Public Property Get HD_Checksum() As Integer
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.HD_Checksum
    HD_Checksum = mvarHD_Checksum
End Property





Public Property Get window() As Integer
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.window
    window = mvarwindow
End Property






Public Property Get flags() As Byte
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.flags
    flags = mvarFLAGS
End Property






Public Property Get Reserved() As Byte
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.Reserved
    Reserved = mvarReserved
End Property





Public Property Get dataOffset() As Byte
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.dataOffset
    dataOffset = mvardataOffset
End Property






Public Property Get ackNumber() As Long
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.ackNumber
    ackNumber = mvarackNumber
End Property





Public Property Get seqNumber() As Long
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.seqNumber
    seqNumber = mvarseqNumber
End Property






Public Property Get dstPort() As String
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.dstPort
    dstPort = mvardstPort
End Property





Public Property Get srcPort() As String
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.srcPort
    srcPort = mvarsrcPort
End Property

Public Property Get flagString() As String

flagString = mvarsFLAGS
End Property

Public Sub Decode(bff() As Byte, start As Long)
mvarsrcPort = MergeShort(bff(START_OFFSET), bff(START_OFFSET + 1))
mvardstPort = MergeShort(bff(START_OFFSET + 2), bff(START_OFFSET + 3))
mvarseqNumber = MakeDWord(MakeWord(bff(START_OFFSET + 7), bff(START_OFFSET + 6)), MakeWord(bff(START_OFFSET + 5), bff(START_OFFSET + 4)))
mvarackNumber = MakeDWord(MakeWord(bff(START_OFFSET + 11), bff(START_OFFSET + 10)), MakeWord(bff(START_OFFSET + 9), bff(START_OFFSET + 8)))
mvarHLEN = (bff(START_OFFSET + 12) And &HF0) / 16
mvarFLAGS = bff(START_OFFSET + 13)
mvarwindow = MergeShort(bff(START_OFFSET + 14), bff(START_OFFSET + 15))
mvarHD_Cheksum = MergeShort(bff(START_OFFSET + 16), bff(START_OFFSET + 17))
mvarurgentPointer = MergeShort(bff(START_OFFSET + 18), bff(START_OFFSET + 19))
' cambiato da decodestringflags(b as byte)
DecodeStringFLAGS
End Sub

Private Sub Class_Initialize()
PROTID = 6 ' this is the ID of TCP protocol inside IP struct
End Sub
Public Property Get HLEN() As Byte
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.srcPort
    HLEN = mvarHLEN
End Property

Private Sub DecodeStringFLAGS()
mvarsFLAGS = ""
If (mvarFLAGS And &H1) Then mvarsFLAGS = "FIN;"
If (mvarFLAGS And &H2) Then mvarsFLAGS = mvarsFLAGS & "SYN;"
If (mvarFLAGS And &H4) Then mvarsFLAGS = mvarsFLAGS & "RESET;"
If (mvarFLAGS And &H8) Then mvarsFLAGS = mvarsFLAGS & "PUSH;"
If (mvarFLAGS And &H10) Then mvarsFLAGS = mvarsFLAGS & "ACK;"
If (mvarFLAGS And &H20) Then mvarsFLAGS = mvarsFLAGS & "URGENT;"

End Sub
