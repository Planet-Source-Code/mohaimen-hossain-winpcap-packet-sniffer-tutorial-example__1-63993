Attribute VB_Name = "modNet"

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer
Public Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Public Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Integer) As Integer
Public Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long



Type ip_address
 byte1 As Byte
 byte2 As Byte
 byte3 As Byte
 byte4 As Byte
End Type

Type Ethernet_header
srcmac As ip_address
destmac As ip_address
type_length As Integer
End Type

Type ip_header
ver_ihl As Byte ' Ver= (ver and &Hf0)/(2^4)  IHL=(ver and &hf)
tos As Byte
tlen As Integer
Identification As Integer
flags_fo As Integer
TTL As Byte
proto As Byte
crc As Integer
saddr As ip_address
daddr As ip_address
op_pad As Long
End Type

'typedef struct ip_header{
 '   u_char  ver_ihl;        // Version (4 bits) + Internet header length (4 bits)
  '  u_char  tos;            // Type of service
   ' u_short tlen;           // Total length
    'u_short identification; // Identification
    'u_short flags_fo;       // Flags (3 bits) + Fragment offset (13 bits)
    'u_char  ttl;            // Time to live
    'u_char  proto;          // Protocol
    'u_short crc;            // Header checksum
    'ip_address  saddr;      // Source address
    'ip_address  daddr;      // Destination address
    'u_int   op_pad;         // Option + Padding
'}ip_header;

Type udp_header
sport As Integer
dport As Integer
len As Integer
crc As Integer
End Type
Public Function MergeShort(x As Byte, y As Byte) As Integer


If x And &H80 Then
MergeShort = (((x * &H100&) Or y) Or &HFFFF0000)
Else
MergeShort = ((x * &H100) Or y)
End If

End Function

Public Function MakeWord(LoByte As Byte, HiByte As Byte) As Integer
  If HiByte And &H80 Then
    MakeWord = ((HiByte * &H100&) Or LoByte) Or &HFFFF0000
  Else
    MakeWord = (HiByte * &H100) Or LoByte
  End If
End Function

Public Function MergeShortl(x As Byte, y As Byte) As Variant


If x And &H80 Then
MergeShortl = (((x * &H100&) Or y) Or &HFFFF0000)
Else
MergeShortl = ((x * &H100) Or y)
End If

End Function
Public Function MakeDWord(LoWord As Integer, HiWord As Integer) As Long
MakeDWord = (HiWord * &H10000) Or (LoWord And &HFFFF&)
End Function

Public Function Byte2Hex(ByRef bytes() As Byte, startPos As Integer, Count As Integer) As String
Dim i As Long
Dim tempstr As String


    For i = startPos To startPos + Count
          tempstr = Hex(bytes(i))
        If Len(tempstr) = 1 Then
            Byte2Hex = Byte2Hex & "0" & tempstr
        Else
            Byte2Hex = Byte2Hex & tempstr
        End If
        If i < startPos + Count Then Byte2Hex = Byte2Hex & ":"
    Next

    'Byte2Hex = tempstr

End Function


Public Function Byte2Hex2(ByRef bytes() As Byte, startPos As Integer, Count As Integer) As String
Dim i As Integer
Dim tempstr As String

    For i = startPos To startPos + Count

        If Len(Hex(bytes(i))) = 1 Then
            tempstr = tempstr & "0" & Hex(bytes(i))
        Else
            tempstr = tempstr & Hex(bytes(i))
        End If
        
    Next i

    Byte2Hex2 = tempstr

End Function

Public Function IPfromByte(ByRef bytes() As Byte, startPos As Integer, Count As Integer) As String
Dim i As Long
Dim tempstr As String

    For i = startPos To startPos + Count

      
            tempstr = tempstr & bytes(i)

         If i < startPos + Count Then tempstr = tempstr & "."
    Next

    IPfromByte = tempstr

End Function


