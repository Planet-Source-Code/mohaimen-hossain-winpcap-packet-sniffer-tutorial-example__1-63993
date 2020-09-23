Attribute VB_Name = "stringGrab"
'this function will return the string between bytes "31 31 37 C0 80" and "C0 80"
Function getChat(cbuff() As Byte) As String

Dim str1 As String
Dim i, packetLength, found As Integer
       
       packetLength = UBound(cbuff)
       i = 54
       Do
            If i + 4 > packetLength Then
                getChat = ""
                Exit Function
            End If
            If cbuff(i) = 49 And cbuff(i + 1) = 49 And cbuff(i + 2) = 55 And cbuff(i + 3) = 192 And cbuff(i + 4) = 128 Then
                found = 1
                i = i + 5
                Exit Do
            End If
            If i + 1 < packetLength Then i = i + 1 Else Exit Do
        Loop
        
        If found = 1 Then
            found = 0
            Do
                If cbuff(i) = 192 And cbuff(i + 1) = 128 Then
                    found = 1
                    i = i + 2
                    Exit Do
                Else:
                    str1 = str1 & Chr(cbuff(i))
                End If
                If i + 1 < packetLength Then i = i + 1 Else Exit Do
            Loop
            getChat = str1
            Exit Function
        End If
        
        getChat = ""
                
End Function

'this function will return the string between bytes "31 30 39 C0 80" and "C0 80"
Function getNick(cbuff() As Byte) As String

Dim str1 As String
Dim i, packetLength, found As Integer
        
       packetLength = UBound(cbuff)
       
       i = 54
       Do
            If i + 4 > packetLength Then
                getNick = ""
                Exit Function
            End If
            If cbuff(i) = 49 And cbuff(i + 1) = 48 And cbuff(i + 2) = 57 And cbuff(i + 3) = 192 And cbuff(i + 4) = 128 Then
                found = 1
                i = i + 5
                Exit Do
            End If
            If i + 1 < packetLength Then i = i + 1 Else Exit Do
        Loop
        
        If found = 1 Then
            found = 0
            Do
                If cbuff(i) = 192 And cbuff(i + 1) = 128 Then
                    found = 1
                    i = i + 2
                    Exit Do
                Else:
                    str1 = str1 & Chr(cbuff(i))
                End If
                If i + 1 < packetLength Then i = i + 1 Else Exit Do
            Loop
            getNick = str1
            Exit Function
        End If
        
        getNick = ""
                
End Function

