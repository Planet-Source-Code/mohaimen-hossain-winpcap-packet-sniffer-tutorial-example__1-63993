Attribute VB_Name = "wCount"
Public anyNum As Integer
Public forNum As Integer
Public fromNum As Integer
Public whoNum As Integer
Public whyNum As Integer

'  a     n      y   = 1003
' 1001  1002   1003
'
'  f     o      r   = 2003
' 2001  2002   2003
'
'  f     r      o      m   = 3004
' 2001  3002   3003   3004
'
'  w     h      0   = 4003
' 4001  4002   4003
'
'  w     h      y   = 5003
' 4001  4002   5003

'this method will count the number of occurence of words "any","for","from","who","why" in the string str1
Sub wordCount(ByVal str1 As String)

Dim i, strLen, Counter As Integer
Dim b As Byte

strLen = Len(str1)
Counter = 0

For i = 1 To strLen
    
    b = Asc(Mid(str1, i, 1))
    Select Case b
        Case 97: 'a
            Counter = 1001
        Case 110: 'n
            Counter = 1002
        Case 121: 'y
            Select Case Counter
                Case 1002: Counter = 1003
                Case 4002: Counter = 5003
                Case Else: Counter = 0
            End Select
        Case 102: 'f
            Counter = 2001
        Case 111: '0
            Select Case Counter
                Case 2001: Counter = 2002
                Case 3002: Counter = 3003
                Case 4002: Counter = 4003
                Case Else: Counter = 0
            End Select
        Case 114: 'r
            Select Case Counter
                Case 2002: Counter = 2003
                Case 2001: Counter = 3002
                Case Else: Counter = 0
            End Select
        Case 109: 'm
            Select Case Counter
                Case 3003: Counter = 3004
                Case Else: Counter = 0
            End Select
        Case 119: 'w
            Counter = 4001
        Case 104: 'h
            Select Case Counter
                Case 4001: Counter = 4002
                Case Else: Counter = 0
            End Select
            
        Case Else:
            Counter = 0
    End Select
    
    If (Counter = 1003) Then ' "any"
        Counter = 0
        anyNum = anyNum + 1
        Form1.lblAny = anyNum
    End If
    
    If (Counter = 2003) Then ' "for"
        Counter = 0
        forNum = forNum + 1
        Form1.lblFor = forNum
    End If
    
    If (Counter = 3004) Then ' "from"
        Counter = 0
        fromNum = fromNum + 1
        Form1.lblFrom = fromNum
    End If
    
    If (Counter = 4003) Then ' "who"
        Counter = 0
        whoNum = whoNum + 1
        Form1.lblWho = whoNum
    End If
    
    If (Counter = 5003) Then ' "why"
        Counter = 0
        whyNum = whyNum + 1
        Form1.lblWhy = whyNum
    End If

Next i


End Sub
