Attribute VB_Name = "NetZero312"
'NetZero 3.1.2 DUN Creator 2.2
'Created by loon
'http://www.electronerdz.com/loon

Function Pass(password)
    Key1 = "`-=~!@#$%^&*()_+[]\{}|;':" _
    & """" & ",./<>?abcdefghijklmnopqrstuvwxyzABCDEFG" _
    & "HIJKLMNOPQRSTUVWXYZ0123456789"
    Key2 = "GFEDCBAzyxwvutsrqponmlkjihgfed" _
    & "cba?></.," & """" & ":';|}{\][+_)(*&^%$#@" _
    & "!~=-`9876543210ZYXWVUTSRQPONMLKJIH"
    


    For i = 1 To Len(password)
        
        A = Mid(password, i, 1)
        
        B = InStr(1, Key1, A)
        
        C = B - (i - 1)
        
        D = Mid(Key2, C, 1)
        
        E = E + D
        
    Next i
    


    If E = "" Then
    Else
        Pass = "0" & E & "1"
    End If
End Function


Function User(Username)
User = "0:3.1.2:" & Username & "@netzero.net"
End Function


