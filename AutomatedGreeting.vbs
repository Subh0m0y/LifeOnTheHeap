Set Sapi = Wscript.CreateObject("SAPI.SpVoice")  
  
Dim strWelcome, strTime, strHour, strMin, strAMPM  
  
If Hour(Time) < 12 Then  
    strWelcome = "Welcome. Good Morning. "  
Else  
    If Hour(Time) > 12 Then  
        If Hour(Time) > 16 Then  
            strWelcome = "Welcome. Good evening. "  
        Else  
            strWelcome = "Welcome. Good afternoon. "  
        End If  
    End If  
End If  
  
strTime = " The current time is "  
If Hour(Time) > 12 Then  
    strHour = Hour(Time) - 12  
    strHour = strHour & " "  
Else  
    If Hour(Time) = 0 Then  
        strHour = "12 "  
    Else  
        strHour = Hour(Time) & " "  
    End If  
End If  
  
  
If Minute(Time) < 10 Then  
    strMin = " O "  
    If Minute(Time) < 1 Then  
        strMin = strMin & " clock"  
    Else  
        strMin = strMin & Minute(Time) & " "  
    End If  
Else  
    strMin = Minute(Time) & " "  
End If  
  
If Hour(Time) > 12 Then  
    strAMPM = " P M"  
Else  
    If Hour(Time) = 0 Then  
        If Minute(Time) = 0 Then  
            strAMPM = " Midnight"  
        Else  
            strAMPM = " A M"  
        End If  
    Else  
        If Hour(Time) = 12 Then  
            If Minute(Time) = 0 Then  
                strAMPM = " Noon"  
            Else  
                strAMPM = " P M"  
            End If  
        Else  
            strAMPM = " A M"  
        End If  
    End If  
End If  
      
Sapi.speak strWelcome & strTime & strHour & strMin & strAMPM
