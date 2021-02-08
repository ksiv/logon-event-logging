 On Error Resume Next
 Set objArgs = WScript.Arguments
 strFullUser = objArgs(0)
 strCode = objArgs(1)
 strFullTime = objArgs(2)

 isitComp = InStr(1, strFullUser, "$")
 strUser = Replace(strFullUser, "DOMAINNAME\", "")
 strMail = strUser & "@domainname.ru"

 If isitComp = 0 Then
   If Left(strFullUser, 2) <> "NT" Then
     Dim rst                ' Объект ADODB.Recordset
     Dim strSQLCommand      ' Query string
     Dim strErr            ' Errors
     Dim objConnect        ' Объект ADODB.Connection
     Dim strCode, strDate, strTime, strUser, strFullTime

     Const adExecuteNoRecords = "&H80"
     Const adUseClient = 3
     Const adOpenStatic = 3
     Const adLockOptimistic = 3

     strErr = 0
     Set objConnect = CreateObject("ADODB.Connection")
     Set rst = CreateObject("ADODB.Recordset")
     strDate = Mid(strFullTime, 7, 2) & "." & Mid(strFullTime, 5, 2) & "." & Left(strFullTime, 4)
     strTime = Mid(strFullTime, 9, 2) & ":" & Mid(strFullTime, 11, 2) & ":" & Mid(strFullTime, 13, 2)
     strSQLinsert = "INSERT INTO history (action,date,time,user) VALUES ('" & strCode & "','" & strDate & "','" & strTime & "','" & strUser & "') "
     strSQLselect = "SELECT * FROM history WHERE user='" & strUser & "' AND date='" & strDate & "'"
     objConnect.Open "DSN=EventLog;"

     If Err.Number <> 0 Then
       strErr = strErr & Err.Description
     End If
     Err.Clear

     rst.Open strSQLselect, objConnect, adOpenStatic, adLockOptimistic
     If Err.Number <> 0 Then
       strErr = strErr & Err.Description
     End If
     Err.Clear

     rst.MoveFirst
     rc = Err.Number 
     If Err.Number <> 0 Then
       strErr = strErr & Err.Description
     End If
     Err.Clear

     rc = 0
     Do Until rst.EOF
       rc = rc + 1
       rst.MoveNext
     Loop


     If rc = 0 Then ' There are not any records for current user
     
         objConnect.Execute strSQLinsert, , adExecuteNoRecords
         If Err.Number <> 0 Then
             strErr = strErr & Err.Description
         End If
         Err.Clear

         rst.Close
         objConnect.Close
         Set rst = Nothing
         Set objConnect = Nothing
    
                                                                                             
         Dim iMsg
         Dim iConf
         Dim Flds
         Dim strHTML
         Dim strSmartHost
    
         Const cdoSendUsingPort = 2
         strSmartHost = "SERVER"
                                                                                             
         Set iMsg = CreateObject("CDO.Message")
         Set iConf = CreateObject("CDO.Configuration")
                                                                                             
         Set Flds = iConf.Fields
                                                                                             
     ' set the CDOSYS configuration fields to use port 25 on the SMTP server
                                                                                             
                                                                                             
         With Flds
             .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
             .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "postmaster"
             .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "vdpost*vdpost"
             .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort
             .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSmartHost
             .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
             .Update
         End With
         strHTML = "text of the day"
                                                                                             
                                                                                             
     ' apply the settings to the message
         With iMsg
             Set .Configuration = iConf
                 .To = strMail
                 .From = "postmaster@domainname.ru"
                 .Subject = "Time Logged in--" & strDate & "--" & strTime & "--" & strUser
                 .HTMLBody = strHTML
                 .Send
         End With
                                                                                             
         If Err.Number <> 0 Then
             strErr = strErr & Err.Description
         End If
         Err.Clear
     
     ' cleanup of variables
         Set iMsg = Nothing
         Set iConf = Nothing
         Set Flds = Nothing
                                                                                            
    
                                                                                                                                    
                                                                                                    
         If Err.Number <> 0 Then
             Dim objFS, objFile
             Set objFS = CreateObject("Scripting.FileSystemObject")
             Set objFile = objFS.OpenTextFile("C:\LOGON.log", 8, True)
             objFile.WriteLine " rc= " & rc & strErr & strSQLinsert & ""
             objFile.Close
         End If
     End If


   End If
 End If

                


