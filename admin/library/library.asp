<%  
  Const SYSTEM_ID = 1
  Const SYSTEM_NAME = "Placement Record Searching"
  Const BLANK_PASSWORD = "DA39A3EE5E6B4B0D3255BFEF95601890AFD80709"

  Const mailPlain = 0
  Const mailHTML = 1 

  function openConnection(ByVal ConnectionStr)
    dim conn
    set conn = Server.CreateObject("ADODB.Connection")
    conn.ConnectionString = ConnectionStr
    conn.open
    set openConnection = conn
  end function
  
  sub closeConnection(ByRef conn)
    conn.close
    Set conn = nothing
  end sub

  Sub MakeHeader(ByVal title)
    Dim user, priv
    Set user = Session("User")
    Response.Write "<div class=""hideElement center"">" & vbNewLine
    Response.Write "<p id=""top"">" & _
                   "<img src=""img/recrdsrch.jpg"" alt=""STCC Placement Record Searching System"" width=""505"" height=""54"" />" & _
                   "</p>" & vbNewLine
    Response.Write "<h2>" & Title & "</h2>" & vbNewLine
    Response.Write "<h5>" & UCase(Session("User")("username")) & " has " 
    priv = ""
    if user("IsAdmin") then
      priv = concatenate(priv, "ADMIN")
    end if
    if user("IsWriter") then
      priv = concatenate(priv, "WRITER")
    end if
    if user("IsEssay") then
      priv = concatenate(priv, "ESSAY")
    end if
    priv = concatenate(priv, "READER")
    Response.Write "(" & priv & ") access.</h5>" & vbNewLine
    Response.Write "<hr />" & vbNewLine
    Response.Write "</div>" & vbNewLine
  end Sub

  Sub mailmessage(ByVal mailStrTo, ByVal mailStrFrom, ByVal strsubject, ByVal strbody)
    Call mailmessageEx(mailStrTo, mailStrFrom, "", "", strsubject, strbody, mailHTML)
  End Sub
  
  Sub mailmessageEx(ByVal strto, ByVal strfrom, ByVal strcc, ByVal strbcc, ByVal strsubject, ByVal strbody, ByVal mailtype)
    Dim msgFormat
    msgFormat = iif(mailtype = mailPlain, "TEXT", "HTML")
    Call SendGmail(strto, strFrom, strcc, strbcc, strsubject, strbody, msgFormat)
  End Sub
  
  Sub SendGmail(ByVal RecipientEmail, ByVal SenderEmail, ByVal strCC, ByVal strBCC, ByVal Subject, ByVal msgBody, ByVal msgFormat)
    On Error Resume Next
    Dim SMTPServer, SMTPusername, SMTPpassword
    Dim sch, cdoConfig, cdoMessage
    SMTPserver   = "smtp.gmail.com"
    SMTPusername = "placementtesting@stcc.edu"
    SMTPpassword = "Hello$World"
    sch = "http://schemas.microsoft.com/cdo/configuration/"
    Set cdoConfig = Server.CreateObject("CDO.Configuration")
    With cdoConfig.Fields
        .Item(sch & "smtpauthenticate") = 1
        .Item(sch & "smtpusessl") = True
        .Item(sch & "smtpserver") = SMTPserver
        .Item(sch & "sendusername") = SMTPusername
        .Item(sch & "sendpassword") = SMTPpassword
        .Item(sch & "smtpserverport") = 465 '587
        .Item(sch & "sendusing") = 2
        .Item(sch & "connectiontimeout") = 100
        .update
    End With
    Set cdoMessage = Server.CreateObject("CDO.Message")
    Set cdoMessage.Configuration = cdoConfig
    cdoMessage.From = SenderEmail
    cdoMessage.To = RecipientEmail
    cdoMessage.Cc = strCC
    cdoMessage.Bcc = strBCC
    cdoMessage.Subject = Subject
    If Ucase(msgFormat) = "TEXT" Then
        cdoMessage.TextBody = msgBody
    Else
        cdoMessage.HTMLBody = msgBody
    End If
    cdoMessage.Send
    Set cdoMessage = Nothing
    Set cdoConfig = Nothing
    If Err.Number <> 0 Then
        Response.Write "error: " & err.Number & " - " & err.Description & "<br /><br />"
    End If
  End Sub
 
  Function ZPadStr(ByVal X, ByVal Length)
    Dim temp
    if IsNull(X) then
      X = ""
    end if
    temp = Right(CStr(X), Length)
    ZPadStr = String(Length - Len(temp), "0") + temp
  End Function
  
  Function FixNull(ByVal str)
    if IsNull(str) or str = "" then
      FixNull = "&nbsp;"
    else
      FixNull = str
    end if  
  End Function
   
  function fixSSN(ByVal str)
    if instr(1, str, "-") = 0 then
      fixSSN = Mid(str, 1, 3) & "-" & Mid(str, 4, 2) & "-" & Mid(str, 6, 4)
    else
      fixSSN = str
    end if
  end function             
  
  function iif(ByVal cond, ByVal tpart, ByVal fpart)
    if cond then
      iif = tpart
    else
      iif = fpart
    end if
  end function     
  
  function KeepJustDigits(ByVal str)
    Dim i, ch, retstr
    retstr = ""
    for i = 1 to Len(str)
      ch = mid(str, i, 1)
      if ch >= "0" and ch <= "9" then
        retstr = retstr & ch
      end if
    next
    KeepJustDigits = retstr
  end function
  
  function PadR(ByVal str, ByVal Len)
    PadR = Left(str & Space(Len), Len)
  end function
  
  Function BuildErrMsg(ByVal ErrMsg, ByVal Msg)
    if Msg = "" then
      BuildErrMsg = ErrMsg
    else
      if ErrMsg <> "" then 
        ErrMsg = ErrMsg & "<br />"
      end if
      BuildErrMsg = ErrMsg & Msg
    end if
  End Function
  
  Function FixQuote(ByVal str)
    FixQuote = Replace(str, "'", "''")
  End Function
  
  Function appendCondition(ByVal cond, ByVal tempstr)
    if cond = "" then
      cond = cond & tempstr
    else
      cond = cond & " AND " & tempstr
    end if
    appendCondition = cond
  End Function

  Function concatenate(ByVal total, ByVal priv)
    if total = "" then
      total = total & priv
    else
      total = total & ", " & priv
    end if
    concatenate = total
  End Function

  Function XORDecryption(ByVal codeKey, ByVal DataIn)
    Dim lonDataPtr, strDataOut, intXOrValue1, intXOrValue2
    For lonDataPtr = 1 To Len(DataIn) / 2
      intXOrValue1 = CInt("&H0" & Mid(DataIn, lonDataPtr * 2 - 1, 2))
      intXOrValue2 = Asc(Mid(codeKey, ((lonDataPtr Mod Len(codeKey)) + 1), 1))
      strDataOut = strDataOut + Chr(intXOrValue1 Xor intXOrValue2)
    Next
    XORDecryption = strDataOut
  End Function

  Function FormatPlacement(ByVal str)
    str = UCase(Trim(str))
    'If str <> "ESSAY" And InStr(1, str, "-") = 0 Then
    '  str = Mid(str, 1, 4) & "-" & Mid(str, 5)
    'End If
    FormatPlacement = str
  End Function

  Function InputFilter(ByVal userInput)
	Dim newString, regEx
	userInput = Replace(userInput, "'", "''")
	Set regEx = New RegExp
	regEx.Pattern = "([^A-Za-z0-9@.' _-]+)"
	regEx.IgnoreCase = True
	regEx.Global = True
	newString = regEx.Replace(userInput, "")
	Set regEx = nothing
	InputFilter = newString
  End Function 
  
  Function ExecuteSQL(ByRef conn, ByVal sql)
    On Error Resume Next 
    Dim ErrMsg
    ErrMsg = ""
    conn.Execute(sql)
    if Err <> 0 then
      ErrMsg = BuildErrMsg(ErrMsg, "Error No: " & Err.Number & " " & Err.Description)
    end if
    ExecuteSQL = ErrMsg
  End Function 
  
  Function ExecuteSQLForRS(ByRef conn, ByVal sql, ByRef rs)
    On Error Resume Next
    Dim ErrMsg
    ErrMsg = ""
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.CursorLocation = adUseClient
    rs.Open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    if Err <> 0 then
      ErrMsg = BuildErrMsg(ErrMsg, "Error No: " & Err.Number & " " & Err.Description)
    end if
    ExecuteSQLForRS = ErrMsg
  End Function
  
  Function TranslateID2SSN(ByRef conn, ByVal StudentID)
    Dim sql, sqlErr, rs, SSN
    SSN = ""
    if IsNumeric(StudentID) then
      sql = "SELECT SSN FROM dbo.Students WHERE StudentID = " & StudentID
      sqlErr = ExecuteSQLForRS(conn, sql, rs)
      if sqlErr = "" then
        if not rs.Eof then
          SSN = rs("SSN")
        end if
        rs.close
      end if
    end if
    TranslateID2SSN = SSN  
    Set rs = Nothing
  End Function
  
  Sub ShootEmail(ByVal rs, ByVal readersrs, ByVal EssayID)
    Dim mailstr, subject

    mailstr = mailstr & "Hello " & readersrs("FullName") & "," & vbNewLine & vbNewLine
    mailstr = mailstr & "A Writing Sample Needs to be Scored.  Access sample by clicking the following link:" & vbNewLine & vbNewLine
    mailstr = mailstr & "https://placement.stcc.edu/admin/review.asp?EssayID=" & EssayID & "&ReaderID=" & readersrs("ID") & vbNewLine & vbNewLine
    
    mailstr = mailstr & "English/Reading Placement Testing Results" & vbNewLine & vbNewLine
    mailstr = mailstr & rs("lastname") & ", " & rs("firstname") & vbNewLine
       
    mailstr = mailstr & "ENGLISH: " 
    If IsNull(rs("EnglPlacement")) Then
      mailstr = mailstr & "NO PLACEMENT TAKEN"
    Else
      mailstr = mailstr & FormatPlacement(rs("EnglPlacement")) & " "
      mailstr = mailstr & FormatDateTime(rs("EnglDate")) & " "
      mailstr = mailstr & ZPadStr(CInt(rs("EnglScore")), 3)
    End If
    mailstr = mailstr & vbNewLine     
    
    mailstr = mailstr & "READING: " 
    If IsNull(rs("ReadPlacement")) Then
      mailstr = mailstr & "NO PLACEMENT TAKEN"
    Else
      mailstr = mailstr & FormatPlacement(rs("ReadPlacement")) & " "
      mailstr = mailstr & FormatDateTime(rs("ReadDate")) & " "
      mailstr = mailstr & ZPadStr(CInt(rs("ReadScore")), 3)
    End If
    mailstr = mailstr & vbNewLine

    subject = "Writing Sample For " & rs("lastname") & ", " & rs("firstname") & " Requires Evaluation"
    Call mailmessageEx(readersrs("Email"), "webmaster@stcc.edu", "", "", subject, mailstr, mailPlain)
  end Sub

  function EncodeSSN(ByVal SSN)
      If Len(SSN) <= 7 Then
        SSN = "XX" & ZPadStr(SSN, 7)
      End If
      EncodeSSN = SSN
  End Function
  
  Function IsBlank(Value)
    'Returns True if Empty or NULL or Zero
    If IsEmpty(Value) or IsNull(Value) Then
        IsBlank = True
    ElseIf IsNumeric(Value) Then
        If Value = 0 Then ' Special Case 
            IsBlank = True  ' Change to suit your needs
        End If      
    ElseIf IsObject(Value) Then
        If Value Is Nothing Then
            IsBlank = True
        End If
    ElseIf VarType(Value) = vbString Then
        If Value = "" Then
            IsBlank = True
        End If      
    Else
        IsBlank = False
    End If
  End Function

%>