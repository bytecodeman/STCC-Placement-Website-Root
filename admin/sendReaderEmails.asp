<% 
  Option Explicit
  On Error Resume Next  
%>
<!-- #include file="library/library.asp" -->
<%
  Dim conn, rs, readers, sql, ErrMsg, sqlErr, tempstr
  Dim StartDate, FinalDate
  Dim tmpStartDate, tmpFinalDate
  Dim readerDict
  Dim tmpReaderID
    
  if IsEmpty(Session("User")) then
    Response.Redirect "login.asp"
  elseif not Session("User")("IsEssay") then 
    Response.Redirect "default.asp"
  end if
  
  ' Disable this page
  Response.Redirect "default.asp"

  Set conn = openConnection(Application("ConnectionString"))
  
  Sub AppendUnique(ByVal x, ByVal val)
    If val <> "" And Not x.Exists(val) Then
      x.Add val, val
    End if
  End Sub
  
  Function StripID(ByVal x)
    Dim pos
    pos = Instr(1, x, "-")
    StripID = iif(pos = 0, x, Left(x, pos - 1))
  End Function

  Function StripName(ByVal x)
    Dim pos
    pos = Instr(1, x, "-")
    StripName = iif(pos = 0, x, Mid(x, pos + 1))
  End Function

  Function GrabReaderIDs(ByVal readerDict)
    Dim i, keys, retval
    keys = readerDict.keys
    retval = ""
    For i = 0 To UBound(keys)
      retval = retval & StripID(keys(i))
      if i < UBound(keys) then
        retval = retval & ", "
      end if
    next
    GrabReaderIDs = retval
  End Function
  
  Function SendEmailsToEssayReaders(ByVal EssayID, ByVal SSN, ByVal readerDict)
    On Error Resume Next
    Dim placers, readersrs, temprs, sql, sqlErr, ErrMsg, sendEmail
    sql = "SELECT * FROM [Full Placement Testing Join] WHERE SSN = '" & SSN & "'"
    sqlErr = ExecuteSQLForRs(conn, sql, placers)
    ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
    if placers.EOF Then
      ErrMsg = BuildErrMsg(ErrMsg, "STCC Placement: Cannot Find SSN " & SSN)
    else
      sql = "Select ID, Code, FullName, Email From EssayReaders WHERE ID IN (" & GrabReaderIDs(readerDict) & ")"
      sqlErr = ExecuteSQLForRs(conn, sql, readersrs)
      ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
      do while not readersrs.eof
        sql = "SELECT EssayID, ReaderID, ReaderPlacement FROM ContactReaders WHERE EssayID = " & EssayID & " AND ReaderID = " & readersrs("ID")
        sqlErr = ExecuteSQLForRS(conn, sql, temprs)
        ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
        sendEmail = False
        if temprs.Eof Then
          sql = "INSERT INTO ContactReaders (EssayID, SSN, ReaderID) VALUES (" & EssayID & ", '" & SSN & "', " & readersrs("ID") & ")"
          sqlErr = ExecuteSQL(conn, sql)
          ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
          sendEmail = true
        elseif IsNull(temprs("ReaderPlacement")) then
          sendEmail = true
        end if
        if ErrMsg <> "" then
          sendEmail = False
        end if
        if sendEmail then
          Call ShootEmail(placers, readersrs, EssayID)
        end if          
        readersrs.moveNext
      loop
    end if
    placers.close
    Set placers = Nothing
    readersrs.close
    Set readersrs = Nothing
    SendEmailsToEssayReaders = ErrMsg
  End Function

  'Sub Append(ByRef x, ByVal val)
  '  Dim Count
  '  On Error Resume Next
  '  Count = UBound(x)
  '  On Error Goto 0
  '  ReDim Preserve x(Count + 1) 
  '  x(Count) = val
  'End Sub

  'Sub AppendUnique(ByRef x, ByVal val)
  '  Dim Count
  '  On Error Resume Next
  '  Count = UBound(x)
  '  On Error Goto 0
  '  Dim i
  '  For i = 0 To Count - 1
  '    If x(i) = val Then
  '      Exit Sub 
  '	 End if
  '  next
  '  ReDim Preserve x(Count + 1) 
  '  x(Count) = val
  'End Sub

  'Sub Print(ByVal x)
  '  Dim i
  '  For i = 0 To UBound(x) - 1
  '    WScript.Echo "* " & x(i)
  '  next
  'End Sub


  ErrMsg = ""
  if Request.ServerVariables("REQUEST_METHOD") = "POST" then
    StartDate = Trim(Request("StartDate"))
    FinalDate = Trim(Request("FinalDate"))
    if StartDate = "" then
      tmpStartDate = "01/01/08"
    else
      tmpStartDate = StartDate
    end if
    if FinalDate = "" then
      tmpFinalDate = Date()
    else
      tmpFinalDate = FinalDate
    end if
    if not IsDate(tmpStartDate) then
      ErrMsg = BuildErrMsg(ErrMsg, "Bad Starting Date")
    end if
    if not IsDate(tmpFinalDate) then
      ErrMsg = BuildErrMsg(ErrMsg, "Bad Final Date")
    end if
    if IsDate(StartDate) and IsDate(FinalDate) then
      if CDate(FinalDate) < CDate(StartDate) then
        ErrMsg = BuildErrMsg(ErrMsg, "Final Date is before Starting Date")
      end if
    end if
    
    Set readerDict= Server.CreateObject("Scripting.Dictionary")
    
    Call AppendUnique(readerDict, Trim(Request("Reader1")))
    Call AppendUnique(readerDict, Trim(Request("Reader2")))
    Call AppendUnique(readerDict, Trim(Request("Reader3")))
    if readerDict.Count = 0 then
      ErrMsg = BuildErrMsg(ErrMsg, "No Readers Specified")
    end if
    
    if ErrMsg = "" then
      sql = "SELECT SSN, LastName, FirstName, EnglPlacement, EssayID FROM [Full Placement Testing Join] WHERE "
      sql = sql & "EssayDate Between '" & tmpStartDate & "' AND '" & tmpFinalDate & "' AND " 
      sql = sql & "EnglPlacement = 'ESSAY' AND EssayID Is Not Null"         
      sqlErr = ExecuteSQLForRs(conn, sql, rs)
      ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
    end if
        
  end if
%>
<!DOCTYPE html>
<html lang="en">

<head>
<meta charset="UTF-8" />
<title><% =SYSTEM_NAME %> - Send Emails To Readers</title>
<link rel="stylesheet" href="css/placement.css" />
<link rel="stylesheet" href="css/calendar.css" />
<link rel="stylesheet" href="css/gototop.css" />
<script src="js/calendar_us.js"></script>
</head>

<body>
<div class="center">
<%
  Call MakeHeader("Send Emails to Readers")
  if ErrMsg <> "" then
    Response.Write "<h3 class=""red"">Error(s) Occurred<br />" & ErrMsg & "</h3>" & vbNewLine
  end if
%>
<p class="hideElement bold">
<a style="margin-right: 10px" href="sendReaderEmails.asp">Send Emails To Readers Utility</a>
<a style="margin-right: 10px" href="EssayUtilities.asp">Essay Utilities</a>
<a style="margin-right: 10px" href="default.asp?ResetQuery=true">Main Menu</a>
<a href="default.asp?Logout=true">Log Out</a>
</p>
<hr class="hideElement" />
<%
  if ErrMsg <> "" OR Request.ServerVariables("REQUEST_METHOD") = "GET" then
    sql = "SELECT ID, FullName FROM EssayReaders Order By FullName"
    Set readers = conn.Execute(sql)
%>
	<form method="post">
<p class="bold">Select Readers:</p>
  <p>
  <select name="Reader1" style="margin-right: 10px">
  <option value="">Choose Reader 1</option>
  <%
    readers.MoveFirst
    do while not readers.EOF
      Response.Write "<option value=""" & readers("ID") & "-" & readers("FullName") & """" & iif(Trim(Request("Reader1")) = readers("ID") & "-" & readers("FullName"), " selected=""selected""", "") & ">" & readers("FullName") & "</option>" & vbNewLine
      readers.MoveNext
    Loop
  %>
  </select>
  <select name="Reader2" style="margin-right: 10px">
  <option value="">Choose Reader 2</option>
  <%
    readers.MoveFirst
    do while not readers.EOF
      Response.Write "<option value=""" & readers("ID") & "-" & readers("FullName") & """" & iif(Trim(Request("Reader2")) = readers("ID") & "-" & readers("FullName"), " selected=""selected""", "") & ">" & readers("FullName") & "</option>" & vbNewLine
      readers.MoveNext
    Loop
  %>
  </select>
  <select name="Reader3">
  <option value="">Choose Reader 3</option>
  <%
    readers.MoveFirst
    do while not readers.EOF
      Response.Write "<option value=""" & readers("ID") & "-" & readers("FullName") & """" & iif(Trim(Request("Reader3")) = readers("ID") & "-" & readers("FullName"), " selected=""selected""", "") & ">" & readers("FullName") & "</option>" & vbNewLine
      readers.MoveNext
    Loop
  %>
  </select>
  </p>
  <%
    readers.Close
    Set readers = Nothing
  %>
  <p class="bold">Select Date Range:</p>
		<p class="bold">Start Date:
		<input type="text" id="StartDate" name="StartDate" size="10" maxlength="10" value="<%=StartDate%>" /> : 12:00am
		<script type="text/javascript">
		  new tcal ({
		    'controlname': 'StartDate'
		  })
		</script>
		&nbsp;&nbsp;&nbsp; Final Date:
		<input type="text" id="FinalDate" name="FinalDate" size="10" maxlength="10" value="<%=FinalDate%>" /> : 12:00am
		<script type="text/javascript">
		  new tcal ({
		    'controlname': 'FinalDate'
		  })
		</script></p>
		<p class="bold">Dates are entered in MM/DD/YY format.<br />
		Empty Dates form an open ended boundary. Dates are inclusive.</p>
		<p><input type="submit" value="Submit" /></p>
	</form>
<%
  else
%>
<div style="width: 400px; text-align: left; margin:auto">
<% 
  
  Response.Write "<h4>Emails Sent to:</h4>" & vbNewLine
  Response.Write "<ol class=""bold"">" & vbNewLine
  Dim keys, i
  keys = readerDict.Keys
  For i = 0 To readerDict.Count - 1 
    Response.Write "<li>" & StripName(keys(i)) & "</li>" & vbNewLine
  Next
  Response.Write "</ol>" & vbNewLine

  Response.Write "<h4>Essays Processed for:</h4>" & vbNewLine
  Response.Write "<ol class=""bold"">" & vbNewLine
  do while not rs.eof
    ErrMsg = SendEmailsToEssayReaders(rs("EssayID"), rs("SSN"), readerDict)
    if ErrMsg = "" then
      Response.Write "<li style=""color: green"">" & "SUCCESS: " & rs("SSN") & " " & rs("FirstName") & " " & rs("LastName") & "</li>" & vbNewLine
    else
      Response.Write "<li style=""color: red"">" & "FAILED: " & ErrMsg & "<br />" & rs("SSN") & " " & rs("FirstName") & " " & rs("LastName") & "</li>" & vbNewLine
    end if
    rs.movenext
  loop
  Response.Write "</ol>" & vbNewLine
%>
</div>
<%
    Set readerDict = Nothing
  end if  
%>
 
</div>
<script src="//ajax.googleapis.com/ajax/libs/jquery/2.2.4/jquery.min.js"></script> 
<script src="js/jquery.gototop.js"></script> 
<script>
$(function(){
  $("#toTop").gototop({ container: "body" });
});
</script>
</body>
</html>
<%
  closeConnection(conn)
%>