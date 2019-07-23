<% 
  Option Explicit
  On Error Resume Next
%>
<!-- #include file="library/library.asp" -->
<%
  Dim conn, rs, sql, ErrMsg, sqlErr, tempstr
  Dim StartDate, FinalDate
  Dim tmpStartDate, tmpFinalDate
  
  if IsEmpty(Session("User")) then
    Response.Redirect "login.asp"
  elseif not Session("User")("IsEssay") then 
    Response.Redirect "default.asp"
  end if
  
  ' Disable this page
  Response.Redirect "default.asp"

  Set conn = openConnection(Application("ConnectionString"))
  
  Sub DisplayRecordSet(ByVal rs, ByVal StartDate, ByVal FinalDate)
    Response.Write "<h1>Essay Readers Count Report</h1>" & vbNewLine
    Response.Write "<h2>"
    if StartDate = "" and FinalDate = "" then
      Response.Write "All Dates"
    elseif StartDate <> "" and FinalDate <> "" then
      Response.Write "From: " & StartDate & " &nbsp; To: " & FinalDate
    elseif StartDate <> "" then
      Response.Write "From: " & StartDate
    else
      Response.Write "To: " & FinalDate
    end if     
    Response.Write "</h2>" & vbNewLine    
    Response.Write "<table border=""1"" cellpadding=""2"" cellspacing=""2"" style=""margin: 20px auto"">" & vbNewLine 
    Response.Write "<tr>" & vbNewLine
    Response.Write "<th>Name</th><th>Total<br />Essays Sent</th><th>Total<br />Essays Read</th><th>Total<br />Essays Not Read</th>"
    Response.Write "</tr>" & vbNewLine
    do while not rs.Eof
      Response.Write "<tr>" & vbNewLine
      Response.Write "<td style=""text-align: left"">" & rs("FullName") & "</td>"
      Response.Write "<td>" & rs("TotalSent") & "</td>"
      Response.Write "<td>" & rs("TotalRead") & "</td>"
      Response.Write "<td>" & rs("TotalUnRead") & "</td>"
      Response.Write "</tr>" & vbNewLine
      rs.MoveNext
    loop
    Response.Write "</table>" & vbNewLine
  End Sub
  
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
    if ErrMsg = "" then
      sql = "select Z.*, (TotalSent - TotalRead) As TotalUnread From (select A.FullName, "
      sql = sql & "(select count(*) from EnglishEssaysJoin where ReaderID = A.ID AND EssayDate BETWEEN '" & tmpStartDate & "' AND '" & tmpFinalDate & "') as TotalSent, "
      sql = sql & "(select count(B.ReaderPlacement) from EnglishEssaysJoin B where ReaderID = A.ID AND EssayDate BETWEEN '" & tmpStartDate & "' AND '" & tmpFinalDate & "') as TotalRead "
      sql = sql & "from essayreaders A) Z ORDER BY FullName"
      sqlErr = ExecuteSQLForRs(conn, sql, rs)
      ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
    end if
  end if
%>
<!DOCTYPE html>
<html lang="en">

<head>
<meta charset="UTF-8" />
<title><% =SYSTEM_NAME %> - Essay Readers Count Report</title>
<link rel="stylesheet" href="css/placement.css" />
<link rel="stylesheet" href="css/calendar.css" />
<link rel="stylesheet" href="css/gototop.css" />
<script src="calendar_us.js"></script>
</head>

<body>
<div class="center">
<%
  Call MakeHeader("Essay Readers Count Report")
  if ErrMsg <> "" then
    Response.Write "<h3 class=""red"">Error(s) Occurred in Essay Readers Count Report<br />" & ErrMsg & "</h3>" & vbNewLine
  end if
%>
<p class="hideElement bold">
<a style="margin-right: 10px" href="essayCountReport.asp">Essay Readers Count Report</a>
<a style="margin-right: 10px" href="EssayUtilities.asp">Essay Utilities</a>
<a style="margin-right: 10px" href="default.asp?ResetQuery=true">Main Menu</a>
<a href="default.asp?Logout=true">Log Out</a>
</p>

<hr class="hideElement" />
<%
  if Request.ServerVariables("REQUEST_METHOD") = "GET" or ErrMsg <> "" then
%>
	<form method="post">
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
    Call DisplayRecordSet(rs, StartDate, FinalDate)
    rs.Close
    Set rs = Nothing
  end if  
%> 
</div>
<script src="//ajax.googleapis.com/ajax/libs/jquery/2.2.4/jquery.min.js"></script> 
<script src="jquery.gototop.js"></script> 
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