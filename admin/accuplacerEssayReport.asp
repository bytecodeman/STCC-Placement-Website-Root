<% 
  Option Explicit
  On Error Resume Next  
%>
<!-- #include file="library/library.asp" -->
<%
  Dim conn, rs, sql, ErrMsg, sqlErr
  Dim StartDate, FinalDate
  Dim tmpStartDate, tmpFinalDate
    
  if IsEmpty(Session("User")) then
    Response.Redirect "login.asp"
  elseif not Session("User")("IsEssay") then 
    Response.Redirect "default.asp"
  end if

  Set conn = openConnection(Application("ConnectionString"))

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
      sql = "SELECT SSN, LastName, FirstName, EnglDate FROM [Full Placement Testing Join] WHERE "
      sql = sql & "EnglDate Between '" & tmpStartDate & "' AND '" & tmpFinalDate & "' AND " 
      sql = sql & "EnglPlacement = 'ESSAY' AND WPScore = -1"         
      sqlErr = ExecuteSQLForRs(conn, sql, rs)
      ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
    end if
        
  end if
%>
<!DOCTYPE html>
<html lang="en">

<head>
<meta charset="UTF-8" />
<title><% =SYSTEM_NAME %> - Accuplacer Essay Report</title>
<link rel="stylesheet" href="css/placement.css" />
<link rel="stylesheet" href="css/calendar.css" />
<link rel="stylesheet" href="css/gototop.css" />
<style>
#undeterminedEssays {
  margin: auto;
  border: thin black solid;
}
#undeterminedEssays td, #undeterminedEssays th {
	margin: 2px;
	padding: 2px;
	border: thin black solid;
}
</style>
<script src="js/calendar_us.js"></script>
</head>

<body>
<div class="center">
<%
  Call MakeHeader("Accuplacer Essay Report")
  if ErrMsg <> "" then
    Response.Write "<h3 class=""red"">Error(s) Occurred<br />" & ErrMsg & "</h3>" & vbNewLine
  end if
%>
<p class="hideElement bold">
<a style="margin-right: 10px" href="accuplacerEssayReport.asp">Accuplacer Essay Report</a>
<a style="margin-right: 10px" href="EssayUtilities.asp">Essay Utilities</a>
<a style="margin-right: 10px" href="default.asp?ResetQuery=true">Main Menu</a>
<a href="default.asp?Logout=true">Log Out</a>
</p>
<hr class="hideElement" />
<%
  if ErrMsg <> "" OR Request.ServerVariables("REQUEST_METHOD") = "GET" then
%>
	<form method="post">
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
  <h4>Students with Accuplacer Undetermined Essay Status</h4>
  <%
  if rs.eof then
    Response.Write "<h4>No Students Found</h4>"
  else
  %>
    <table id="undeterminedEssays">
    <tr>
    <th>Student ID</th>
    <th>First name</th>
    <th>Last Name</th>
    <th>Test Date</th>
    </tr>
    <%
      do while not rs.eof
        Response.Write "<tr><td>" & iif(Left(rs("SSN"), 2) = "XX", Right(rs("SSN"), 7), rs("SSN")) & _
                       "</td><td>" & rs("FirstName") & "</td><td>" & rs("LastName") & _
                       "</td><td>" & rs("EnglDate") & _
                       "</td></tr>" & vbNewLine
        rs.movenext
      loop
    %>
    </table>
<%
  end if
%>
</div>
<%
  end if  
%>
 
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