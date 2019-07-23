<%
  Option Explicit
  On Error Resume Next
%>
<!-- #include file="library/library.asp" -->
<%
  Dim conn, rs, sql, count, i, ErrMsg, sqlErr

  if IsEmpty(Session("User")) then
    Response.Redirect "login.asp"
  elseif not Session("User")("IsAdmin") then 
    Response.Redirect "default.asp"
  end if
  
  Set conn = openConnection(Application("ConnectionString"))
  
  ErrMsg = ""
  if Request.ServerVariables("REQUEST_METHOD") = "POST" and Request("MakeChanges") = "Make Changes" then
    conn.BeginTrans
    for i = 1 to Request("MathStatus").Count
      sql = "Update [MathPlacement] Set Status = 0 WHERE SSN='" & Request("MathStatus")(i) & "'"
      sqlErr = ExecuteSQL(conn, sql)
      ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
    next
    
    for i = 1 to Request("EnglStatus").Count
      sql = "Update [EnglishPlacement] Set Status = 0 WHERE SSN='" & Request("EnglStatus")(i) & "'"
      sqlErr = ExecuteSQL(conn, sql)
      ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
    next
   
    for i = 1 to Request("ReadStatus").Count
      sql = "Update [ReadingPlacement] Set Status = 0 WHERE SSN='" & Request("ReadStatus")(i) & "'"
      sqlErr = ExecuteSQL(conn, sql)
      ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
    next
   
    for i = 1 to Request("TypeStatus").Count
      sql = "Update [TypingPlacement] Set Status = 0 WHERE SSN='" & Request("TypeStatus")(i) & "'"
      sqlErr = ExecuteSQL(conn, sql)
      ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
    next
    
    for i = 1 to Request("ExitStatus").Count
      sql = "Update [ReadingExitPlacement] Set Status = 0 WHERE SSN='" & Request("ExitStatus")(i) & "'"
      sqlErr = ExecuteSQL(conn, sql)
      ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
    next

    if ErrMsg = "" then
      conn.CommitTrans
    else
      conn.RollbackTrans
    end if
  end if 
%>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<title><% =SYSTEM_NAME %> - Upload Failure Report</title>
<link rel="stylesheet" href="css/placement.css" />
<link rel="stylesheet" href="css/gototop.css" />
<style>
table {
	margin-right: auto;
	margin-left: auto;
}
.detailRow {
	height: 50px;
}
@media print {
  #report a {
	text-decoration:none;
	color:inherit;
  }
}
</style>
</head>

<body>
<div class="center">
<%
  Call MakeHeader("Upload Failure Report")
  if ErrMsg <> "" then
    Response.Write "<h3 class=""red center"">Error(s) Occurred in Submission<br />" & ErrMsg & "</h3>" & vbNewLine
  end if
%>
<p class="bold">
<a style="margin-right: 10px" href="utilrept.asp">Utilities &amp; Reports</a>
<a style="margin-right: 10px" href="default.asp?ResetQuery=true">Main Menu</a>
<a href="default.asp?Logout=true">Log Out</a></p>
<hr />

<form method="post">
<table id="report" style="margin-top: 20px;" border="1" cellpadding="2" cellspacing="2">
	<tr>
		<th colspan="5">&nbsp;</th>
		<th colspan="5">Status Values (Click to Remove from Upload)</th>
	</tr>
	<tr>
		<th>Count</th>
		<th>System ID</th>
		<th>SSN</th>
		<th>Last Name</th>
		<th>First Name</th>
		<th>Math </th>
		<th>English</th>
		<th>Reading</th>
		<th>Keyboarding</th>
		<th>ReadExit</th>
	</tr>
	<%
	  sql = "SELECT StudentID, SSN, LastName, FirstName, MathStatus, EnglStatus, ReadStatus, TypeStatus, ExitStatus FROM "
      sql = sql & "dbo.[Full Placement Testing Join] WHERE "
      sql = sql & "EnglStatus <> 0 or MathStatus <> 0 or ExitStatus <> 0 or ReadStatus <> 0 or TypeStatus <> 0 "
      sql = sql & "ORDER BY LastName, FirstName"

      Set rs = Server.CreateObject("ADODB.Recordset")
      rs.CursorLocation = adUseClient
      rs.Open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
      Set rs.ActiveConnection = Nothing
      count = 0  
	  do while not rs.eof
	    count = count + 1
	    Response.Write "<tr>" & vbNewLine
	    Response.Write "<td align=""center"">" & count & "</td>" & vbNewLine
	    Response.Write "<td><a href=""editRecord.asp?value=" & rs("StudentID") & """>" & rs("StudentID") & "</a></td>" & vbNewLine
	    Response.Write "<td>" & rs("SSN") & "</td>" & vbNewLine
        Response.Write "<td>" & rs("LastName") & "</td>" & vbNewLine
	    Response.Write "<td>" & rs("FirstName") & "</td>" & vbNewLine
	    Response.Write "<td align=""center"">"
	    if rs("MathStatus") <> 0 then
	      Response.Write "<input type=""checkbox"" name=""MathStatus"" value=""" & rs("SSN") & """ />"
	    else
  	      Response.Write "&nbsp;"
	    end if
	    Response.Write "</td>" & vbNewLine
	    Response.Write "<td align=""center"">"
	    if rs("EnglStatus") <> 0 then
	      Response.Write "<input type=""checkbox"" name=""EnglStatus"" value=""" & rs("SSN") & """ />"
	    else
  	      Response.Write "&nbsp;"
	    end if
	    Response.Write "</td>" & vbNewLine
	    Response.Write "<td align=""center"">"
  	    if rs("ReadStatus") <> 0 then
	      Response.Write "<input type=""checkbox"" name=""ReadStatus"" value=""" & rs("SSN") & """ />"
	    else
  	      Response.Write "&nbsp;"
	    end if
	    Response.Write "</td>" & vbNewLine
	    Response.Write "<td align=""center"">"
  	    if rs("TypeStatus") <> 0 then
	      Response.Write "<input type=""checkbox"" name=""TypeStatus"" value=""" & rs("SSN") & """ />"
	    else
  	      Response.Write "&nbsp;"
	    end if
	    Response.Write "</td>" & vbNewLine
	    Response.Write "<td align=""center"">"
   	    if rs("ExitStatus") <> 0 then
	      Response.Write "<input type=""checkbox"" name=""ExitStatus"" value=""" & rs("SSN") & """ />"
	    else
  	      Response.Write "&nbsp;" 
	    end if
	    Response.Write "</td>" & vbNewLine
	    Response.Write "</tr>" & vbNewLine
	    rs.movenext
	  loop
	  rs.close
	  Set rs = Nothing
    %>
    <tr>
    <td class="detailRow center" colspan="10">
    <%
      if count = 0 then
        Response.Write "<b>No Upload Failure Records Found</b>"
      else
    %>
      <input name="MakeChanges" type="submit" value="Make Changes" style="margin-right: 100px" />
      <input name="Cancel Changes" type="submit" value="Cancel Changes" />
    <%
      end if
    %>
    </td>
    </tr>
</table>
</form>
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