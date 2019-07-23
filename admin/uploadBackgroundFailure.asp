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

    for i = 1 to Request("BackgroundStatus").Count
      sql = "Update dbo.BackgroundQuestionResponses Set Status = 0 WHERE ID='" & Request("BackgroundStatus")(i) & "'"
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
  Call MakeHeader("Upload Background Questions Failure Report")
  if ErrMsg <> "" then
    Response.Write "<h3 class=""red center"">Error(s) Occurred in Submission<br />" & ErrMsg & "</h3>" & vbNewLine
  end if
%>
<p class="bold">
<a style="margin-right: 10px" href="uploadBackgroundFailure.asp">Upload Background Questions Failure Report</a>
<a style="margin-right: 10px" href="utilrept.asp">Utilities &amp; Reports</a>
<a style="margin-right: 10px" href="default.asp?ResetQuery=true">Main Menu</a>
<a href="default.asp?Logout=true">Log Out</a></p>
<hr />

<form method="post">
<table id="report" style="margin-top: 20px;" border="1" cellpadding="2" cellspacing="2">
	<tr>
		<th colspan="5">&nbsp;</th>
		<th>Status Values<br/>(Click to Remove from Upload)</th>
	</tr>
	<tr>
		<th>Count</th>
		<th>System ID</th>
		<th>Last Name</th>
		<th>First Name</th>
		<th>Date</th>
		<th><label for="SetClearStatus">Select/Clear:</label> <input type="checkbox" id="SetClearStatus" style="margin-right: 20px" />
		    <a href="#" id="ToggleStatus">Toggle</a></th>
	</tr>
	<%
      sql = "SELECT S.StudentID, S.SSN, S.LastName, S.FirstName, B.Status As BackgroundStatus, B.RespDate FROM "
      sql = sql & "dbo.[Students] S INNER JOIN dbo.BackgroundQuestionResponses B ON S.SSN = B.ID WHERE B.Status <> 0 "
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
 	    Response.Write "<td>" & rs("StudentID") & "</td>" & vbNewLine
        Response.Write "<td>" & rs("LastName") & "</td>" & vbNewLine
	    Response.Write "<td>" & rs("FirstName") & "</td>" & vbNewLine
	    Response.Write "<td>" & rs("RespDate") & "</td>" & vbNewLine
	    Response.Write "<td align=""center"">"
   	    if not IsNull(rs("BackgroundStatus")) then
   	      if rs("BackgroundStatus") <> 0 then
	        Response.Write "<input type=""checkbox"" name=""BackgroundStatus"" value=""" & rs("SSN") & """ />"
	      else
  	        Response.Write "&nbsp;" 
  	      end if	      
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
  
  function allTrue() {
    var falseFound = false;
    $("input[name=BackgroundStatus]").each(function() {
       if (!$(this).prop("checked")) {
          falseFound = true;
          return false;
       }
    });
    return !falseFound;
  }
  $("#SetClearStatus").click(function() {
    var checked = $(this).is(":checked");
    $("input[name=BackgroundStatus]").prop("checked", checked);
  });
  $("#ToggleStatus").click(function() {
    $("input[name=BackgroundStatus]").each(function() {
        $(this).prop("checked", !$(this).is(":checked"));
    });
    $("#SetClearStatus").prop("checked", allTrue());
    return false;
  });
  $("input[name=BackgroundStatus]").click(function() {
    $("#SetClearStatus").prop("checked", allTrue());
  });   
});
</script>
</body>
</html>
<%
  closeConnection(conn)
%>