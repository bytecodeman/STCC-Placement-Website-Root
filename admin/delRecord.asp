<% 
  Option Explicit
  On Error Resume Next
%>
<!-- #include file="library/library.asp" -->
<%
  Dim conn, rs, sql
  Dim ErrMsg, sqlErr
  Dim SSN, StudentID
  
  if IsEmpty(Session("User")) then
    Response.Redirect "login.asp"
  elseif not Session("User")("IsWriter") then 
    Response.Redirect "default.asp"
  elseif Request("value") = "" then
    Response.Redirect "default.asp?ResetQuery=true"
  else
    StudentID = Request("value")
    if Request("Phase") = "3" and Request("Response") = "OK" then
      Response.Redirect "default.asp"
    end if
  end if

  Set conn = openConnection(Application("ConnectionString"))

  ErrMsg = ""
  SSN = TranslateID2SSN(conn, StudentID)
  if SSN = "" then
    ErrMsg = BuildErrMsg(ErrMsg, "Record Cannot Be Located")
  else
    sql = "SELECT LastName, FirstName FROM dbo.Students WHERE StudentID = " & StudentID
    sqlErr = ExecuteSQLForRS(conn, sql, rs)
    ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
    if Request.ServerVariables("REQUEST_METHOD") = "POST" then
      if Request("Phase") = "2" and Request("Response") = "YES" then
        conn.BeginTrans
        sql =       "DELETE STUDENTS             WHERE SSN = '" & SSN & "' "
        sql = sql & "DELETE MathPlacement        WHERE SSN = '" & SSN & "' "
        sql = sql & "DELETE EnglishPlacement     WHERE SSN = '" & SSN & "' "
        sql = sql & "DELETE ReadingPlacement     WHERE SSN = '" & SSN & "' " 
        sql = sql & "DELETE TypingPlacement      WHERE SSN = '" & SSN & "' "
        sql = sql & "DELETE ReadingExitPlacement WHERE SSN = '" & SSN & "' "
        sql = sql & "DELETE ContactReaders       WHERE SSN = '" & SSN & "' "
        sql = sql & "DELETE EnglishEssays        WHERE SSN = '" & SSN & "' "
        sqlErr = ExecuteSQL(conn, sql)
        ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
        if ErrMsg = "" then
          conn.CommitTrans
        else
          conn.RollbackTrans
        end if
      elseif Request("Response") = "NO" then
        Response.Redirect "default.asp"
      end if
    end if
  end if
%>
<!DOCTYPE html>
<html lang="en">

<head>
<meta charset="UTF-8" />
<title><% =SYSTEM_NAME %> - Delete Record</title>
<link rel="stylesheet" href="css/placement.css" />
<link rel="stylesheet" href="css/gototop.css" />
</head>

<body>
<div class="center">

<%
  Call MakeHeader("Delete Record")
  if ErrMsg <> "" then
    Response.Write "<h3 class=""red"">Error(s) Occurred in Record Deletion<br />" & ErrMsg & "</h3>" & vbNewLine
  end if
%>
	<p class="bold">
    <a style="margin-right: 10px" href="addRecord.asp">Add Record</a>
	<a style="margin-right: 10px" href="viewRecord.asp?value=<%=StudentID%>">View Record</a>
    <a style="margin-right: 10px" href="editRecord.asp?value=<%=StudentID%>">Edit Record</a>
    <a style="margin-right: 10px" href="testdetails.asp?value=<%=StudentID%>">Edit Test Details</a>
    <a style="margin-right: 10px" href="delRecord.asp?value=<%=StudentID%>">Delete Record</a>
    <a style="margin-right: 10px" href="default.asp">Query List</a>
    <a style="margin-right: 10px" href="default.asp?RefineSearch=true">Refine Search</a>
    <a style="margin-right: 10px" href="default.asp?ResetQuery=true">New Query</a>
    <a href="default.asp?Logout=true">Log Out</a>
    </p>
	<hr />
	
	<form method="post">
<%
  if Request("Phase") = "" then
    if ErrMsg = "" then
%>
		<h3>Are you sure you want to delete<br />
		the Placement Records for <%=rs("FirstName") & " " & rs("LastName")%></h3>
		<p><input type="hidden" name="Phase" value="1" />
		<input type="submit" name="Response" value="YES" style="margin-right: 50px" />
		<input type="submit" name="Response" value="NO" /></p>
<%
    end if
  elseif Request("Phase") = "1" then
%>
		<h2>Are you ABSOLUTELY sure you want to delete<br />
		the Placement Records for <%=rs("FirstName") & " " & rs("LastName")%></h2>
		<p><input type="hidden" name="Phase" value="2" />
		<input type="submit" name="Response" value="YES" style="margin-right: 50px" />
		<input type="submit" name="Response" value="NO" /></p>
<%
  elseif Request("Phase") = "2" then
    if ErrMsg = "" then
      Response.Write "<h3 align=""center"">Record Deleted!!!</h3>" & vbNewLine
    end if
%>
		<p><input type="hidden" name="Phase" value="3" />
		   <input type="submit" name="Response" value="OK" /></p>
<%   
  end if
%>
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