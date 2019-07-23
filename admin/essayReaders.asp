<%
  Option Explicit
  On Error Resume Next
%>
<!-- #include file="library/library.asp" -->
<%
  Dim rs, conn, sql
  Dim ErrMsg, sqlErr
  Dim id
  Dim Operation
  Dim SubOperation
  Dim Code
  Dim FirstName, LastName, FullName
  Dim Email
  
  if IsEmpty(Session("User")) then
    Response.Redirect "login.asp"
  elseif not Session("User")("IsAdmin") and not Session("User")("IsEssay") then 
    Response.Redirect "default.asp"
  end if

  ' Disable this page
  Response.Redirect "default.asp"

  Set conn = openConnection(Application("ConnectionString"))

  ErrMsg = ""
  Operation = "LIST"
  if Request.ServerVariables("REQUEST_METHOD") = "POST" then
    Operation = UCase(Trim(Request("Operation")))
    ID = UCase(Trim(Request("ID")))
    if Operation = "ADD" then
      Code = ""
      LastName = ""
      FirstName = ""
      Email = ""
    elseif Operation = "EDIT" or Operation = "DELETE" then
      if IsNumeric(ID) then
        sql = "SELECT Code, FirstName, LastName, FullName, Email FROM EssayReaders WHERE ID = " & ID
        sqlErr = ExecuteSqlForRs(conn, sql, rs)
        if sqlErr = "" then
          Code = rs("Code")
          LastName = rs("LastName")
          FirstName = rs("FirstName")
          FullName = rs("FullName")    
          Email = rs("Email")      
          rs.close
        else
          ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
        end if
        Set rs = Nothing
      else
        ErrMsg = BuildErrMsg(ErrMsg, "No Message Selected!")
        Operation = "LIST"
      end if
    elseif Operation = "CANCEL" then
      Operation = "LIST"
    elseif Operation = "SUBMIT" then
      SubOperation = UCase(Trim(Request("SubOperation")))
      Code = InputFilter(Trim(Request("Code")))
      FirstName = Replace(Trim(Request("FirstName")), "'", "''")
      LastName = Replace(Trim(Request("LastName")), "'", "''")
      Email = InputFilter(Trim(Request("Email")))
      if FirstName = "" then
        ErrMsg = BuildErrMsg(ErrMsg, "FirstName MUST be Specified")
      end if  
      if LastName = "" then
        ErrMsg = BuildErrMsg(ErrMsg, "LastName MUST be Specified")
      end if  
      if not ValidCode(Code) then
        ErrMsg = BuildErrMsg(ErrMsg, "Code must be 1, 2, 3, or Blank")
      end if
      if not EmailOK(Email) then
        ErrMsg = BuildErrMsg(ErrMsg, "Invalid Email Specified")
      end if      
      if ErrMsg = "" then
        sqlErr = PerformUpdate(SubOperation, ID)
        if sqlErr = "" then
          Operation = "LIST"
        else
          ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
          Operation = SubOperation
        end if
      else
        Operation = SubOperation
      end if
    elseif Operation = "YES" then    
      sqlErr = PerformUpdate("DELETE", ID)
      if sqlErr = "" then
        Operation = "LIST"
      else
        ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
        Operation = "DELETE"
      end if
    end if
  end if 
  if Operation = "LIST" then  
    sql = "SELECT ID, Code, FullName, Email FROM EssayReaders ORDER BY FullName ASC"
    sqlErr = ExecuteSQLForRs(conn, sql, rs)
    ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
  end if 
  
  Function PerformUpdate(ByVal Operation, ByVal ID)
    On Error Resume Next
    Dim conn, sql, ErrMsg, sqlErr
    Dim tmpCode, tmpFirstName, tmpLastName, tmpEmail
    tmpCode = iif(Trim(Code) = "", "NULL", Code)
    tmpFirstName = iif(Trim(FirstName) = "", "NULL", "'" & FirstName & "'")
    tmpLastName = iif(Trim(LastName) = "", "NULL", "'" & LastName & "'")
    tmpEmail = iif(Trim(Email) = "", "NULL", "'" & Email & "'")    
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open Application("ConnectionString")
    conn.BeginTrans
    ErrMsg = ""
    if Operation = "ADD" then
      sql = "Insert Into dbo.EssayReaders (Code, FirstName, LastName, Email) VALUES (" & tmpCode & ", " & tmpFirstName & ", " & tmpLastName & ", " & tmpEmail & ")"
    elseif Operation = "EDIT" then
      sql = "Update dbo.EssayReaders SET Code = " & tmpCode & ", FirstName = " & tmpFirstName & ", LastName = " & tmpLastName & ", Email = " & tmpEmail & " WHERE ID = " & ID
    elseif Operation = "DELETE" then
      sql = "Delete From dbo.EssayReaders Where ID = " & ID
    end if
    sqlErr = ExecuteSql(conn, sql)
    ErrMsg = BuildErrMsg(ErrMsg, sqlErr) 
    if ErrMsg = "" then
      conn.CommitTrans
    else
      conn.RollbackTrans
    end if
    PerformUpdate = ErrMsg
  End Function
  
  Function EmailOK(ByVal str)
    Dim newString, regEx
	Set regEx = New RegExp
	regEx.Pattern = "^([0-9a-z]([-.\w]*[0-9a-z])*@(([0-9a-z])+([-\w]*[0-9a-z])*\.)+[a-z]{2,9})$"
	regEx.IgnoreCase = True
	regEx.Global = True
	EmailOK = regEx.test(str)
	Set regEx = nothing
  End Function
  
  Function ValidCode(ByVal code)
    ValidCode = code = "" or code = "1" or code = "2" or code = "3"
  End Function
%>
<!DOCTYPE html>
<html lang="en">

<head>
<meta charset="UTF-8" />
<title><% =SYSTEM_NAME %> - Essay Readers Database Maintenance</title>
<link rel="stylesheet" href="css/placement.css" />
<link rel="stylesheet" href="css/gototop.css" />
<style>
form table {
	margin: 25px auto;
}
td.selectCol {
	text-align: center;
	width: 25px;	
}
th.selectCol {
	text-align: center;
	width: 25px;	
}
td.codeCol, th.codeCol {
	text-align: center;
	width: 50px;
}
td.nameCol, td.emailCol {
	text-align: left;
	width: 225px;
}
td.FirstNameCol, td.LastNameCol {
	text-align: left;
	width: 175px;
}
th.nameCol, th.emailCol {
	text-align: center;
	width: 225px;
}
th.FirstNameCol, th.LastNameCol {
	text-align: center;
	width: 175px;
}
.detailRow {
	height: 35px;
}

.buttonCell {
	text-align: center;
}
</style>
</head>

<body>
<div class="center">
<%
  Call MakeHeader("Essay Readers Database Maintenance")
  if ErrMsg <> "" then
    Response.Write "<h3 class=""red"">Error(s) Occurred in Submission<br />" & ErrMsg & "</h3>" & vbNewLine
  end if
%>
<p class="hideElement bold">
<a href="EssayReaders.asp">Essay Readers Database Maintenance</a><span style="margin-right: 10px">&nbsp;</span>
<a href="EssayUtilities.asp">Essay Utilities</a><span style="margin-right: 10px">&nbsp;</span>
<a href="default.asp?ResetQuery=true">Main Menu</a><span style="margin-right: 10px">&nbsp;</span>
<a href="default.asp?Logout=true">Log Out</a>
</p>
<hr />

<form method="post">
<%
  if Operation = "LIST" then
%>
	<table border="1" cellspacing="2" cellpadding="5">
		<tr>
			<th class="selectCol">&nbsp;</th>
			<th class="codeCol">Code</th>
			<th class="nameCol">Name</th>
			<th class="emailCol">Email</th>
		</tr>
<%  
  if Typename(rs) = "Recordset" then 
    if rs.state = adStateOpen then
      do while not rs.Eof
        id = rs("ID")
        Response.Write "<tr class=""detailRow"">" & vbNewline
        Response.Write "<td class=""selectCol""><input name=""ID"" type=""radio"" value=""" & id & """ /></td>" & vbNewLine
        Response.Write "<td class=""codeCol"">" & FixNull(rs("Code")) & "</td>" & vbNewLine
        Response.Write "<td class=""nameCol"">" & FixNull(rs("FullName")) & "</td>" & vbNewLine
        Response.Write "<td class=""emailCol"">" & FixNull(rs("Email")) & "</td>" & vbNewLine
        Response.Write "</tr>" & vbNewLine
        rs.MoveNext
      Loop
      rs.Close
    end if
    Set rs = Nothing
  end if
%>
		<tr>
			<td class="buttonCell" colspan="4">
			<input type="submit" style="margin-right: 35px" name="Operation" value="Add" />
			<input type="submit" style="margin-right: 35px" name="Operation" value="Edit" />
			<input type="submit" style="margin-right: 35px" name="Operation" value="Delete" />
			<input type="reset" name="Reset" /></td>
		</tr>
	</table>
<%
  elseif Operation = "ADD" or Operation = "EDIT" then
%>
	<table border="1" cellspacing="2" cellpadding="5">
		<tr>
			<th class="selectCol">&nbsp;</th>
			<th class="codeCol">Code</th>
			<th class="FirstNameCol">First Name</th>
			<th class="LastNameCol">Last Name</th>
			<th class="emailCol">Email</th>
		</tr>
		<tr class="detailRow">
			<td class="selectCol">&nbsp;</td>
			<td class="codeCol"><input name="Code" type="text" maxlength="1" size="1" value="<% =Code %>" /></td>
			<td class="FirstNameCol"><input name="FirstName" type="text" maxlength="20" size="20" value="<% =FirstName %>" /></td>
			<td class="LastNameCol"><input name="LastName" type="text" maxlength="25" size="20" value="<% =LastName %>" /></td>
			<td class="emailCol"><input name="Email" type="text" maxlength="100" size="30" value="<% =Email %>" /></td>
		</tr>
		<tr>
			<td class="buttonCell" colspan="5">
			<input type="submit" style="margin-right: 35px" name="Operation" value="Submit" />
			<input type="submit" style="margin-right: 35px" name="Operation" value="Cancel" />
			<input type="reset" name="Reset" />
			<input type="hidden" name="SubOperation" value="<%=Operation%>" />
			<input type="hidden" name="ID" value="<%=ID%>" /></td>
		</tr>
	</table>
<%
  elseif Operation = "DELETE" then
%>
	<table border="1" cellspacing="2" cellpadding="5">
		<tr>
			<th class="selectCol">&nbsp;</th>
			<th class="codeCol">Code</th>
			<th class="nameCol">Name</th>
			<th class="emailCol">Email</th>
		</tr>
		<tr class="detailRow">
			<td class="selectCol">&nbsp;</td>
			<td class="codeCol"><% =FixNull(Code) %></td>
			<td class="nameCol"><% =FixNull(FullName) %></td>
			<td class="emailCol"><% =FixNull(Email) %></td>
		</tr>
		<tr>
			<th class="buttonCell" colspan="4"><span style="margin-right: 35px">Delete this Record?</span>
			<input type="submit" style="margin-right: 35px" name="Operation" value="YES" />
			<input type="submit" name="Operation" value="Cancel" />
			<input type="hidden" name="ID" value="<%=ID%>" /></th>
		</tr>
	</table>
<%
  end if
%>
</form>

<p>Code Value must be one of the following values:</p>
<div>
<ul style="text-align:left;width:275px;margin:auto;">
	<li>1 - First Reader</li>
	<li>2 - Second Reader</li>
	<li>3 - Tie Breaker</li>
	<li>Blank - Not currently involved in essay reading</li>
</ul>
</div>
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