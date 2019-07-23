<% 
  Option Explicit
  On Error Resume Next
%>
<!-- #include file="library/library.asp" -->
<%
  if IsEmpty(Session("User")) then
    Response.Redirect "login.asp"
  elseif not Session("User")("IsAdmin") then 
    Response.Redirect "default.asp"
  end if
  
  Dim username, password, confirmation
  Dim encUsername, encPassword
  Dim ErrMsg, sqlErr
  Dim conn, rs, sql
  
  Const codeKey = "strassh"
 
  function XORDecryption(ByVal DataIn)
    Dim lonDataPtr
    Dim strDataOut : strDataOut = ""
    Dim intXOrValue1, intXOrValue2 
    For lonDataPtr = 1 To Len(DataIn) \ 2
        'The first value to be XOr-ed comes from the data to be encrypted
        intXOrValue1 = CInt("&H0" & Mid(DataIn, lonDataPtr * 2 - 1, 2))
        'The second value comes from the code key
        intXOrValue2 = Asc(Mid(codeKey, (lonDataPtr Mod Len(codeKey)) + 1, 1))
        strDataOut = strDataOut & Chr(intXOrValue1 Xor intXOrValue2)
    Next
    XORDecryption = strDataOut
  end function

  function XOREncryption(ByVal DataIn)
    Dim lonDataPtr
    Dim strDataOut : strDataOut = ""
    Dim intXOrValue1, intXOrValue2
    For lonDataPtr = 1 To Len(DataIn)
        'The first value to be XOr-ed comes from the data to be encrypted
        intXOrValue1 = Asc(Mid(DataIn, lonDataPtr, 1))
        'The second value comes from the code key
        intXOrValue2 = Asc(Mid(codeKey, (lonDataPtr Mod Len(codeKey)) + 1, 1))
        strDataOut = strDataOut & Right("00" & Hex(intXOrValue1 Xor intXOrValue2), 2)
    Next
    XOREncryption =strDataOut
  End Function
    
  Set conn = openConnection(Application("ConnectionString"))

  ErrMsg = ""
  if Request.ServerVariables("REQUEST_METHOD") = "POST" then
    username = Trim(Request("username"))
    password = Trim(Request("password"))
    confirmation = Trim(Request("confirmation"))
    if username = "" then
      ErrMsg = BuildErrMsg(ErrMsg, "Username is empty.")
    end if
    if password = "" then
      ErrMsg = BuildErrMsg(ErrMsg, "Password is empty.")
    end if
    if confirmation = "" then
      ErrMsg = BuildErrMsg(ErrMsg, "Confirm Password is empty.")
    end if
    if password <> "" or confirmation <> "" then
      if password <> confirmation then
        ErrMsg = BuildErrMsg(ErrMsg, "Password and Confirm Password do not match.")
      end if
    end if
    if ErrMsg = "" then
      encUsername = XOREncryption(username)
      encPassword = XOREncryption(password)  
      sql = "UPDATE dbo.Configuration SET value = '" & encUsername & "' WHERE [key] = 'AccuPlacerSiteID'; "
      sql = sql & "UPDATE dbo.Configuration SET value = '" & encPassword & "' WHERE [key] = 'AccuPlacerPassword';"
      sqlErr = ExecuteSQL(conn, sql)
      ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
    end if
  else
    sql = "SELECT value FROM dbo.Configuration WHERE [key] = 'AccuPlacerSiteID'"
    sqlErr = ExecuteSQLForRS(conn, sql, rs)
    if sqlErr = "" then
      username = XORDecryption(rs("value"))
    else
      ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
    end if
  end if
%>
<!DOCTYPE html>
<html lang="en">

<head>
<meta charset="UTF-8" />
<title><% =SYSTEM_NAME %> - Change Password</title>
<link rel="stylesheet" href="css/placement.css" />
<link rel="stylesheet" href="css/gototop.css" />
<style>
table#pass {
	width: 600px;
	margin-left: auto;
	margin-right: auto;
	text-align: left;
	font-weight: bold;
}
table#pass input[type=text], table#pass input[type=password] {
	font-family: 'Courier New', Courier, monospace;
}
table#pass td {
	width: 50%;
	padding: 10px;
	text-align: left;
}

table#pass td:first-child {
	text-align: right;
}

table#pass td#submitter {
	text-align: center;
}

</style>
</head>

<body>

<div class="center">
	<%
  Call MakeHeader("Configure Accuplacer Login")
  if ErrMsg <> "" then
    Response.Write "<h3 class=""red"">Error(s) Occurred in Submission<br />" & ErrMsg & "</h3>" & vbNewLine
  elseif Request.ServerVariables("REQUEST_METHOD") = "POST" then
    Response.Write "<h3 class=""green"">Accuplacer Login Info Modified!</h3>" & vbNewLine
  end if
%>
	<p class="bold">
	<a href="default.asp?ResetQuery=true" style="margin-right: 10px">Main Menu</a>
	<a href="default.asp?Logout=true">Log Out</a></p>
	<hr />
	<div style="margin:auto; width: 80%">
	<h2 style="text-align: center">Change the Login Information for the Accuplacer website.</h2></div>
	<form id="changePasswordForm" method="post" autocomplete="off">
		<table id="pass">
			<tr>
				<td>Username:</td>
				<td>
				<input name="username" size="20" type="text" value="<%=username%>" required="required" autocomplete="off" /></td>
			</tr>
			<tr>
				<td>Password:</td>
				<td>
				<input name="password" size="20" type="password" required="required" autocomplete="off" /></td>
			</tr>
			<tr>
				<td>Confirm Password:</td>
				<td>
				<input name="confirmation" size="20" type="password" required="required" autocomplete="off"/></td>
			</tr>
			<tr>
				<td id="submitter" colspan="2">
				<input type="submit" value="Change Accuplacer Login Info" /></td>
			</tr>
		</table>
	</form>
</div>
<script src="//ajax.googleapis.com/ajax/libs/jquery/2.2.4/jquery.min.js"></script> 
<script src="js/jquery.gototop.js"></script> 
<script>
  $(function() {
    "use strict";
    $("#toTop").gototop({ container: "body" });
  });
</script>
</body>
</html>
<%
  closeConnection(conn)
%>