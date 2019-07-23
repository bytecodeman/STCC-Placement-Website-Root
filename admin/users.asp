<%
  Option Explicit
  On Error Resume Next
%>
<!-- #include file="library/library.asp" -->
<%
  Dim rs, conn, sql
  Dim ErrMsg, sqlErr
  Dim Operation, SubOperation
  Dim ID, Username, Password, MustChange, IsAdmin, IsWriter, IsEssay
  
  if IsEmpty(Session("User")) then
    Response.Redirect "login.asp"
  elseif not Session("User")("IsUserManager") then 
    Response.Redirect "default.asp"
  end if

  Set conn = openConnection(Application("UserDirectoryConnectionString"))

  ErrMsg = ""
  Operation = "LIST"
  if Request.ServerVariables("REQUEST_METHOD") = "POST" then
    Operation = UCase(Trim(Request("Operation")))
    ID = Trim(Request("ID"))
    if Operation = "ADD" then
      if ID = "" then
        Username = ""
        Password = ""
        MustChange = false
        IsAdmin = false
        IsWriter = false
        IsEssay = false
      else
        ErrMsg = BuildErrMsg(ErrMsg, "Do Not Select a User When Adding")
        Operation = "LIST"
      end if
    elseif Operation = "ADD EXTERNAL USER" then
      if ID <> "" then
        ErrMsg = BuildErrMsg(ErrMsg, "Do Not Select a User When Adding External")
        Operation = "LIST"
      end if
    elseif Operation = "EDIT" or Operation = "DELETE" then
      if ID <> "" then
        sql = "SELECT UserName, Password, MustChangePassword, IsAdmin, IsWriter, IsEssay FROM dbo.PlacementTestingUsers WHERE Username = '" & ID & "'" 
        sqlErr = ExecuteSqlForRs(conn, sql, rs)
        if sqlErr = "" then
          Username = rs("Username")
          Password = rs("Password")
          MustChange = rs("MustChangePassword")
          IsAdmin = rs("IsAdmin")      
          IsWriter = rs("IsWriter")  
          IsEssay = rs("IsEssay")
          rs.close
        else
          ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
        end if
        Set rs = Nothing
      else
        ErrMsg = BuildErrMsg(ErrMsg, "Plase Select a User For the " & Operation & " Operation!")
        Operation = "LIST"
      end if
    elseif Operation = "CANCEL" then
      Operation = "LIST"
    elseif Operation = "SUBMIT" then
      SubOperation = UCase(Trim(Request("SubOperation")))
      if SubOperation = "ADD" or SubOperation = "EDIT" then
        Username = InputFilter(Trim(Request("Username")))
        Password = Trim(Request("md5Password"))
        MustChange = InputFilter(Trim(Request("MustChange"))) = "YES"
        IsAdmin = InputFilter(Trim(Request("IsAdmin"))) = "YES"      
        IsWriter = InputFilter(Trim(Request("IsWriter"))) = "YES"
        IsEssay = InputFilter(Trim(Request("IsEssay"))) = "YES"
        if UserName = "" then
          ErrMsg = BuildErrMsg(ErrMsg, "Username MUST be Specified")
        end if  
        if Password = BLANK_PASSWORD then
          ErrMsg = BuildErrMsg(ErrMsg, "Password MUST Be Specified")
        end if
        if ErrMsg = "" then
          sqlErr = PerformUpdate(conn, SubOperation, ID)
          if sqlErr = "" then
            Operation = "LIST"
          else
            ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
            Operation = SubOperation
          end if
        else
          Operation = SubOperation
        end if
      elseif SubOperation = "ADD EXTERNAL USER" then
        if ID <> "" then
          sqlErr = PerformUpdate(conn, "ADD EXTERNAL USER", ID)
          if sqlErr = "" then
            Operation = "LIST"
          else
            ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
            Operation = SubOperation
          end if         
        else
          ErrMsg = BuildErrMsg(ErrMsg, "No User Selected!")
          Operation = SubOperation
       end if
      end if
    elseif Operation = "YES" then    
      sqlErr = PerformUpdate(conn, "DELETE", ID)
      if sqlErr = "" then
        Operation = "LIST"
      else
        ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
        Operation = "DELETE"
      end if
    end if
  end if 
  if Operation = "LIST" then  
    sql = "SELECT UserName, Password, MustChangePassword, IsAdmin, IsWriter, IsEssay FROM dbo.PlacementTestingUsers ORDER BY UserName" 
    sqlErr = ExecuteSQLForRs(conn, sql, rs)
    ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
  elseif Operation = "ADD EXTERNAL USER" then
    sql = ""
    sql = sql + "SELECT DISTINCT NU.username FROM "
    sql = sql + "(SELECT U.username FROM [dbo].[SystemUser] SU INNER JOIN [dbo].[Users] U ON SU.username = U.username WHERE SystemID = " &  SYSTEM_ID & ") U "
    sql = sql + "RIGHT JOIN "
    sql = sql + "(SELECT U.username FROM [dbo].[SystemUser] SU INNER JOIN [dbo].[Users] U ON SU.username = U.username WHERE SystemID <> " &  SYSTEM_ID & ") NU "
    sql = sql + "ON U.Username = NU.Username WHERE U.Username Is Null"      
    sqlErr = ExecuteSQLForRs(conn, sql, rs)
    ErrMsg = BuildErrMsg(ErrMsg, sqlErr)  
  end if 
  
  Function PerformUpdate(ByVal conn, ByVal Operation, ByVal ID)
    On Error Resume Next
    Dim sql, ErrMsg, sqlErr
    conn.BeginTrans
    ErrMsg = ""
    if Operation = "ADD" then
      sql = ""
      sql = sql & "INSERT Into dbo.Users (Username, Password, MustChangePassword) VALUES ('" & Username & "', '" & Password & "', " & CInt(MustChange) & ")" & vbNewLine
      sql = sql & "INSERT Into dbo.PlacementUserAttributes (Username, IsAdmin, IsWriter, IsEssay) VALUES ('" & Username & "', " & CInt(IsAdmin) & ", " & CInt(IsWriter) & ", " & CInt(IsEssay) & ")" & vbNewLine
      sql = sql & "INSERT Into dbo.SystemUser (SystemID, Username) VALUES (" & SYSTEM_ID & ", '" & Username & "')" & vbNewLine    
    elseif Operation = "EDIT" then
      sql = ""
      sql = sql & "UPDATE dbo.Users SET Username = '" & Username & "', Password = '" & Password & "', MustChangePassword = " & CInt(MustChange) & " WHERE Username = '" & ID & "'" & vbNewLine
      sql = sql & "UPDATE dbo.PlacementUserAttributes SET Username = '" & Username & "', IsAdmin = " & CInt(IsAdmin) & ", IsWriter = " & CInt(IsWriter) & ", IsEssay = " & CInt(IsEssay) & " WHERE Username = '" & ID & "'" & vbNewLine
      sql = sql & "UPDATE dbo.SystemUser SET Username = '" & Username & "' WHERE SystemID = " & SYSTEM_ID & " AND Username = '" & ID & "'" & vbNewLine
    elseif Operation = "DELETE" then
      sql = ""
      sql = sql & "DELETE FROM dbo.SystemUser WHERE SystemID = " & SYSTEM_ID & " AND Username = '" & ID & "'" & vbNewLine
      sql = sql & "DELETE FROM dbo.PlacementUserAttributes WHERE Username = '" & ID & "'" & vbNewLine
      sql = sql & "DELETE FROM dbo.Users WHERE Username = '" & ID & "' AND NOT EXISTS(SELECT 1 FROM dbo.SystemUser WHERE Username = '" & ID & "')" & vbNewLine
    elseif Operation = "ADD EXTERNAL USER" then
      sql = ""
      sql = sql & "INSERT Into dbo.PlacementUserAttributes (Username, IsAdmin, IsWriter, IsEssay) VALUES ('" & ID & "', " & CInt(False) & ", " & CInt(False) & ", " & CInt(False) & ")" & vbNewLine
      sql = sql & "INSERT Into dbo.SystemUser (SystemID, Username) VALUES (" &  SYSTEM_ID & ", '" & ID & "')" & vbNewLine    
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
%>
<!DOCTYPE html>
<html lang="en">

<head>
<meta charset="UTF-8" />
<title><% =SYSTEM_NAME %> - Users Editor</title>
<link rel="stylesheet" href="css/placement.css" />
<link rel="stylesheet" href="css/gototop.css" />
<style>
table#users {
	margin: 10px auto;
	border: thin black inset;
	border-collapse: separate;
}
table#users td, table#users th {
	border: thin black inset;
}
table#users .titleRow {
	height: 40px;
}

table#users .headerRow {
	height: 40px;
}

table#users .detailRow {
  height: 35px;
}

table#users .buttonRow {
  text-align: center;
  height: 50px;
}

table#users .selectCol {
  text-align: center;
  width: 25px;	
}

table#users .userNameCol {
  width: 150px;
}

table#users .passWordCol {
  width: 150px;
}

table#users .IsAdminCol {
  text-align: center;
  width: 75px;
}

table#users .IsWriterCol {
  text-align: center;
  width: 75px;
}

table#users .IsEssayCol {
  text-align: center;
  width: 75px;
}
</style>
</head>

<body>
<div class="center">
<%
  Call MakeHeader("Users Editor")
  if ErrMsg <> "" then
    Response.Write "<h3 class=""red"">Error(s) Occurred in Submission<br />" & ErrMsg & "</h3>" & vbNewLine
  end if
%>
<p class="hideElement bold">
<a style="margin-right: 10px" href="users.asp">Users Editor</a>
<a style="margin-right: 10px" href="utilrept.asp">Utilities and Reports</a>
<a style="margin-right: 10px" href="default.asp?ResetQuery=true">Main Menu</a>
<a href="default.asp?Logout=true">Log Out</a>
</p>
<hr />

<div style="width: 80%; margin: auto">
<form id="changeUserForm" method="post" autocomplete="off">
<%
  if Operation = "LIST" then
%>
	<table id="users" >
		<tr class="headerRow">
			<th class="selectCol">&nbsp;</th>
			<th class="userNameCol">Username</th>
			<th class="mustChange">Must Change<br/>Password</th>
			<th class="IsAdminCol">Is Admin</th>
			<th class="IsWriterCol">Is Writer</th>
			<th class="IsEssayCol">Is Essay</th>
		</tr>
<%  
    do while not rs.Eof
      Response.Write "<tr class=""detailRow"">" & vbNewline
      Response.Write "<td class=""selectCol""><input name=""ID"" type=""radio"" value=""" & rs("UserName") & """ /></td>" & vbNewLine
      Response.Write "<td class=""userNameCol"">" & rs("UserName") & "</td>" & vbNewLine
      Response.Write "<td class=""mustChange"">" & rs("mustChangePassword") & "</td>" & vbNewLine
      Response.Write "<td class=""IsAdminCol"">" & rs("IsAdmin") & "</td>" & vbNewLine
      Response.Write "<td class=""IsWriterCol"">" & rs("IsWriter") & "</td>" & vbNewLine
      Response.Write "<td class=""IsEssayCol"">" & rs("IsEssay") & "</td>" & vbNewLine
      Response.Write "</tr>" & vbNewLine
      rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
%>
		<tr class="buttonRow">
			<td colspan="6">
			<input type="submit" style="margin-right: 30px" name="Operation" value="Add" />
			<input type="submit" style="margin-right: 30px" name="Operation" value="Add External User" />
			<input type="submit" style="margin-right: 30px" name="Operation" value="Edit" />
			<input type="submit" style="margin-right: 30px" name="Operation" value="Delete" />
			<input type="reset" name="Reset" /></td>
		</tr>
	</table>
<%
  elseif Operation = "ADD" or Operation = "EDIT" then
%>
	<table id="users" >
		<tr class="titleRow">
		<th colspan="6"><%=Operation%> a User</th>
		</tr>
		<tr class="headerRow">
			<th class="userNameCol">Username</th>
			<th class="passWordCol">Password</th>
			<th class="mustChange">Must Change<br/>Password</th>
			<th class="IsAdminCol">Is Admin</th>
			<th class="IsWriterCol">Is Writer</th>
			<th class="IsEssayCol">Is Essay</th>
		</tr>
		<tr class="detailRow">
			<td class="userNameCol"><input type="text" name="UserName" maxlength="50" size="15" value="<% =Username %>" required="required" /></td>
			<td class="passWordCol"><input type="password" id="Password" maxlength="50" size="15" value="<% =Password %>" required="required" autocomplete="off" /></td>
			<td class="mustChange"><input type="checkbox" name="MustChange" value="YES" <%=iif(MustChange, "checked=""checked""", "")%> /></td>
			<td class="IsAdminCol"><input type="checkbox" name="IsAdmin" value="YES" <%=iif(IsAdmin, "checked=""checked""", "")%> /></td>
			<td class="IsWriterCol"><input type="checkbox" name="IsWriter" value="YES" <%=iif(IsWriter, "checked=""checked""", "")%> /></td>
			<td class="IsEssayCol"><input type="checkbox" name="IsEssay" value="YES" <%=iif(IsEssay, "checked=""checked""", "")%> /></td>
		</tr>
		<tr class="buttonRow">
			<td colspan="6">
			<input type="submit" style="margin-right: 30px" name="Operation" value="Submit" />
			<input type="submit" style="margin-right: 30px" name="Operation" value="Cancel" formnovalidate="formnovalidate" />
			<input type="reset" name="Reset" />
			<input type="hidden" id="md5Password" name="md5Password" />
			<input type="hidden" name="SubOperation" value="<%=Operation%>" />
			<input type="hidden" name="ID" value="<%=ID%>" />
			</td>
		</tr>
	</table>
<%
  elseif Operation = "ADD EXTERNAL USER" then
%>
	<table id="users" >
		<tr class="titleRow">
		<th colspan="2">Add a User From Another<br/>Testing Center System</th>
		</tr>
		<tr class="headerRow">
			<th class="selectCol">&nbsp;</th>
			<th class="userNameCol">Username</th>
		</tr>
<%  
    do while not rs.Eof
      Response.Write "<tr class=""detailRow"">" & vbNewline
      Response.Write "<td class=""selectCol""><input name=""ID"" type=""radio"" value=""" & rs("UserName") & """ /></td>" & vbNewLine
      Response.Write "<td class=""userNameCol"">" & rs("UserName") & "</td>" & vbNewLine
      Response.Write "</tr>" & vbNewLine
      rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
%>
		<tr class="buttonRow">
			<td colspan="2">
			<input type="submit" style="margin-right: 30px" name="Operation" value="Submit" />
			<input type="submit" name="Operation" value="Cancel" />
			<input type="hidden" name="SubOperation" value="<%=Operation%>" />
			</td>
		</tr>
	</table>

<%
  elseif Operation = "DELETE" then
%>
	<table id="users" >
		<tr class="titleRow">
		  <th colspan="6">Delete This User?</th>
		</tr>
		<tr class="headerRow">
		  <th class="userNameCol">Username</th>
		  <th class="mustChange">Must Change<br/>Password</th>
		  <th class="IsAdminCol">Is Admin</th>
		  <th class="IsWriterCol">Is Writer</th>
		  <th class="IsEssayCol">Is Essay</th>
		</tr>
		<tr class="detailRow">
		  <td class="userNameCol"><% =Username %></td>
		  <td class="mustChange"><% =MustChange %></td>
		  <td class="IsAdminCol"><% =IsAdmin %></td>
		  <td class="IsWriterCol"><% =IsWriter %></td>
		  <td class="IsEssayCol"><% =IsEssay %></td>
		</tr>
		<tr class="buttonRow">
		  <td colspan="5">
			<input type="submit" style="margin-right: 30px" name="Operation" value="YES" />
			<input type="submit" name="Operation" value="Cancel" />
			<input type="hidden" name="ID" value="<%=ID%>" />
		  </td>
		</tr>
	</table>
<%
  end if
%>
</form>
</div>
</div>
<script src="//ajax.googleapis.com/ajax/libs/jquery/2.2.4/jquery.min.js"></script> 
<script src="js/crypto-js.min.js"></script>
<script src="js/jquery.gototop.js"></script> 
<script>
  $(function() {
    "use strict";
    
    function passwordVerify(str) {
      var i, ch, count, ucase, lcase, digit, punct, illegal;
      if (str.length < 8) {
        return "Password is less than 8 characters";
        }
      for (i = 0; i < str.length; i += 1) {
        ch = str.charAt(i);
        if ("ABCDEFGHIJKLMNOPQRSTUVWXYZ".indexOf(ch) !== -1) {
          ucase = true;
          }
        else if ("abcdefghijklmnopqrstuvwxyz".indexOf(ch) !== -1) {
          lcase = true;
          }
        else if ("0123456789".indexOf(ch) !== -1) {
          digit = true;
          }
        else if ("{}[],.<>;:'\"?/|\\`~!@#$%^&*()_-+=".indexOf(ch) !== -1) {
          punct = true;
          }
        else {
          illegal = true;
          break;
          }
        }
      if (illegal) {
         return "Illegal Character Encountered";
         }
      count = 0;
      if (ucase) {
        count += 1;
        }
      if (lcase) {
        count += 1;
        }
      if (digit) {
        count += 1;
        }
      if (punct) {
        count += 1;
        }
      if (count < 3) {
        return "Password must contain at least 3 of the 4 character groups";
        }
      return "";      
    }
    
    $("#toTop").gototop({ container: "body" });
    $("#changeUserForm").on("submit", function() {
      var password;
      var md5Password;
      var passwordState;
      var $password = $("#Password");
      var btn = $(this).find("input[type=submit]:focus" );
    
      if (btn.val() === "Cancel") {
          return true;
          }
          
      if ($password.length > 0) {
        password = $.trim($("#Password").val());
        if (password !== $password.prop("defaultValue")) { 
          passwordState = passwordVerify(password);
          if (passwordState !== "") {
            window.alert("ILLEGAL PASSWORD: " + passwordState);
            return false;
            }
          md5Password = CryptoJS.MD5(password).toString().toUpperCase();
          $("#md5Password").val(md5Password);
          }
        else {
          $("#md5Password").val(password);
          }
        }
      
      return true;
    });
    
  });
</script>
</body>
</html>
<%
  closeConnection(conn)
%>