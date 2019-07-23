<% 
  Option Explicit
  On Error Resume Next
%>
<!-- #include file="library/library.asp" -->
<%
  Dim username, oldpassword, newpassword, verifypassword, ErrMsg, sqlErr
  Dim conn, rs, sql
  Dim newpw, verpw
    
  if IsEmpty(Session("User")) then
    Response.Redirect "login.asp"
  end if

  Set conn = openConnection(Application("UserDirectoryConnectionString"))

  ErrMsg = ""
  if Request.ServerVariables("REQUEST_METHOD") = "POST" then
    newpw = False
    verpw = False
    username = Trim(Request("username"))
    oldpassword = Trim(Request("md5OldPassword"))
    newpassword = Trim(Request("md5NewPassword"))
    verifypassword = Trim(Request("md5VerifyPassword"))
    if username = "" then
      ErrMsg = BuildErrMsg(ErrMsg, "Username is empty.")
    end if
    if oldpassword = BLANK_PASSWORD then
      ErrMsg = BuildErrMsg(ErrMsg, "Old Password is empty.")
    end if
    if newpassword = BLANK_PASSWORD then
      ErrMsg = BuildErrMsg(ErrMsg, "New Password is empty.")
    else
      newpw = true
    end if
    if verifypassword = BLANK_PASSWORD then
      ErrMsg = BuildErrMsg(ErrMsg, "Verify Password is empty.")
    else
      verpw = true
    end if
    if newpw and verpw then
      if newpassword <> verifypassword then
        ErrMsg = BuildErrMsg(ErrMsg, "New Password and Verification do not match.")
      elseif newpassword = oldpassword then
        ErrMsg = BuildErrMsg(ErrMsg, "New Password Is The Same As The Old One. Come on you can't fool me!")
      end if
    end if
    if ErrMsg = "" then
      sql = "SELECT Username FROM dbo.Users WHERE username = '" & username & "' AND password = '" & oldpassword & "'"
      sqlErr = ExecuteSQLForRs(conn, sql, rs)
      ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
      
      if ErrMsg = "" then
        if rs.Eof then
          ErrMsg = BuildErrMsg(ErrMsg, "Username/Password Not Found")
        else
          sql = "UPDATE dbo.Users SET PASSWORD = '" & newpassword & "', MustChangePassword = 0 WHERE Username = '" & rs("Username") & "'"
          sqlErr = ExecuteSQL(conn, sql)
          ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
          Session("User")("mustChangePassword") = Empty
        end if
        rs.Close
        Set rs = Nothing
      end if
    end if
  else
    username = Session("User")("username")
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
	margin-left: auto;
	margin-right: auto;
	text-align: left;
	font-weight: bold;
}
table#pass input[type=text], table#pass input[type=password] {
	font-family: 'Courier New', Courier, monospace;
}
table#info {
	width: 510px;
}
table#info td {
	width: 50%;
}
</style>
</head>

<body>

<div class="center">
	<%
  Call MakeHeader("Change Password")
  if ErrMsg <> "" then
    Response.Write "<h3 class=""red"">Error(s) Occurred in Submission<br />" & ErrMsg & "</h3>" & vbNewLine
  elseif Request.ServerVariables("REQUEST_METHOD") = "POST" then
    Response.Write "<h3 class=""green"">Password changed!!!</h3>" & vbNewLine
  end if
%>
	<p class="bold">
	<a href="default.asp?ResetQuery=true" style="margin-right: 10px">Main Menu</a>
	<a href="default.asp?Logout=true">Log Out</a></p>
	<hr /><%
	  if Session("User")("mustChangePassword") then
	    Response.Write "<h2>You Must Change Your Password!!!</h2>"
	  end if
	%>
	<div style="margin:auto; width: 80%">
	<p style="text-align: left">This page allows you to change the password 
	on accounts established on the STCC Placement Record Searching System.&nbsp; 
	You are free to change your password from anywhere you have access to a web 
	browser.&nbsp; For security reasons, you are encouraged to change your password 
	at regular intervals.</p></div>
	<form id="changePasswordForm" method="post">
		<table id="pass">
			<tr>
				<td>User Name:</td>
				<td>
				<input name="username" size="20" type="text" value="<%=username%>" readonly="readonly" required="required" /></td>
			</tr>
			<tr>
				<td>Old Password:</td>
				<td>
				<input id="oldPassword" size="20" type="password" required="required" /></td>
			</tr>
			<tr>
				<td>New Password:</td>
				<td>
				<input id="newPassword" size="20" type="password" required="required" /></td>
			</tr>
			<tr>
				<td>Verification:</td>
				<td>
				<input id="verifyPassword" size="20" type="password" required="required" /></td>
			</tr>
			<tr style="text-align: center; height: 50px">
				<td colspan="2">
				<input type="hidden" id="md5OldPassword" name="md5OldPassword" />
				<input type="hidden" id="md5NewPassword" name="md5NewPassword" />
				<input type="hidden" id="md5VerifyPassword" name="md5VerifyPassword" />
				<input name="changeit" type="submit" value="Change Password" /></td>
			</tr>
		</table>
	</form>
	<hr />
	<h1>Valid Secure Password Characteristics</h1>
	<p>Valid Accounts and Passwords must have the following characteristics:</p>
	<div style="text-align: left; width: 600px; margin-left: auto; margin-right: auto">
		<ul>
			<li>Are Case Sensitive. </li>
			<li>Be at least 8 characters long </li>
			<li>May not contain user account name, or any portion of the user’s 
			full name </li>
			<li>May not be reused or reset to a previous password</li>
			<li>Your password must contain characters from at least 3 of the following 
			4 classes:</li>
		</ul>
		<blockquote>
			<table id="info" border="1" cellpadding="7" cellspacing="1">
				<tr>
					<th>Description</th>
					<th>Examples</th>
				</tr>
				<tr>
					<td>1. Upper Case Letters</td>
					<td><pre>A, B, C, … Z</pre></td>
				</tr>
				<tr>
					<td>2. Lower Case Letters</td>
					<td>
					<pre>a, b, c, … z</pre>
					</td>
				</tr>
				<tr>
					<td>3. Digits</td>
					<td>
					<pre>0, 1, 2, … 9</pre>
					</td>
				</tr>
				<tr>
					<td>4. Non-alphanumeric </td>
					<td>For example, punctuation, symbols.
					<pre>({}[],.&lt;&gt;;:&#39;&quot;?/|\`~!@#$%^&amp;()_=)</pre>
					</td>
				</tr>
			</table>
		</blockquote>
		<ul>
			<li>Your password should not be a &quot;common&quot; word (for example, it should 
			not be a word in the dictionary or slang in common use). Your password 
			should not contain words from any language, because numerous password-cracking 
			programs exist that can run through millions of possible word combinations 
			in seconds.</li>
			<li>A complex password that cannot be broken is useless if you cannot 
			remember it. For security to function, you must choose a password you 
			can remember and yet is complex. For example, Msi5!YOld (My Son is 5 
			years old) OR IhliCf5#yN (I have lived in California for 5 years now).</li>
		</ul>
	</div>
</div>
<script src="//ajax.googleapis.com/ajax/libs/jquery/2.2.4/jquery.min.js"></script> 
<script src="js/crypto-js.min.js"></script>
<script src="js/jquery.gototop.js"></script> 
<script>
  $(function() {
    "use strict";
    
    var newPassword;
    var md5NewPassword;
    var verifyPassword;
    var md5VerifyPassword;
    var passwordState;
    
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
    $("#changePasswordForm").on("submit", function(e) {
       oldPassword = $.trim($("#oldPassword").val());
       passwordState = passwordVerify(oldPassword);
       if (passwordState !== "") {
         window.alert("ILLEGAL OLD PASSWORD: " + passwordState);
         return false;
         }
       newPassword = $.trim($("#newPassword").val());
       passwordState = passwordVerify(newPassword);
       if (passwordState !== "") {
         window.alert("ILLEGAL NEW PASSWORD: " + passwordState);
         return false;
         }
       verifyPassword = $.trim($("#verifyPassword").val());
       passwordState = passwordVerify(verifyPassword);
       if (passwordState !== "") {
         window.alert("ILLEGAL VERIFY PASSWORD: " + passwordState);
         return false;
         }
      
       md5OldPassword = CryptoJS.MD5(oldPassword).toString().toUpperCase();
       $("#md5OldPassword").val(md5OldPassword);
       md5NewPassword = CryptoJS.MD5(newPassword).toString().toUpperCase();
       $("#md5NewPassword").val(md5NewPassword);
       md5VerifyPassword = CryptoJS.MD5(verifyPassword).toString().toUpperCase();
       $("#md5VerifyPassword").val(md5VerifyPassword);
      
      return true;
    });
    
  });
</script>
</body>
</html>
<%
  closeConnection(conn)
%>