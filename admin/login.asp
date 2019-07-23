<%
  Option Explicit
  On Error Resume Next

  Function ParseDomainFromURL(ByVal url) 
    Dim urlParts
    if instr(url, "//") > 0 then 
      urlParts = split(url,"/") 
      ParseDomainFromURL = urlParts(2) 
    else 
      ParseDomainFromURL = "" 
    end if 
  End Function 
%>
<!-- #include file="library/library.asp" -->
<!-- #include file="library/UserSecurity.asp" -->
<!-- #include file="library/CheckIP.asp" -->
<%
  Dim Username, ErrMsg, IPAddressAllowed
  Dim tmpUsername, tmpPassword

  IPAddressAllowed = true
  if not NorthAmericanVisitor(Request.ServerVariables("REMOTE_ADDR")) then
    IPAddressAllowed = false
  end if

  ErrMsg = ""
  if IPAddressAllowed and Request.ServerVariables("REQUEST_METHOD") = "POST" then
    Username= Request.Form("Username")
    tmpUsername = Trim(Username)
    tmpPassword = Trim(Request("MD5Password"))
    
    if Len(tmpUsername) = 0 then	
      ErrMsg = BuildErrMsg(ErrMsg, "You must enter your User Name.")
    end if
    if tmpPassword = BLANK_PASSWORD then
      ErrMsg = BuildErrMsg(ErrMsg, "You must enter your Password.")
    end if
    if ErrMsg = "" then
      if not signUserOn(tmpUserName, tmpPassword) then
        ErrMsg = BuildErrMsg(ErrMsg, "Invalid User Account/Password")
      else
        Response.Redirect "default.asp"
      end if
    end if   
  end if
%>
<!DOCTYPE html>
<html lang="en">

<head>
<meta charset="UTF-8" />
<title><% =SYSTEM_NAME %></title>
<meta name="robots" content="NOINDEX,NOFOLLOW" />
<link rel="stylesheet" href="css/placement.css" />
<link rel="stylesheet" href="css/gototop.css" />
<style>
#capsWarn {
	display: none;
	text-align:center;
	font-weight:bold;
	color:red;
	padding:0;
	margin:0;
}
table#header {
	text-align: left;
	margin: 0 auto 15px auto;
}
table#header td.firstcol {
	width: 108px;
}
table#header td.secondcol {
	width: 500px;
	text-align: center;
}
table#logon {
	text-align: left;
	margin-left: auto;
	margin-right: auto;
	margin-bottom: 10px;
}
table#logon input[type=text], table#logon input[type=password] {
	font-family: 'Courier New', Courier, monospace;
}
table#logon td.firstcol {
	text-align: right;
	height: 35px;
	font-weight: bold;
}
table#logon td.secondcol {
	width: 5px;
}
table#logon td.thirdcol {
	height: 35px;
}
table#logon td.submit {
	text-align: center;
	height: 35px;
}
</style>
</head>

<body>

<div class="center">
	<table id="header">
		<tr>
			<td class="firstcol"><a href="http://www.stcc.edu/"><img src="img/smsealbw.png" width="108" height="108" alt="STCC Home Page" /></a></td>
			<td class="secondcol"><img src="img/recrdsrch.jpg" width="505" height="54" alt="<% =SYSTEM_NAME %>" /></td>
		</tr>
	</table>
<%
  if not IPAddressAllowed then
%>
    <h2 class="red">Your IP Address is <%=Request.ServerVariables("REMOTE_ADDR")%><br/>
    This system Is not allowed to execute on this machine.</h2>
<%
  else
%>
	<hr />
<%
  if ErrMsg <> "" then
    Response.Write "<h3 class=""red"">" & ErrMsg & "</h3>" & vbNewLine
  end if
%>
    <noscript>
    <h1 class="red">Javascript Needed To Use This System!!!</h1>
    </noscript>
	<h3>Enter Logon Information to use this search system</h3>
	<form id="loginForm" method="post" autocomplete="off">
		<table id="logon">
			<tr>
				<td class="firstcol">User Name: </td>
				<td class="secondcol">&nbsp;</td>
				<td class="thirdcol"><input type="text" id="Username" name="Username" maxlength="20" size="20" value="<%=Username%>" required /></td>
			</tr>
			<tr>
				<td class="firstcol">Password: </td>
				<td class="secondcol">&nbsp;</td>
				<td class="thirdcol"><input type="password" id="Password" maxlength="20" size="20" required autocomplete="off"/></td>
			</tr>
			<tr>
				<td colspan="3"><p id="capsWarn">Is your CAPSLOCK on?</p></td>
			</tr>
			<tr>
				<td class="submit" colspan="3">
				<input type="hidden" id="md5password" name="md5password" />
				<input type="submit" value="Signon" /></td>
			</tr>
		</table>
	</form>
	<hr />
	
	<p><a href="https://www.rapidssl.com/" target="_blank">
	<img alt="RapidSSL Certificate " height="50" src="img/rapidssl_ssl_certificate.gif" width="90" /></a></p>
<%
  end if
%>
</div>

<script src="//ajax.googleapis.com/ajax/libs/jquery/2.2.4/jquery.min.js"></script> 
<script src="js/crypto-js.min.js"></script>
<script src="js/jquery.gototop.js"></script>
<script>
$(function() {
  "use strict";
  $("#Username").focus();
  $("#toTop").gototop({ container: "body" });
  $("#loginForm").on("submit", function() { 
     var md5password = CryptoJS.MD5($.trim($("#Password").val())).toString().toUpperCase();
     $("#md5password").val(md5password);
     return true;
     });
  $("input[type='password']").keypress(function(e) {
    var $warn = $("#capsWarn"); 
    var kc = e.which;
    var isUp = kc >= 65 && kc <= 90;
    var isLow = kc >= 97 && kc <= 122;
    var isShift = e.shiftKey ? e.shiftKey : kc === 16;

    if ((isUp && !isShift) || (isLow && isShift)) {
        $warn.show();
    } else {
        $warn.hide();
    }
  });
});
</script>
</body>

</html>
