<% 
  Option Explicit
  On Error Resume Next
%>
<!-- #include file="library/library.asp" -->
<%
  if IsEmpty(Session("User")) then
    Response.Redirect "login.asp"
  elseif not Session("User")("IsAdmin") and not Session("User")("IsEssay") then 
    Response.Redirect "default.asp"
  end if
%>
<!DOCTYPE html>
<html lang="en">

<head>
<meta charset="UTF-8" />
<title><% =SYSTEM_NAME %> - Essay Utilities and Reports</title>
<link rel="stylesheet" href="css/placement.css" />
<link rel="stylesheet" href="css/gototop.css" />
<style>
table.options {
	text-align: left;
	margin-top: 20px;
	margin-left: auto;
	margin-right: auto;
}

table.options li {
	font-weight: bold;
	font-size: 14pt;
	margin-bottom: 10px;
}
.isDisabled {
  cursor: not-allowed;
}
.isDisabled > a {
  color: currentColor;
  display: inline-block;  /* For IE11/ MS Edge bug */
  pointer-events: none;
  text-decoration: none;
  opacity: 0.5;
}
</style>
</head>

<body>
<div class="center">
<%
  MakeHeader "Essay Utilities and Reports"
%>
<p class="hideElement bold">
<a href="EssayUtilities.asp">Essay Utilities</a><span style="margin-right: 10px">&nbsp;</span>
<a href="default.asp?ResetQuery=true">Main Menu</a><span style="margin-right: 10px">&nbsp;</span>
<a href="default.asp?Logout=true">Log Out</a>
</p>
<hr />
	<table class="options" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td>
			<ul id="menu">
				<li><a href="accuplacerEssayReport.asp">Accuplacer Essay Report</a></li>
				<!--
				<li><span class="isDisabled"><a data-href="EssayReaders.asp" href="#">Establish Essay Readers</a></span></li>
				<li><span class="isDisabled"><a data-href="essayCountReport.asp" href="#">Essay Readers Count Report</a></span></li>
				<li><span class="isDisabled"><a data-href="sendReaderEmails.asp" href="#">Send Emails to Readers</a></span></li>
				-->
			</ul>
			</td>
		</tr>
	</table>
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
