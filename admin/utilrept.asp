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
%>
<!DOCTYPE html>
<html lang="en">

<head>
<meta charset="UTF-8" />
<title><% =SYSTEM_NAME %> - Utilities and Reports Page</title>
<link rel="stylesheet" href="css/placement.css" />
<link rel="stylesheet" href="css/gototop.css" />
<style>
#menuparent {
    text-align: center;
	font-weight: bold;
	font-size: 14pt;
	margin-bottom: 10px;
}
</style>
</head>

<body>
<div class="center">
<%
  MakeHeader "Utilities and Reports"
%>
<p class="bold">
<a style="margin-right: 10px" href="utilrept.asp" style="margin-right: 10px">Utilities/Reports</a>
<a style="margin-right: 10px" href="default.asp?ResetQuery=true">Main Menu</a>
<a href="default.asp?Logout=true">Log Out</a></p>
<hr />
	<div id="menuparent">
	<ul id="menu">
	    <li><a href="placementsFromTestScores.asp">Calculate Placements From Test Scores</a></li>
		<li><a href="report.asp">Placement Testing Report</a></li>
		<li><a href="uploadPlacementDataFailure.asp">Upload Placement Data Failure Report</a></li>
		<li><a href="uploadBackgroundFailure.asp">Upload Background Question Failure Report</a></li>
		<li><a href="countsReport.asp">Counts Report</a></li>
		<% if Session("User")("IsUserManager") then %> 
		<li><a href="users.asp">Users Editor</a></li>
		<% end if %>
		<li><a href="accuplacerLogin.asp">Configure Accuplacer Login</a></li>
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
