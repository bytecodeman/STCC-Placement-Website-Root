<% 
  Option Explicit
  On Error Resume Next
%>
<!-- #include file="library/library.asp" -->
<!-- #include file="library/calculatePlacements.asp" -->
<%
  if IsEmpty(Session("User")) then
    Response.Redirect "login.asp"
  elseif not Session("User")("IsAdmin") then 
    Response.Redirect "default.asp"
  end if
  
  Dim ArithScore, AlgScore, CollegeScore, MathPlacement
  Dim EnglScore, WritePlacerScore, EnglReadingScore, EnglishPlacement
  Dim ReadingScore, ReadingPlacement
  Dim KeyBoardWordsPerMin, KeyBoardErrors, KeyboardPlacement
  Dim ExitScore, ExitPlacement

  if Request.ServerVariables("REQUEST_METHOD") = "POST" then
    Select Case True
      Case not IsEmpty(Request("MathCalculate"))
         MathPlacement = GetMathPlacement(ArithScore, AlgScore, CollegeScore)
      Case not IsEmpty(Request("EnglishCalculate"))
         EnglishPlacement = GetEnglishPlacement(WritePlacerScore)
      Case not IsEmpty(Request("ReadingCalculate"))
         ReadingPlacement = GetReadingPlacement(ReadingScore)
      Case not IsEmpty(Request("KeyboardPlacement"))
         KeyboardPlacement = GetKeyboardPlacement(KeyBoardWordsPerMin, KeyBoardErrors)
      Case not IsEmpty(Request("ExitPlacement")) 
         ExitPlacement = GetExitPlacement(ExitScore)
    End Select
  end if
  
  Function GetMathPlacement(ByRef ArithScore, ByRef AlgScore, ByRef CollegeScore)
    On Error Resume Next
    Dim tmpArith, tmpAlg, tmpCollege, Placement
    ArithScore = Trim(Request("ArithScore"))
    AlgScore = Trim(Request("AlgScore"))
    CollegeScore = Trim(Request("CollegeScore"))
    if ArithScore & AlgScore & CollegeScore = "" then
      Placement = "ERROR"
    else
      tmpArith = ArithScore
      tmpAlg = AlgScore
      tmpCollege = CollegeScore
      if tmpArith = "" then tmpArith = "0"
      if tmpAlg = "" then tmpAlg = "0"
      if tmpCollege = "" then tmpCollege = "0"
      if IsNumeric(tmpArith) and IsNumeric(tmpAlg) and IsNumeric(tmpCollege) then
        Placement = CalculateMathPlacement(CDbl(tmpArith), CDbl(tmpAlg), CDbl(tmpCollege))
      else
        Placement = "ERROR"
      end if
    end if
    if Err.Number <> 0 then
      Placement = "ERROR"
    end if
    GetMathPlacement = Placement
  End Function
  
  Function GetEnglishPlacement(ByRef WritePlacerScore)
    On Error Resume Next
    Dim tmpEngl, tmpEnglReading, tmpWPScore, Placement
    WritePlacerScore = Trim(Request("WritePlacerScore"))
    if EnglScore & EnglReadingScore & WritePlacerScore = "" then
      Placement = "ERROR"
    else
      tmpEngl = EnglScore
      tmpEnglReading = EnglReadingScore
      tmpWPScore = WritePlacerScore
      if tmpEngl = "" then tmpEngl = "0"
      if tmpEnglReading = "" then tmpEnglReading = "0"
      if tmpWPScore = "" then tmpWPScore = "-2" 
      if IsNumeric(tmpEngl) and IsNumeric(tmpEnglReading) and IsNumeric(tmpWPScore) then
        Placement = CalculateEnglishPlacement(CDbl(tmpEngl), CDbl(tmpEnglReading), CInt(tmpWPScore))
      else
        Placement = "ERROR"
      end if
    end if
    if Err.Number <> 0 then
      Placement = "ERROR"
    end if
    GetEnglishPlacement = Placement
  End Function 
  
  Function GetReadingPlacement(ByRef ReadingScore)
    On Error Resume Next
    Dim tmpRead, Placement
    ReadingScore = Trim(Request("ReadingScore"))
    if ReadingScore = "" then
      Placement = "ERROR"
    else
      tmpRead = ReadingScore
      if tmpRead = "" then tmpRead = "0"
      if IsNumeric(tmpRead) then
        Placement = CalculateReadingPlacement(CDbl(tmpRead))
      else
        Placement = "ERROR"
      end if
    end if
    if Err.Number <> 0 then
      Placement = "ERROR"
    end if
    GetReadingPlacement = Placement
  End Function 

  Function GetKeyboardPlacement(ByRef KeyBoardWordsPerMin, ByRef KeyBoardErrors)
    On Error Resume Next
    Dim tmpWordsPerMin, tmpErrors, Placement
    KeyBoardWordsPerMin = Trim(Request("KeyBoardWordsPerMin"))
    KeyBoardErrors = Trim(Request("KeyBoardErrors"))
    if KeyBoardWordsPerMin & KeyBoardErrors = "" then
      Placement = "ERROR"
    else
      tmpWordsPerMin = KeyBoardWordsPerMin
      tmpErrors = KeyBoardErrors
      if tmpWordsPerMin = "" then tmpWordsPerMin = "0"
      if tmpErrors = "" then tmpErrors = "0"
      if IsNumeric(tmpWordsPerMin) and IsNumeric(tmpErrors) then
        Placement = CalculateTypingPlacement(CDbl(tmpWordsPerMin), CDbl(tmpErrors))
      else
        Placement = "ERROR"
      end if
    end if
    if Err.Number <> 0 then
      Placement = "ERROR"
    end if
    GetKeyboardPlacement = Placement
  End Function 

  Function GetExitPlacement(ByRef ExitScore)
    On Error Resume Next
    Dim tmpExit, Placement
    ExitScore = Trim(Request("ExitScore"))
    if ExitScore = "" then
      Placement = "ERROR"
    else
      tmpExit = ExitScore
      if tmpExit = "" then tmpExit = "0"
      if IsNumeric(tmpExit) then
        Placement = CalculateExitPlacement(CDbl(tmpExit))
      else
        Placement = "ERROR"
      end if
    end if
    if Err.Number <> 0 then
      Placement = "ERROR"
    end if
    GetExitPlacement = Placement
  End Function 
%>
<!DOCTYPE html>
<html lang="en">

<head>
<meta charset="UTF-8" />
<title><% =SYSTEM_NAME %> - Utilities and Reports Page</title>
<link rel="stylesheet" href="css/placement.css" />
<link rel="stylesheet" href="css/gototop.css" />
<style>
table {
	text-align: left;
	margin-right: auto;
	margin-left: auto;
    border-style: solid;
    border-width: thin;
	border-color: #ccc black black #ccc;
	border-spacing: 2px;
	border-collapse:separate
}
td, th  {
    border-style: solid;
    border-width: thin;
	border-color: black #ccc #ccc black;
	text-align: center;
	padding: 4px;
}
hr {
	margin-top: 15px;
}
body {
	margin-bottom: 50px;
}
table#options {
	margin-top: 10px;
	border: none;
}
table#options td {
	border: none;
	text-align: left;
	vertical-align: top;
}
</style>
</head>

<body>

<div class="center">
	<%
      MakeHeader "Calculate Placements From Test Scores"
    %>
	<p class="bold">
	<a href="placementsFromTestScores.asp" style="margin-right: 10px">Calculate Placements From Test Scores</a>
    <a href="utilrept.asp" style="margin-right: 10px">Utilities/Reports</a>
	<a href="default.asp?ResetQuery=true" style="margin-right: 10px">Main Menu</a>
	<a href="default.asp?Logout=true">Log Out</a></p>
	<hr />
	<table id="options" border="0" cellpadding="2" cellspacing="1">
		<tr>
			<td>
			<ul>
				<li><a href="#math">Math Placement Calculation</a></li>
				<li><a href="#english">English Placement Calculation</a></li>
				<li><a href="#reading">Reading Placement Calculation</a></li>
			</ul>
			</td>
			<td>
			<ul>
				<li><a href="#keyboard">Keyboarding Placement Calculation</a></li>
				<li><a href="#exit">Reading Exit Placement Calculation</a></li>
			</ul>
			</td>
		</tr>
	</table>
	<hr />
	<h3 id="math">Math Placement Calculation</h3>
	<form action="#math" method="post">
		<div>
			<table>
				<tr>
					<th colspan="3">Scores</th>
					<th rowspan="2">Placement</th>
				</tr>
				<tr>
					<th>Arithmetic</th>
					<th>Algebra</th>
					<th>College</th>
				</tr>
				<tr>
					<td><input type="text" name="ArithScore" size="5" maxlength="5" value="<% =ArithScore%>"/></td>
					<td><input type="text" name="AlgScore" size="5" maxlength="5" value="<% =AlgScore%>" /></td>
					<td><input type="text" name="CollegeScore" size="5" maxlength="5" value="<%=CollegeScore%>" /></td>
					<td><input type="text" name="MathPlacement" size="10" readonly="readonly" value="<%=MathPlacement%>" /></td>
				</tr>
				<tr>
					<td colspan="4">
					<input name="MathCalculate" style="margin-right: 25px" type="submit" value="Calculate" /><input name="Reset1" type="reset" /></td>
				</tr>
			</table>
		</div>
	</form>
	<hr />
	<h3 id="english" style="margin-bottom: 0">English Placement Calculation</h3>
	<form action="#english" method="post">
		<div>
			<table>
				<tr>
					<th>WritePlacer Score</th>
					<th>Placement</th>
				</tr>
				<tr>
					<td><input type="text" name="WritePlacerScore" size="2" maxlength="2" value="<% =WritePlacerScore%>" /></td>
					<td><input type="text" name="EnglishPlacement" size="10" readonly="readonly" value="<% =EnglishPlacement%>"  /></td>
				</tr>
				<tr>
					<td colspan="2">
					<input name="EnglishCalculate" style="margin-right: 25px" type="submit" value="Calculate" /><input name="Reset2" type="reset" /></td>
				</tr>
			</table>
		</div>
	</form>
	<hr />
	<h3 id="reading">Reading Placement Calculation</h3>
	<form action="#reading" method="post">
		<div>
			<table>
				<tr>
					<th>Score</th>
					<th>Placement</th>
				</tr>
				<tr>
					<td><input type="text" name="ReadingScore" size="5" maxlength="5" value="<% =ReadingScore%>" /></td>
					<td><input type="text" name="ReadingPlacement" size="10" readonly="readonly" value="<% =ReadingPlacement%>" /></td>
				</tr>
				<tr>
					<td colspan="2">
					<input name="ReadingCalculate" style="margin-right: 25px" type="submit" value="Calculate" /><input name="Reset3" type="reset" /></td>
				</tr>
			</table>
		</div>
	</form>
	<hr />
	<h3 id="keyboard">Keyboarding Placement Calculation</h3>
	<form action="" method="post">
		<div>
			<table>
				<tr>
					<th>Words/Min</th>
					<th>Errors</th>
					<th>Placement</th>
				</tr>
				<tr>
					<td><input type="text" name="KeyBoardWordsPerMin" maxlength="5" size="5" value="<%=KeyBoardWordsPerMin%>" /></td>
					<td><input type="text" name="KeyBoardErrors" maxlength="5" size="5" value="<% =KeyBoardErrors%>" /></td>
					<td><input type="text" name="KeyboardPlacement" size="10" readonly="readonly" value="<% =KeyboardPlacement%>" /></td>
				</tr>
				<tr>
					<td colspan="3">
					<input name="KeyboardCalculate" style="margin-right: 25px" type="submit" value="Calculate" /><input name="Reset4" type="reset" /></td>
				</tr>
			</table>
		</div>
	</form>
	<hr />
	<h3 id="exit">Reading Exit Placement Calculation</h3>
	<form action="#exit" method="post">
		<div>
			<table>
				<tr>
					<th>Score</th>
					<th>Placement</th>
				</tr>
				<tr>
					<td><input type="text" name="ExitScore" maxlength="5" size="5" value="<% =ExitScore%>" /></td>
					<td><input type="text" name="ExitPlacement" size="10" readonly="readonly" value="<% =ExitPlacement%>" /></td>
				</tr>
				<tr>
					<td colspan="2">
					<input name="ExitCalculate" style="margin-right: 25px" type="submit" value="Calculate" /><input name="Reset" type="reset" /></td>
				</tr>
			</table>
		</div>
	</form>
</div>
<script src="//ajax.googleapis.com/ajax/libs/jquery/2.2.4/jquery.min.js"></script>
<script src="js/jquery.gototop.js"></script>
<script>
$(function(){
  $("#toTop").gototop({ container: "body" });
  
  $("input[type=reset").click(function() {
     $("input[type=text]", $(this).parents("form")).val("");
     return false;
  });
});
</script>
</body>

</html>
