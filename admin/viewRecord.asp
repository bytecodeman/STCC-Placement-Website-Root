<%
  Option Explicit
  On Error Resume Next
%>
<!-- #include file="library/library.asp" -->
<%
  Const ESSAYNOTTAKEN = -2
  
  Dim conn, rs, sql, cmd, pm, tempstr, SSN, StudentID, ErrMsg, sqlErr
  Dim englscore, wpscore

  if IsEmpty(Session("User")) then
    Response.Redirect "login.asp"
  elseif Request("value") = "" then
    Response.Redirect "default.asp?ResetQuery=true"
  else
    StudentID = Request("value")
  end if

  Set conn = openConnection(Application("ConnectionString"))

  ErrMsg = ""
  SSN = TranslateID2SSN(conn, StudentID)
  if SSN = "" then
    ErrMsg = BuildErrMsg(ErrMsg, "Record Cannot Be Located")
  else
    sql = "dbo.StudentPlacementTestingRecord"
    
    Set cmd = Server.CreateObject("ADODB.Command")
    With cmd
      .ActiveConnection = conn
      .CommandText = sql
      .commandType = adCmdStoredProc
      Set pm = .CreateParameter("SSN", adVarChar, adParamInput, 9, SSN)
      .Parameters.Append(pm)
    End With
    Set rs = Server.CreateObject("ADODB.RecordSet")
    With rs
      .CursorType = adOpenForwardOnly
      .CursorLocation = adUseClient
      .Open cmd
      .ActiveConnection = Nothing
    End With
    if Err <> 0 then
      ErrMsg = BuildErrMsg(ErrMsg, "Error No: " & Err.Number & " " & Err.Description)
    end if
  end if
  
  Function Format(ByVal d)
    if IsNull(d) then
      Format = "&nbsp;"
      Exit Function
    end if
    Dim strTime, arr, ampm
    strTime = CStr(FormatDateTime(d,vbLongTime))
    arr = Split(strTime, ":")
    ampm = Split(arr(2), " ")(1)
    Format = FormatDateTime(d,vbShortDate) & " " & arr(0) & ":" & arr(1) & " " & ampm
  End Function
  
  Function MathCourse(ByVal rs)
  	Dim Placement
	Placement = iif(IsNull(rs("MathPlacement")), "", rs("MathPlacement"))
    Select Case Placement
      Case ""
        MathCourse = "NOT TAKEN"
      Case "ARTH071", "MAT071"
        Dim arith 
        arith = CDbl(rs("MathArithScore"))
        If arith <= 29 Then
          MathCourse = "MAT079"
        Else
          MathCourse = "MAT078"
        End If
      Case "ARTH071U", "MAT071U"
        MathCourse = "MAT089 or MAT078"
      Case "ALGB081", "MAT081"
        MathCourse = "MAT087"
      Case "ALGB081U", "MAT081U"
        MathCourse = "MAT099 or MAT087"
      Case "ALGB091", "MAT091"
        MathCourse = "MAT097, MAT101, or MAT115"
      Case "MATH101", "MAT101"
        MathCourse = "Any Math Course requiring Algebra 2"
      Case "MATH105", "MAT105"
        MathCourse = "MAT130 or any Math Course requiring Algebra 2"
      Case "MATH155", "MAT131"
        MathCourse = "MAT131 or any Math Course requiring Algebra 2"
      Case Else
        MathCourse = "Unknown Placement: " & Placement
    End Select
  End Function
  
  Function EnglishCourse(ByVal rs)
  	Dim Placement, wpscore, englscore
	Placement = iif(IsNull(rs("EnglPlacement")), "", rs("EnglPlacement"))
	Select Case Placement
	  Case ""
	    EnglishCourse = "NOT TAKEN"
	  Case "ENG101H", "ENGL110"
        EnglishCourse = "ENG101H or ENG101"
	  Case "ENG101", "ENGL100"
        EnglishCourse = "ENG101"
	  Case "DWT099U"
	    EnglishCourse = "DWT099U, Take ENG-101 and corequisite PHL-120 with matching section no."
	  Case "DWT099C"
		EnglishCourse = "DWT099C, Take DWT-099 with C in section no."
	  Case "DWT099"
		EnglishCourse = "DWT099"
	  Case "ESSAY"
		EnglishCourse = "Essay Not Yet Scored"
	  Case Else
        EnglishCourse = "Unknown Placement: " & Placement
    End Select
  End Function
	    
  Function ReadingCourse(ByVal rs)
  	Dim Placement
	Placement = iif(IsNull(rs("ReadPlacement")), "", rs("ReadPlacement"))
    Select Case Placement
      Case ""
        ReadingCourse = "NOT TAKEN"
      Case "DRG091", "DRDG091"
        ReadingCourse = "DRG091"
      Case "DRG092", "DRDG092"
        ReadingCourse = "DRG092"
      Case "READ105"
        ReadingCourse = "EXEMPT From Reading Course"
      Case Else
        ReadingCourse = "Unknown Placement: " & Placement
    End Select
  End Function

  Function TypingCourse(ByVal rs)
  	Dim Placement
	Placement = iif(IsNull(rs("TypePlacement")), "", rs("TypePlacement"))
    Select Case Placement
      Case ""
        TypingCourse = "NOT TAKEN"
      Case "OFFS100", "OIT100"
        TypingCourse = "OIT100"
      Case "OFFS110", "OIT110"
        TypingCourse = "EXEMPT From Typing Course"
      Case Else
        TypingCourse = "Unknown Placement: " & Placement
    End Select
  End Function
  
  Function WritePlacerScore(wp)
    If IsNull(wp) Then
      WritePlacerScore = ESSAYNOTTAKEN
    Else
      WritePlacerScore = CInt(wp)
    End if
  End Function

%>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<title><% =SYSTEM_NAME %> - View Record</title>
<link rel="stylesheet" href="css/placement.css" />
<link rel="stylesheet" href="css/gototop.css" />
<style>
#resultsHeader {
	text-align: left;
	margin-left: auto;
	margin-right: auto;
}
#results {
	margin-top: 15px;
	margin-left: auto;
	margin-right: auto;
}
#results TD {
	text-align: left;
	vertical-align: top;
	font-size: 12pt;
	font-weight: bold;
}
#placementInfo {
    font-weight: bold;
    margin: auto;
    text-align: left;
}
#placementInfo td {
	vertical-align:top;
}
#scheduleInfo {
    font-weight: bold;
    margin: auto;
    text-align: left;
    width: 600px;
}
#scheduleInfo td:first-child {
	text-align:right;
	width: 100px;
}
</style>
</head>

<body>

<div class="center">
	<%
  Call MakeHeader("View Record")
  if ErrMsg <> "" then
    Response.Write "<h3 class=""red"">Error(s) Occurred in Viewing Record<br />" & ErrMsg & "</h3>" & vbNewLine
  end if
%>
	<p class="hideElement bold"><% if Session("User")("IsWriter") then %>
	<a href="addRecord.asp" style="margin-right: 10px">Add Record</a>
	<a href="viewRecord.asp?value=<%=StudentID%>" style="margin-right: 10px">View 
	Record</a>
	<a href="editRecord.asp?value=<%=StudentID%>" style="margin-right: 10px">Edit 
	Record</a>
	<a href="testdetails.asp?value=<%=StudentID%>" style="margin-right: 10px">Edit 
	Test Details</a>
	<a href="delRecord.asp?value=<%=StudentID%>" style="margin-right: 10px">Delete 
	Record</a> <% end if %><a href="default.asp" style="margin-right: 10px">Query 
	List</a> <a href="default.asp?RefineSearch=true" style="margin-right: 10px">
	Refine Search</a>
	<a href="default.asp?ResetQuery=true" style="margin-right: 10px">New Query</a>
	<a href="default.asp?Logout=true">Log Out</a> </p>
	<hr class="hideElement" /><%
  if ErrMsg = "" then
    if not rs.eof then
%>
	<h1>Placement Testing Report</h1>
	<table id="resultsHeader" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td style="width: 108px">
			<img alt="STCC Home Page" height="108" src="img/smsealbw.png" width="108" /></td>
			<td style="width: 70px">&nbsp;</td>
			<td style="width: 250px" valign="top"><%
          Response.Write "<p style=""font-size: 12pt; font-weight: bold; margin-top: 15px"">" & vbNewLine
          Response.Write rs("LastName") & ", " & rs("FirstName") & "<br />" & vbNewLine
          Response.Write rs("Street") & "<br />" & vbNewLine
          Response.Write rs("City") & ", " & rs("State") & "&nbsp;&nbsp;" & rs("ZipCode") & vbNewLine
          Response.Write "</p>" & vbNewLine
        %></td>
		</tr>
	</table>
	<table id="results" border="0" cellpadding="5" cellspacing="5">
		<tr>
			<td>Test</td>
			<td>Placement</td>
			<td>Date</td>
			<td>Raw Scores</td>
		</tr>
		<tr>
			<td>Math:</td>
			<td><%=iif(IsNull(rs("MathPlacement")), "NOT TAKEN", rs("MathPlacement"))%>
			</td>
			<td><%=iif(IsNull(rs("MathPlacement")), "&nbsp;", Format(rs("MathDate")))%>
			</td>
			<td><%
       if IsNull(rs("MathPlacement")) then
         tempstr = "&nbsp;"
       else
         tempstr = Int(Cdbl(rs("MathArithScore"))) & "-" & Int(CDbl(rs("MathAlgScore"))) & "-" & Int(CDbl(rs("MathCollegeScore")))
       end if
       Response.Write tempstr
     %></td>
		</tr>
		<tr>
			<td>English:</td>
			<td><%
		if IsNull(rs("EnglPlacement")) then
		  tempstr = "NOT TAKEN"
		elseif rs("EnglPlacement") = "ESSAY" then
		  tempstr = ""
		  if IsNull(rs("EssayID")) then
		    tempstr = "NEEDS "
		  end if
		  tempstr = tempstr & "ESSAY"
		else
		  tempstr = rs("EnglPlacement")
		end if
		Response.Write tempstr
		%></td>
			<td><%=iif(IsNull(rs("EnglPlacement")), "&nbsp;", Format(rs("EnglDate")))%>
			</td>
			<td><%
       if IsNull(rs("EnglPlacement")) then
         tempstr = "&nbsp;"
       else
         englscore = Int(CDbl(rs("EnglScore")))
         if englscore > 0 then
           tempstr = englscore
           wpscore = WritePlacerScore(rs("WPScore"))
           if wpscore >= 0 then
             tempstr = tempstr & "-" & wpscore
           end if
         else
           wpscore = WritePlacerScore(rs("WPScore"))
           tempstr = wpscore
         end if
       end if
       Response.Write tempstr
     %></td>
		</tr>
		<tr>
			<td>Reading:</td>
			<td><%=iif(IsNull(rs("ReadPlacement")), "NOT TAKEN", rs("ReadPlacement"))%>
			</td>
			<td><%=iif(IsNull(rs("ReadPlacement")), "&nbsp;", Format(rs("ReadDate")))%>
			</td>
			<td><%
       if IsNull(rs("ReadPlacement")) then
         tempstr = "&nbsp;"
       else
         tempstr = Int(CDbl(rs("ReadScore")))
       end if
       Response.Write tempstr
     %></td>
		</tr>
		<tr>
			<td>Keyboarding:</td>
			<td><%=iif(IsNull(rs("TypePlacement")), "NOT TAKEN", rs("TypePlacement"))%>
			</td>
			<td><%=iif(IsNull(rs("TypePlacement")), "&nbsp;", Format(rs("TypeDate")))%>
			</td>
			<td><%
       if IsNull(rs("TypePlacement")) then
         tempstr = "&nbsp;"
       else
         tempstr = rs("PassageNo") & "-" & rs("WordsPerMin") & "-" & rs("Errors")
       end if
       Response.Write tempstr
     %></td>
		</tr>
		<tr>
			<td>Reading Exit:</td>
			<td><%=iif(IsNull(rs("ExitPlacement")), "NOT TAKEN", rs("ExitPlacement"))%>
			</td>
			<td><%=iif(IsNull(rs("ExitPlacement")), "&nbsp;", Format(rs("ExitDate")))%>
			</td>
			<td><% 
       if IsNull(rs("ExitPlacement")) then
         tempstr = "&nbsp;"
       else
         tempstr = Int(CDbl(rs("ExitScore")))
       end if
       Response.Write tempstr
     %></td>
		</tr>
	</table>
	<h2>Course Placement Information</h2>
	<table id="placementInfo" border="0" cellpadding="0" cellspacing="20">
		<tr>
			<td>
			MAT071, PreAlgebra<br />
			MAT071U, Elementary Algebra 1 (5 Day)<br />
			MAT081, Elementary Algebra 1<br />
			MAT081U, Elementary Algebra 2 (5 Day)<br />
			MAT091, Elementary Algebra 2<br />
			MAT101, College Level 1<br />
			MAT105, College Level 2<br />
			MAT131, Calculus 1 </td>
			<td>
			DWT099, Open English or Review for College Writing<br />
			DWT099C, Open English Combined<br/>
			DWT099U, English/Critical Thinking Combined<br/>
			ENG101, English Composition 1<br />
			ENG101H, Honors English Composition 1<br />
			<br />
			DRG091, Reading Level 1<br />
			DRG092, Reading Level 2<br />
			READ105, Exempt from Reading<br />
			<br />
			OIT100, Basic Keyboarding Skills<br />
			OIT110, Exempt from Keyboarding </td>
		</tr>
	</table>
	<h2>Course Scheduling Advisor</h2>
	<table id="scheduleInfo">
		<tr>
		<td>Math:</td>
		<td><% =MathCourse(rs) %></td>
		</tr>
		<tr>
		<td>English:</td>
		<td><% =EnglishCourse(rs) %></td>
		</tr>
		<tr>
		<td>Reading:</td>
		<td><% =ReadingCourse(rs) %></td>
		</tr>
		<tr>
		<td>Typing:</td>
		<td><% =TypingCourse(rs) %></td>
		</tr>
	</table>

<%
    end if
    rs.close
    Set rs = Nothing
  end if
%></div>
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