<% 
  Option Explicit
  On Error Resume Next
  Response.Buffer = False
%>
<!-- #include file="library/library.asp" -->
<%
  Dim sql, fields, cond, startRow
  Dim conn, rs, strQuery, count, ErrMsg, sqlErr, tempstr, NoOfTests
  Dim StartDate, FinalDate, tmpStart, tmpFinal
  Dim MathPlacement, EnglPlacement, ReadPlacement, ExitPlacement, TypePlacement
  Dim todayDate, dateDifference
  Dim WPstr, ESstr
  
  if IsEmpty(Session("User")) then
    Response.Redirect "login.asp"
  elseif not Session("User")("IsAdmin") then 
    Response.Redirect "default.asp"
  end if

  Set conn = openConnection(Application("ConnectionString"))

  ErrMsg = ""
  if Request.ServerVariables("REQUEST_METHOD") = "POST" and Request("Action") = "Generate Report" then
    StartDate = Trim(Request("StartDate"))
    FinalDate = Trim(Request("FinalDate"))
    MathPlacement = Trim(Request("MathPlacement")) = "ON"
    EnglPlacement = Trim(Request("EnglPlacement")) = "ON"
    ReadPlacement = Trim(Request("ReadPlacement")) = "ON"
    TypePlacement = Trim(Request("TypePlacement")) = "ON"
    ExitPlacement = Trim(Request("ExitPlacement")) = "ON"
    
    NoOfTests = 0
    if MathPlacement then
      NoOfTests = NoOfTests + 1
    end if
    if EnglPlacement then
      NoOftests = NoOfTests + 1
    end if
    if ReadPlacement then
      NoOftests = NoOfTests + 1
    end if
    if TypePlacement then
      NoOftests = NoOfTests + 1
    end if
    if ExitPlacement then
      NoOftests = NoOfTests + 1
    end if

    todayDate = now()
    todayDate = CDate(month(todayDate) & "/" & day(todayDate) & "/" & year(todayDate))
    if StartDate <> "" and not IsDate(StartDate) then
      ErrMsg = BuildErrMsg(ErrMsg, "Bad Start Date")
    elseif FinalDate <> "" and not IsDate(FinalDate) then
      ErrMsg = BuildErrMsg(ErrMsg, "Bad Final Date")
    elseif IsDate(StartDate) and IsDate(FinalDate) then
      if CDate(FinalDate) < CDate(StartDate) then
        ErrMsg = BuildErrMsg(ErrMsg, "Final Date is before Starting Date")
      else
        dateDifference = DateDiff("d", StartDate, FinalDate)
        if dateDifference > 365 then
          ErrMsg = BuildErrMsg(ErrMsg, "Must Have Range <= 1 year")
        else
          tmpFinal = CDate(FinalDate)
          tmpStart = StartDate
        end if
      end if
    elseif StartDate = "" and FinalDate = "" then
      tmpFinal = todayDate + 1
      tmpStart = tmpFinal - 365
    elseif StartDate <> "" then
      tmpStart = CDate(StartDate)
      tmpFinal = iif(CDate(StartDate) + 365 < todayDate + 1, CDate(StartDate) + 365, todayDate + 1)
    else
      tmpFinal = CDate(FinalDate) + 1
      tmpStart = tmpFinal - 365   
    end if
         
    fields = ""
    cond = ""
    if ErrMsg = "" then
      if MathPlacement then
        if fields <> "" then
          fields = fields & ", "
        end if 
        fields = fields & "MathDate, MathArithScore, MathAlgScore, MathCollegeScore, MathPlacement"
        if cond <> "" then
          cond = cond & " OR "
        end if
        cond = cond & "(MathDate BETWEEN '" & tmpStart & "' AND '" & tmpFinal & "') "
      end if
      if EnglPlacement then
        if fields <> "" then
          fields = fields & ", "
        end if 
        fields = fields & "EnglDate, EnglScore, WPScore, EnglPlacement"
        if cond <> "" then
          cond = cond & " OR "
        end if
        cond = cond & "(EnglDate BETWEEN '" & tmpStart & "' AND '" & tmpFinal & "') "
       end if
      if ReadPlacement then
        if fields <> "" then
          fields = fields & ", "
        end if 
        fields = fields & "ReadDate, ReadScore, ReadPlacement"
        if cond <> "" then
          cond = cond & " OR "
        end if
        cond = cond & "(ReadDate BETWEEN '" & tmpStart & "' AND '" & tmpFinal & "') "
      end if
      if TypePlacement then
        if fields <> "" then
          fields = fields & ", "
        end if 
        fields = fields & "TypeDate, PassageNo, WordsPerMin, Errors, TypePlacement"
        if cond <> "" then
          cond = cond & " OR "
        end if
        cond = cond & "(TypeDate BETWEEN '" & tmpStart & "' AND '" & tmpFinal & "') "
      end if
      if ExitPlacement then
        if fields <> "" then
          fields = fields & ", "
        end if 
        fields = fields & "ExitDate, ExitScore, ExitPlacement"
        if cond <> "" then
          cond = cond & " OR "
        end if
        cond = cond & "(ExitDate BETWEEN '" & tmpStart & "' AND '" & tmpFinal & "') "
      end if
      
      if fields = "" then
        fields = "ProfDate"
      end if
      if cond = "" then
        cond = "(ProfDate BETWEEN '" & tmpStart & "' AND '" & tmpFinal & "')"
      end if
      sql = "SELECT StudentID, LastName, FirstName, " & fields & " FROM dbo.[Full Placement Testing Join]"
      if cond <> "" then
        sql = sql & " WHERE " & cond
      end if  
      sql = sql & " ORDER BY LastName, FirstName"
      
      sqlErr = ExecuteSQLForRs(conn, sql, rs)
      ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
      if rs.recordCount <= 0 then
        rs.Close
        Set rs = Nothing
        Errmsg = BuildErrMsg(ErrMsg, "There are no records that match the specified criteria")
      end if
    end if
  end if
%>
<!DOCTYPE html>
<html lang="en">

<head>
<meta charset="UTF-8" />
<title><% =SYSTEM_NAME %> - Placement Testings Reports</title>
<link rel="stylesheet" href="css/placement.css" />
<link rel="stylesheet" href="css/calendar.css" />
<link rel="stylesheet" href="css/gototop.css" />
<script src="js/calendar_us.js"></script>
<style>
table {
  margin-left: auto; 
  margin-right: auto;
  text-align: left;
}

th {
  text-align: center;	
}

td {
  vertical-align: top;
}
@media print {
  #report a {
	text-decoration:none;
	color:inherit;
  }
}
</style>
</head>

<body>
<div class="center">
<%
  Call MakeHeader("Placement Testing Reports")
  if ErrMsg <> "" then
    Response.Write "<h3 class=""red center"">Error(s) Occurred in Submission<br />" & ErrMsg & "</h3>" & vbNewLine
  end if
%>
<p class="hideElement bold">
<a style="margin-right: 10px" href="report.asp">Placement Testing Report</a>
<a style="margin-right: 10px" href="utilrept.asp">Utilities &amp; Reports</a>
<a style="margin-right: 10px" href="default.asp?ResetQuery=true">Main Menu</a>
<a href="default.asp?Logout=true">Log Out</a></p>
<hr class="hideElement" />
<%
  if ErrMsg <> "" or Request.ServerVariables("REQUEST_METHOD") = "GET" then
%>
<form method="post">
    <p class="bold">Start Date: <input type="text" id="StartDate" name="StartDate" size="10" maxlength="10" value="<%=StartDate%>" /> : 12:00am
    <script type="text/javascript">
	  new tcal ({
	    'controlname': 'StartDate'
	  })
	</script>
	&nbsp;&nbsp;&nbsp; Final Date: <input type="text" id="FinalDate" name="FinalDate" size="10" maxlength="10" value="<%=FinalDate%>" /> : 12:00am
    <script type="text/javascript">
	  new tcal ({
	    'controlname': 'FinalDate'
	  })
	</script><br />
  <span style="font-size:x-small" >Dates are entered in MM/DD/YY format.<br />
  	Empty Dates form an open ended boundary. Dates are inclusive.<br>Maximum 1 
	Year Time Span</span></p>

<table border="1" cellpadding="2" cellspacing="2">
  <tr>
    <th colspan="5">Include Placement Exam Information of Report</th>
  </tr>
  <tr>
    <th>STCC Math</th>
    <th>STCC English</th>
    <th>STCC Reading</th>
    <th>STCC Keyboarding</th>
    <th>Reading Exit</th>
  </tr>
  <tr class="center">
    <td><input type="checkbox" name="MathPlacement" value="ON" /></td>
    <td><input type="checkbox" name="EnglPlacement" value="ON" /></td>
    <td><input type="checkbox" name="ReadPlacement" value="ON" /></td>
    <td><input type="checkbox" name="TypePlacement" value="ON" /></td>
    <td><input type="checkbox" name="ExitPlacement" value="ON" /></td>
  </tr>
  </table>
<p><input name="Action" type="submit" value="Generate Report" /></p>
</form>

<%
  else
%>

<h2>STCC Placement Testing Report - <%=Date()%></h2>
<%
    Response.Write "<table id=""report"" border=""1"" cellspacing=""2"" cellpadding=""2"">" & vbNEwLine
    Response.Write "<tr>"
    Response.Write "<th>System<br/>ID</th>"
    Response.Write "<th>Name</th>"
    if NoOfTests > 0 then
      Response.Write "<th>Test</th>"
      Response.Write "<th>Placement</th>"
      Response.Write "<th>Date</th>"
      Response.Write "<th>Scores</th>"
    else
      Response.Write "<th>Date</th>"
    end if 
    Response.Write "</tr>" & vbNewLine
    do while not rs.eof 
      if NoOfTests = 0 then
        Response.Write "<tr class=""mainrow"">" & vbNewLine
        Response.Write "<td><a href=""editRecord.asp?value=" & rs("StudentID") & """>" & rs("StudentID") & "</a></td>" & vbNewLine
        Response.Write "<td>" & rs("LastName") & ", " & rs("FirstName") & "</td>" & vbNewLine
        Response.Write "<td>" & rs("ProfDate") & "</td>" & vbNewLine
      else
        StartRow = false
        Response.Write "<tr class=""mainrow"">" & vbNewLine
        Response.Write "<td rowspan=""" & NoOfTests & """><a href=""editRecord.asp?value=" & rs("StudentID") & """>" & rs("StudentID") & "</a></td>" & vbNewLine
        Response.Write "<td rowspan=""" & NoOfTests & """>" & rs("LastName") & ", " & rs("FirstName") & "</td>" & vbNewLine
        if MathPlacement then
          if StartRow then
            Response.Write "<tr>" & vbNewLine
          end if
          Response.Write "<td>Math</td>" & vbNewLine
          Response.Write "<td>" & iif(IsNull(rs("MathPlacement")), "&nbsp;", rs("MathPlacement")) & "</td>" & vbNewLine
          Response.Write "<td>" & iif(IsNull(rs("MathPlacement")), "&nbsp;", rs("MathDate")) & "</td>" & vbNewLine
          if IsNull(rs("MathPlacement")) then
            tempstr = "&nbsp;"
          else
            tempstr = Round(rs("MathArithScore")) & "-" & Round(rs("MathAlgScore")) & "-" & Round(rs("MathCollegeScore"))
          end if
          Response.Write "<td>" & tempstr & "</td>" & vbNewLine
          Response.Write "</tr>" & vbNewLine
          StartRow = true
        end if
        if EnglPlacement then
          if StartRow then
            Response.Write "<tr>" & vbNewLine
          end if
          Response.Write "<td>English</td>" & vbNewLine
          Response.Write "<td>" & iif(IsNull(rs("EnglPlacement")), "&nbsp;", rs("EnglPlacement")) & "</td>" & vbNewLine
          Response.Write "<td>" & iif(IsNull(rs("EnglPlacement")), "&nbsp;", rs("EnglDate")) & "</td>" & vbNewLine
          if IsNull(rs("EnglPlacement")) then
            tempstr = "&nbsp;"
          else
            if IsNull(rs("WPScore")) then
              WPStr = ""
            elseif rs("WPScore") < 0 then
              WPStr = ""
            else
              WPStr = rs("WPScore")
            end if
            if IsNull(rs("EnglScore")) then
              ESstr = ""
            elseif CDbl(rs("EnglScore")) <= 0 then
              ESstr = ""
            else
              ESstr = rs("EnglScore")
            end if
            tempstr = ESstr
            if tempstr <> "" AND WPstr <> "" then
              tempstr = tempstr & "-"
            end if
            tempstr = tempstr & WPStr
          end if
          Response.Write "<td>" & tempstr & "</td>" & vbNewLine
          Response.Write "</tr>" & vbNewLine
          StartRow = true
        end if
        if ReadPlacement then
          if StartRow then
            Response.Write "<tr>" & vbNewLine
          end if
          Response.Write "<td>Reading</td>" & vbNewLine
          Response.Write "<td>" & iif(IsNull(rs("ReadPlacement")), "&nbsp;", rs("ReadPlacement")) & "</td>" & vbNewLine
          Response.Write "<td>" & iif(IsNull(rs("ReadPlacement")), "&nbsp;", rs("ReadDate")) & "</td>" & vbNewLine
          if IsNull(rs("ReadPlacement")) then
            tempstr = "&nbsp;"
          else
            tempstr = Round(rs("ReadScore"))
          end if 
          Response.Write "<td>" & tempstr & "</td>" & vbNewLine
          Response.Write "</tr>" & vbNewLine
          StartRow = true
        end if
        if TypePlacement then
          if StartRow then
            Response.Write "<tr>" & vbNewLine
          end if
          Response.Write "<td>Keyboarding</td>" & vbNewLine
          Response.Write "<td>" & iif(IsNull(rs("TypePlacement")), "&nbsp;", rs("TypePlacement")) & "</td>" & vbNewLine
          Response.Write "<td>" & iif(IsNull(rs("TypePlacement")), "&nbsp;", rs("TypeDate")) & "</td>" & vbNewLine
          if IsNull(rs("TypePlacement")) then
            tempstr = "&nbsp;"
          else
            tempstr = rs("PassageNo") & "-" & rs("WordsPerMin") & "-" & rs("Errors")
          end if 
          Response.Write "<td>" & tempstr & "</td>" & vbNewLine
          Response.Write "</tr>" & vbNewLine
          StartRow = true
        end if
        if ExitPlacement then
          if StartRow then
            Response.Write "<tr>" & vbNewLine
          end if
          Response.Write "<td>Reading Exit</td>" & vbNewLine
          Response.Write "<td>" & iif(IsNull(rs("ExitPlacement")), "&nbsp;", rs("ExitPlacement")) & "</td>" & vbNewLine
          Response.Write "<td>" & iif(IsNull(rs("ExitPlacement")), "&nbsp;", rs("ExitDate")) & "</td>" & vbNewLine
          if IsNull(rs("ExitPlacement")) then
            tempstr = "&nbsp;"
          else
            tempstr = Round(rs("ExitScore"))
          end if
          Response.Write "<td>" & tempstr & "</td>" & vbNewLine
          Response.Write "</tr>" & vbNewLine
          StartRow = true
        end if
      end if
      rs.MoveNext
    loop
    rs.Close
    Set rs = Nothing
    Response.Write "</table>" & vbNewLine
  end if
%>
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

