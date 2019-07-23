<% 
  Option Explicit
  On Error Resume Next
%>
<!-- #include file="library/library.asp" -->
<%
  Dim Current_Page, Page_Count
  Dim conn, rs, sql, cond, tempstr, ErrMsg

  Dim StudentID, SSN, LastName, FirstName
  Dim DateFrom, DateTo, strFrom, strTo
  Dim j, tmppage
  Dim ItemsPerPage
  Dim PageURL: PageURL = ReplaceString(Request.ServerVariables("URL"), "(\?.*)", "", True)
  
  function decodeID(ByVal ID)
    Dim i
    decodeID = ""
    if Trim(ID) = "" then
      Exit Function
    end if
    if Mid(ID, 1, 2) = "XX" then
      ID = Mid(ID, 3)
      i = 1
      do while i < Len(ID)
        if Mid(ID, i, 1) <> "0" then
          Exit Do
        end if
        i = i + 1
      loop
      ID = Mid(ID, i)
    end if
    decodeID = ID
  end function
  
  Function ReplaceString( strOriginalString, strPattern, strReplacement, varIgnoreCase )
	If strOriginalString <> "" AND strPattern <> "" Then
		Dim objRegExp
		Set objRegExp = New RegExp
		
		With objRegExp
			.Pattern = strPattern
			.IgnoreCase = varIgnoreCase
			.Global = False
		End With
	
		ReplaceString = objRegExp.replace( strOriginalString, strReplacement )
		
		Set objRegExp = Nothing
	Else
		ReplaceString = strOriginalString
	End If
  End Function
  
  if IsEmpty(Session("User")) then
    Response.Redirect "login.asp"
  elseif Request("Logout") = "true" then
    if IsObject(Session("QueryResultsRs")) then
      Session("QueryResultsRs").Close
    end if
    Session.Abandon
    Response.Redirect "login.asp"
  elseif Session("User")("mustChangePassword") then
    Response.Redirect "changepassword.asp"
  elseif Request("ResetQuery") = "true" then
    if IsObject(Session("QueryResultRs")) then
      Session("QueryResultsRs").Close
    end if
    Session("QueryResultsRs") = Empty
    Session("Current_Page") = Empty
    Response.Redirect "default.asp"
  elseif Request("RefineSearch") = "true" then
    SSN = Session("SSN")
    LastName = Session("LastName")
    FirstName = Session("FirstName")
    DateFrom = Session("DateFrom")
    DateTo = Session("DateTo")
    StudentID = Session("StudentID")
    if IsObject(Session("QueryResultRs")) then
      Session("QueryResultsRs").Close
    end if
    Session("QueryResultsRs") = Empty 
    Session("Current_Page") = Empty
  end if

  Set conn = openConnection(Application("ConnectionString"))

  ErrMsg = ""
  if Request.ServerVariables("REQUEST_METHOD") = "POST" and Request("Action") = "Submit" then
    SSN = Replace(Trim(Request("SSN")), "-", "")
    
    LastName = Trim(Request("LastName"))
    FirstName = Trim(Request("FirstName"))
    DateFrom = Trim(Request("DateFrom"))
    DateTo = Trim(Request("DateTo"))
    StudentID = Trim(Request("StudentID"))
        
    strFrom = Replace(DateFrom, "/", "")
    strTo = Replace(DateTo, "/", "")

    if SSN & LastName & FirstName & strFrom & strTo & StudentID = "" then
      ErrMsg = BuildErrMsg(ErrMsg, "You Must Enter Some Searching Criteria.")
    end if
    if Len(SSN) = 8 Then
      ErrMsg = BuildErrMsg(ErrMsg, "Illegal ID/SSN")
    elseif Len(SSN) > 0 then
      If Len(SSN) <= 7 Then
        SSN = ZPadStr(SSN, 7)
      End If
      If Len(SSN) = 7 Then
        SSN = "XX" & SSN
      End If
    End if
    if strFrom <> "" then
      if not IsDate(DateFrom) then
        ErrMsg = BuildErrMsg(ErrMsg, "Start From Date is Illegal.")
      else
        DateFrom = CDate(DateFrom)
        if DateFrom < CDate("01/01/1900") then
          ErrMsg = BuildErrMsg(ErrMsg, "Start From Date is out of range.")
        end if
      end if
    end if
    if strTo <> "" then
      if not IsDate(DateTo) then
        ErrMsg = BuildErrMsg(ErrMsg, "Start To Date is Illegal.")
      else
        DateTo = CDate(DateTo)
        if DateTo  < CDate("01/01/1900") then
          ErrMsg = BuildErrMsg(ErrMsg, "Start To Date is out of range.")
        end if
      end if
    end if
    if IsDate(DateFrom) and IsDate(DateTo) then
      if CDate(DateFrom) > CDate(DateTo) then
        ErrMsg = BuildErrMsg(ErrMsg, "From Date is Greater than To Date.")
      end if
    end if

    if ErrMsg <> "" then
      Session("QueryResultsRs") = Empty
    else
      sql = "SELECT StudentID, LastName, FirstName, MathPlacement, EnglPlacement, EssayID, ReadPlacement, " & _
            "       ExitPlacement, TypePlacement FROM dbo.[Full Placement Testing Join] WHERE "
      if SSN <> "" then
        tempstr = "SSN = '" & SSN & "'"
        cond = appendCondition(cond, tempstr)
      end if
      if LastName <> "" then
        tempstr = "LastName Like '" & Replace(Replace(Replace(LastName, "*", "%"), "?", "_"), "'", "''") & "'" 
        cond = appendCondition(cond, tempstr)
      end if
      if FirstName <> "" then
        tempstr = "FirstName Like '" & Replace(Replace(Replace(FirstName, "*", "%"), "?", "_"), "'", "''") & "'" 
        cond = appendCondition(cond, tempstr)
      end if
      if strFrom & strTo <> "" then
        if strFrom <> "" then 
          if strTo <> "" then
            tempstr = "((ProfDate BETWEEN '" & DateFrom & "' AND '" & DateTo & "') OR "
            tempstr = tempstr & "(MathDate BETWEEN '" & DateFrom & "' AND '" & DateTo & "') OR "
            tempstr = tempstr & "(EnglDate BETWEEN '" & DateFrom & "' AND '" & DateTo & "') OR "
            tempstr = tempstr & "(ReadDate BETWEEN '" & DateFrom & "' AND '" & DateTo & "') OR "
            tempstr = tempstr & "(ExitDate BETWEEN '" & DateFrom & "' AND '" & DateTo & "') OR "
            tempstr = tempstr & "(TypeDate BETWEEN '" & DateFrom & "' AND '" & DateTo & "'))"
          else
            tempstr = "((ProfDate >= '" & DateFrom & "') OR "
            tempstr = tempstr & "(MathDate >= '" & DateFrom & "') OR "
            tempstr = tempstr & "(EnglDate >= '" & DateFrom & "') OR "
            tempstr = tempstr & "(ReadDate >= '" & DateFrom & "') OR "
            tempstr = tempstr & "(ExitDate >= '" & DateFrom & "') OR "
            tempstr = tempstr & "(TypeDate >= '" & DateFrom & "'))"
          end if
        elseif strTo <> "" then
            tempstr = "((ProfDate <= '" & DateTo & "') OR "
            tempstr = tempstr & "(MathDate <= '" & DateTo & "') OR "
            tempstr = tempstr & "(EnglDate <= '" & DateTo & "') OR "
            tempstr = tempstr & "(ReadDate <= '" & DateTo & "') OR "
            tempstr = tempstr & "(ExitDate <= '" & DateTo & "') OR "
            tempstr = tempstr & "(TypeDate <= '" & DateTo & "'))"
        end if
        cond = appendCondition(cond, tempstr)
      end if
      if StudentID <> "" then
        tempstr = "StudentID = " & StudentID 
        cond = appendCondition(cond, tempstr)
      end if
      
      sql = sql & cond
      ErrMsg = ExecuteSQLForRS(conn, sql, rs)

      if ErrMsg = "" then
        if rs.recordCount <= 0 then
          rs.Close
          Set rs = Nothing
          Errmsg = BuildErrMsg(ErrMsg, "There are no records that match the specified criteria.")
        else
          Page_Count = RS.PageCount
          ItemsPerPage = CInt(iif(IsEmpty(Session("ItemsPerPage")), Application("ItemsPerPage"), Session("ItemsPerPage")))
          rs.PageSize = iif(ItemsPerPage <= 0, rs.recordCount, ItemsPerPage)
          rs.Sort = "Lastname ASC, Firstname ASC"
          Set Session("QueryResultsRs") = rs
        end if
      end if
      
      Session("SSN") = Request("SSN")
      Session("LastName") = Request("LastName")
      Session("FirstName") = Request("FirstName")
      Session("DateFrom") = Request("DateFrom")
      Session("DateTo") = Request("DateTo")
      Session("StudentID") = Request("StudentID")
    end if
  ElseIf IsObject(Session("QueryResultsRs")) Then
    Set rs = Session("QueryResultsRs")
    rs.ActiveConnection = conn
    Call rs.Requery()
    rs.ActiveConnection = Nothing
    if Request("SortBy") = "Name" then 
      if rs.Sort <> "Lastname ASC, Firstname ASC" then
        rs.Sort = "Lastname ASC, Firstname ASC"
      else 
        rs.Sort = "Lastname DESC, Firstname DESC"
      end if
    end if
    if rs.recordCount <= 0 then
      'Errmsg = ErrMsg & "There are no records that match the specified criteria"
      Session("QueryResultsRs").Close
      Session("QueryResultsRs") = Empty
    end if
  Else
    Session("QueryResultsRs") = Empty
  End If
%>
<!DOCTYPE html>
<html lang="en">

<head>
<meta charset="UTF-8" />
<title><% =SYSTEM_NAME %> - Query List / Main Menu</title>
<link rel="stylesheet" href="css/placement.css" />
<link rel="stylesheet" href="css/gototop.css" />
<% if not IsObject(Session("QueryResultsRs")) Then %>
<link rel="stylesheet" href="css/calendar.css" />
<script src="js/calendar_us.js"></script>
<% end if %>
<style>
.firstcol  {
	width: 260px;
	height: 30px;
	text-align: right;
	font-weight: bold;
}
.secondcol {
	width: 5px;
	height: 30px;
}
.thirdcol {
	width: 400px;
	height: 30px;
	text-align: left;
	font-weight: bold;
}
.submitrow {
	height: 40px;
	text-align: center;
}
#resultsTable td {
	text-align: left;
}
#resultsTable td.center {
	text-align: center;
}
</style>
</head>

<body>
<div class="center">

<%
  Call MakeHeader("Query List / Main Menu")
  if ErrMsg <> "" then
    Response.Write "<h3 class=""red "">" & ErrMsg & "</h3>" & vbNewLine
  end if
%>
	<p class="bold">
	<% if Session("User")("IsWriter") then %>
	  <a style="margin-right: 10px" href="addRecord.asp">Add Record</a>
	<% end if %>
	<% if Session("User")("IsEssay")  then %>
	  <a style="margin-right: 10px" href="EssayUtilities.asp">Essay Utilities</a>
	<% end if %>
	<% if Session("User")("IsAdmin") then %>
	  <a style="margin-right: 10px" href="utilrept.asp">Utilities/Reports</a>
	<% end if %>
	<a style="margin-right: 10px" href="changepassword.asp">Change Password</a>
	<a style="margin-right: 10px" href="default.asp?RefineSearch=true">Refine Search</a>
	<a style="margin-right: 10px" href="default.asp?ResetQuery=true">New Query</a>
	<a href="default.asp?Logout=true">Log Out</a>
	</p>
	<hr />
<%
  if not IsObject(Session("QueryResultsRs")) then 
%>
	<div style="width: 680px; margin-left: auto; margin-right: auto;">
		<form method="post" style="margin: 5px 0 5px 0">
			<table border="0" cellspacing="0" cellpadding="0" width="680px">
				<tr>
					<td class="firstcol"><label for="SSNText">Student ID/SSN:</label></td>
					<td class="secondcol">&nbsp;</td>
					<td class="thirdcol"><input type="text" id="SSNText" name="SSN" size="11" value="<%=decodeID(SSN)%>" maxlength="11" /></td>
				</tr>
				<tr>
					<td class="firstcol green"><label for="LastNameText">Last Name:</label></td>
					<td class="secondcol">&nbsp;</td>
					<td class="thirdcol"><input type="text" id="LastNameText" name="LastName" size="25" maxlength="25" value="<%=LastName%>" /></td>
				</tr>
				<tr>
					<td class="firstcol green"><label for="FirstNameText">First Name:</label></td>
					<td class="secondcol">&nbsp;</td>
					<td class="thirdcol"><input type="text" id="FirstNameText" name="FirstName" size="25" maxlength="25" value="<%=FirstName%>" /></td>
				</tr>
				<tr>
					<td class="firstcol"><label for="DateFromText">From Date:</label></td>
					<td class="secondcol">&nbsp;</td>
					<td class="thirdcol"><input type="text" id="DateFromText" name="DateFrom" size="10" maxlength="10" value="<%=DateFrom%>" /> : 12:00am
					<script type="text/javascript">
					  new tcal ({
					    'controlname': 'DateFromText'
					  })
					</script></td>
				</tr>
				<tr>
					<td class="firstcol"><label for="DateToText">To Date:</label></td>
					<td class="secondcol">&nbsp;</td>
					<td class="thirdcol"><input type="text" id="DateToText" name="DateTo" size="10" maxlength="10" value="<%=DateTo%>" /> : 12:00am
					<script type="text/javascript">
					  new tcal ({
					    'controlname': 'DateToText'
					  })
					</script></td>
				</tr>
				<tr>
					<td class="firstcol"><label for="StudentIDText">System ID:</label></td>
					<td class="secondcol">&nbsp;</td>
					<td class="thirdcol">
					<input type="text" id="StudentIDText" name="StudentID" size="7" value="<%=StudentID%>" maxlength="7" /></td>
				</tr>
				<tr>
					<td colspan="3" class="submitrow"><input type="submit" value="Submit" name="Action" /></td>
				</tr>
			</table>
		</form>
	</div>
				<hr />
	<div style="width: 500px; margin-left: auto; margin-right: auto;">

				<h4>Search Field Entry Instructions</h4>
					<ul style="text-align: left">
						<li>Leave fields blank if not part of the searching criteria.</li>
						<li>Multiple entries are AND&#39;d together for a more restrictive 
						search.</li>
						<li>Fields marked in <b class="green">green</b> are searched 
						for likeness;<br />
						use * to match any 0 or more characters; use 
						? to match any one character.</li>
						<li>If no Dates are entered, Date is <b>not</b> part of 
						the search.</li>
						<li>If only one Date is entered, results will be open ended 
						on the empty date.</li>
						<li>Dates are inclusive and have this input format <b>MM/DD/YYYY</b></li>
					</ul>
	</div>
<%
  else
    If Request("ItemsPerPage") = "" Then
      ItemsPerPage = Cint(iif(IsEmpty(Session("ItemsPerPage")), Application("ItemsPerPage"), Session("ItemsPerPage")))
      if Request("Page") = "" Then
        Current_Page = CInt(iif(IsEmpty(Session("Current_Page")), 1, Session("Current_Page")))
      else
        Current_Page = CInt(Request("Page"))
      end if
    else
      ItemsPerPage = CInt(Request("ItemsPerPage"))
      Current_Page = 1
    end if
    rs.PageSize = iif(ItemsPerPage <= 0, rs.recordCount, ItemsPerPage)
    Session("ItemsPerPage") = ItemsPerPage
    
    Page_Count = rs.pageCount
    If Current_Page < 1 Then 
      Current_Page = 1
    elseif Current_Page > Page_Count Then 
      Current_Page = Page_Count
    end if
    rs.AbsolutePage = Current_Page
    Session("Current_Page") = Current_Page
     
    Response.Write "<form method=""post"" action=""" & PageURL & """>"
    Response.Write "<p class=""bold""><label>Display: "
    Response.Write "<select id=""ItemsPerPage"" name=""ItemsPerPage"">"
    Response.Write "<option value=""10"">10</option>"
    Response.Write "<option value=""25"">25</option>"
    Response.Write "<option value=""50"">50</option>"
    Response.Write "<option value=""75"">75</option>"
    Response.Write "<option value=""100"">100</option>"
    Response.Write "<option value=""-1"">All</option>"
    Response.Write "</select> records</label>" 
    Response.Write "<span style=""margin-left: 70px"">Results Page " & Current_Page & " of " & Page_Count & vbNewLine
    Response.Write " -- " & rs.recordCount & " Total Records <small><small>(" & rs.PageSize & " records per page)</small></small></span>" & vbNewLine
    Response.Write "</p></form>" & vbNewLine
    
    if Page_Count > 1 then
      Response.Write "<p class=""bold"">"
      if Current_Page > 1 Then
        Response.Write "<a href=""?Page=" & 1 & """>First</a><span style=""margin-right: 10px"">&nbsp;</span>"
        Response.Write "<a href=""?Page=" & Current_Page - 1 & """>Prev</a><span style=""margin-right: 10px"">&nbsp;</span>"
      else
         Response.Write "First<span style=""margin-right: 10px"">&nbsp;</span>"
         Response.Write "Prev<span style=""margin-right: 10px"">&nbsp;</span>"
      end If    
      for j = -10 to 9
        tmppage = Current_Page + j
        if tmppage >= 1 and tmppage <= Page_Count then
          if tmppage = Current_Page then
	        Response.Write "<span style=""color:#a90a08"">" & tmppage & "</span>"
	      else
	        Response.Write "<a href=""?Page=" & tmppage & """>" & tmppage & "</a>"
	      end if
	      if j <> 9 and tmppage < Page_Count then
	        Response.Write "<span style=""margin-right: 10px"">&nbsp;</span>"
	      end if
	    end if
      next 
      if Current_Page < Page_Count then
        Response.Write "<span style=""margin-right: 10px"">&nbsp;</span><a href=""?Page=" & Current_Page + 1 & """>Next</a>"
        Response.Write "<span style=""margin-right: 10px"">&nbsp;</span><a href=""?Page=" & Page_Count & """>Last</a>"
      else
        Response.Write "<span style=""margin-right: 10px"">&nbsp;</span>Next"
        Response.Write "<span style=""margin-right: 10px"">&nbsp;</span>Last"
      end if
      Response.Write "</p>" & vbNewLine
      if Page_Count >= 20 then
        Response.Write "<form method=""post"" action=""" & PageURL & """ onsubmit=""return checkValue(this)"">"
        Response.Write "<p class=""center"">" & vbNewLine
        Response.Write "<input style=""margin-right: 10px"" type=""text"" name=""Page"" size=""5"" value=""" & Current_Page & """ />" & vbNewLine
        Response.Write "<input type=""submit"" value=""Go"" />" & vbNewLine
        Response.Write "</p></form>"  & vbNewLine
        Response.Write "<p class=""center"" id=""resultstring""></p>" & vbNewLine
      end if    
    end if

    Response.Write "<table id=""resultsTable"" border=""1"" cellpadding=""4"" cellspacing=""2"" style=""margin-left: auto; margin-right: auto;"">" & vbNewLine
    Response.Write "<tr>" & vbNewLine
    Response.Write "<th rowspan=""2"">" & "<a href=""?SortBy=Name"">" & "Name" & "</a></th>" & vbNewLine
    if Session("User")("IsWriter") then 
      Response.Write "<th rowspan=""2"">" & "Edit<br />Record" & "</th>" & vbNewLine
    end if
    Response.Write "<th colspan=""5"">" & "STCC Placement Exam" & "</th>" & vbNewLine
    Response.Write "<th rowspan=""2"">" & "Reading Exit<br />Placement" & "</th>" & vbNewLine
    Response.Write "</tr>" & vbNewLine
    
    Response.Write "<tr>" & vbNewLine
    Response.Write "<th>" & "Math"     & "</th>" & vbNewLine
    Response.Write "<th>" & "English"  & "</th>" & vbNewLine
    Response.Write "<th>" & "Essay"    & "</th>" & vbNewLine
    Response.Write "<th>" & "Reading"  & "</th>" & vbNewLine    
    Response.Write "<th>" & "Keyboard" & "</th>" & vbNewLine    
    Response.Write "</tr>" & vbNewLine

    Do While RS.AbsolutePage = Current_Page AND Not RS.EOF
      Response.Write "<tr>" & vbNewLine
      Response.Write "<td>" & "<a href=""viewRecord.asp?value=" & rs("StudentID") & """>" & rs("LastName") & ", " & rs("FirstName") & "</a>" & "</td>" & vbNewLine
      if Session("User")("IsWriter") then 
        Response.Write "<td class=""center"">" & "<a href=""editRecord.asp?value=" & rs("StudentID") & """>" & "<img src=""img/edit-white.gif"" height=""16"" width=""35"" alt=""Edit"" />" & "</a>" & "</td>" & vbNewLine
      end if
      Response.Write "<td>" & FixNull(rs("MathPlacement")) & "</td>" & vbNewLine
      Response.Write "<td>" & FixNull(rs("EnglPlacement")) & "</td>" & vbNewLine 
      Response.Write "<td class=""center"">" 
      if IsNull(rs("EssayID")) then
        Response.Write "&nbsp;"
      else
        Response.Write "<a href=""viewEssay.asp?EssayID=" & rs("EssayID") & """>View</a>"
      end if
      Response.Write "</td>" & vbNewLine
      Response.Write "<td>" & FixNull(rs("ReadPlacement")) & "</td>" & vbNewLine
      Response.Write "<td>" & FixNull(rs("TypePlacement")) & "</td>" & vbNewLine
      Response.Write "<td>" & FixNull(rs("ExitPlacement")) & "</td>" & vbNewLine
      Response.Write "</tr>" & vbNewLine
      rs.MoveNext
    Loop
    Response.Write "</table>" & vbNewLine
  end if
%>
</div>
<script src="//ajax.googleapis.com/ajax/libs/jquery/2.2.4/jquery.min.js"></script> 
<script src="js/jquery.gototop.js"></script> 
<script>
$(function(){

  window.checkValue = function (frm) {
    if (isNaN(frm.Page.value)) {
      resultstring.innerText = 'Not a Number'
      return false
      }
    var pageno = parseInt(frm.Page.value)
    frm.Page.value = pageno
    if (pageno >= 1 && pageno <= 20)
      return true
    resultstring.innerText = 'Page Number Out of Range'
    return false
  }
  
  $("#ItemsPerPage").val("<%=ItemsPerPage%>");
  
  $("#ItemsPerPage").change(function() {
     $(this).closest("form").submit();
  });
  
  $("#toTop").gototop({ container: "body" });
  $("#SSNText").focus();
});
</script>
</body>
</html>
<%
  closeConnection(conn)
%>