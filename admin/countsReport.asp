<%
  Option Explicit
  On Error Resume Next
%>
<!-- #include file="library/library.asp" -->
<%
  Dim conn, rs, sql, count, i, ErrMsg, sqlErr, cond
  Dim StartDate, FinalDate
  Dim TestType
  Dim MathPlacement, EnglishPlacement, ReadingPlacement, ExitPlacement, TypePlacement
  Dim MathAll, MathNone, EnglishAll, EnglishNone, ReadingAll, ReadingNone, ExitAll, ExitNone, TypeAll, TypeNone
  
  if IsEmpty(Session("User")) then
    Response.Redirect "login.asp"
  elseif not Session("User")("IsAdmin") then 
    Response.Redirect "default.asp"
  end if
  
  Set conn = openConnection(Application("ConnectionString"))

  Function IsSelected(items, item)
    Dim i
    if not items is Nothing then
      for i = 1 to items.Count
        if items(i) = item then
          IsSelected = "selected=""selected"""
          Exit Function
        end if
      next
    end if
    IsSelected = ""
  End Function
  
  Function IsChecked(item, value)
    IsChecked = iif(item = value, "checked=""checked""", "")
  End Function
  
  Function DateCond(ByVal StartDate, ByVal FinalDate)
    Dim temp
 
    temp = ""
    if IsDate(StartDate) then
      if IsDate(FinalDate) then
        temp = temp & "ProfDate BETWEEN '" & StartDate & "' AND '" & FinalDate & "'"
      else 
        temp = temp & "ProfDate >= '" & StartDate & "'"
      end if
    elseif IsDate(FinalDate) then
      temp = temp & "ProfDate <= '" & FinalDate & "'"
    end if
    
    DateCond = temp
  End Function
  
  Function BuildCond(ByVal TotalCondition, ByVal Cond)
    if Cond = "" then
      BuildCond = TotalCondition
    else
      Cond = "(" & Cond & ")"
      if TotalCondition = "" then 
        BuildCond = Cond
      else
        BuildCond = TotalCondition & " AND " & Cond
      end if
    end if
  End Function

  Function PlacementCond(ByVal MathAll, ByVal MathNone, ByVal MathPlacement, _
                         ByVal EnglishAll, ByVal EnglishNone, ByVal EnglishPlacement, _
                         ByVal ReadingAll, ByVal ReadingNone, ByVal ReadingPlacement, _
                         ByVal ExitAll, ByVal ExitNone, ByVal ExitPlacement, _
                         ByVal TypeAll, ByVal TypeNone, ByVal TypePlacement, _
                         ByVal StartDate, ByVal FinalDate)
    Dim dcond: dcond = ""
    Dim cond: cond = ""
    Dim temp
    
    temp = ""
    if MathAll <> "" then
      temp = "MathPlacement Is Not Null"
    elseif MathNone <> "" then
      temp = "MathPlacement Is Null"
    elseif not IsEmpty(MathPlacement) then
      temp = "MathPlacement IN ("
      for i = 1 to MathPlacement.Count
        temp = temp & "'" & MathPlacement(i) & "'"
        if i < MathPlacement.Count then
          temp = temp & ", "
        end if
      next
      temp = temp & ")"
    end if
    if temp <> "" then
      cond = BuildCond(cond, temp)
    end if
      
    temp = ""
    if EnglishAll <> "" then
      temp = "EnglPlacement Is Not Null"
    elseif EnglishNone <> "" then
      temp = "EnglPlacement Is Null"
    elseif not IsEmpty(EnglishPlacement) then
      temp = "EnglPlacement IN ("
      for i = 1 to EnglishPlacement.Count
        temp = temp & "'" & EnglishPlacement(i) & "'"
        if i < EnglishPlacement.Count then
          temp = temp & ", "
        end if
      next
      temp = temp & ")"
    end if
    if temp <> "" then
      cond = BuildCond(cond, temp)
    end if
    
    temp = ""
    if ReadingAll <> "" then
      temp = "ReadPlacement Is Not Null"
    elseif ReadingNone <> "" then
      temp = "ReadPlacement Is Null"
    elseif not IsEmpty(ReadingPlacement) then
      temp = "ReadPlacement IN ("
      for i = 1 to ReadingPlacement.Count
        temp = temp & "'" & ReadingPlacement(i) & "'"
        if i < ReadingPlacement.Count then
          temp = temp & ", "
        end if
      next
      temp = temp & ")"
    end if
    if temp <> "" then
      cond = BuildCond(cond, temp)
    end if
   
    temp = ""
    if ExitAll <> "" then
      temp = "ExitPlacement Is Not Null"
    elseif ExitNone <> "" then
      temp = "ExitPlacement Is Null"
    elseif not IsEmpty(ExitPlacement) then
      temp = "ExitPlacement IN ("
      for i = 1 to ExitPlacement.Count
        temp = temp & "'" & ExitPlacement(i) & "'"
        if i < ExitPlacement.Count then
          temp = temp & ", "
        end if
      next
      temp = temp & ")"
    end if
    if temp <> "" then
      cond = BuildCond(cond, temp)
    end if

    temp = ""
    if TypeAll <> "" then
      temp = "TypePlacement Is Not Null"
    elseif TypeNone <> "" then
      temp = "TypePlacement Is Null"
    elseif not IsEmpty(TypePlacement) then
      temp = "TypePlacement IN ("
      for i = 1 to TypePlacement.Count
        temp = temp & "'" & TypePlacement(i) & "'"
        if i < TypePlacement.Count then
          temp = temp & ", "
        end if
      next
      temp = temp & ")"
    end if
    if temp <> "" then
      cond = BuildCond(cond, temp)
    end if 
    
    dCond = DateCond(StartDate, FinalDate)
    cond = BuildCond(cond, dCond)

    PlacementCond = cond                     
  End Function
  
  ErrMsg = ""
  StartDate = Trim(Request("StartDate"))
  FinalDate = Trim(Request("FinalDate"))
  MathAll = Request("MathAll")
  MathNone = Request("MathNone")
  EnglishAll = Request("EnglishAll")
  EnglishNone = Request("EnglishNone")
  ReadingAll = Request("ReadingAll")
  ReadingNone = Request("ReadingNone")
  ExitAll = Request("ExitAll")
  ExitNone = Request("ExitNone")
  TypeAll = Request("TypeAll")
  TypeNone = Request("TypeNone")
  Set TestType = Request("TestType")
  if not IsEmpty(TestType) then
    Set MathPlacement = Nothing
    Set EnglishPlacement = Nothing
    Set ReadingPlacement = Nothing
    Set ExitPlacement = Nothing
    Set TypePlacement = Nothing
  else
    Set MathPlacement = Request("MathPlacement")
    Set EnglishPlacement = Request("EnglishPlacement")
    Set ReadingPlacement = Request("ReadingPlacement")
    Set ExitPlacement = Request("ExitPlacement")
    Set TypePlacement = Request("TypePlacement")
  end if
  if Request.ServerVariables("REQUEST_METHOD") = "POST" and Request("Action") = "Generate Report" then
    if StartDate <> "" and not IsDate(StartDate) then
      ErrMsg = ErrMsg & "Bad Start Date" & "<br>"
    end if
    if FinalDate <> "" and not IsDate(FinalDate) then
      ErrMsg = ErrMsg & "Bad Final Date" & "<br>"
    end if
    if IsDate(StartDate) and IsDate(FinalDate) then
      if CDate(FinalDate) < CDate(StartDate) then
        ErrMsg = ErrMsg & "Final Date is before Starting Date" & "<br>"
      end if
    end if
    if ErrMsg = "" then
      cond = PlacementCond(MathAll, MathNone, MathPlacement, _
                           EnglishAll, EnglishNone, EnglishPlacement, _
                           ReadingAll, ReadingNone, ReadingPlacement, _
                           ExitAll, ExitNone, ExitPlacement, _
                           TypeAll, TypeNone, TypePlacement, _
                           StartDate, FinalDate)
      
      sql = "SELECT Count(*) FROM dbo.[Full Placement Testing Join]"
      if cond <> "" then
        sql = sql & " WHERE " & cond  
      end if
      
      sqlErr = ExecuteSQLForRs(conn, sql, rs)
      ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
            
      if rs.recordCount <= 0 then
        ErrMsg = BuildErrMsg(ErrMsg, "There are no records that match the specified criteria")
      else
        count = rs(0)
      end if
      rs.Close
      Set rs = Nothing
    end if
  end if
%>
<!DOCTYPE html>
<html lang="en">

<head>
<title><% =SYSTEM_NAME %> - Counts Report</title>
<meta charset="UTF-8" />
<link rel="stylesheet" href="css/placement.css" />
<link rel="stylesheet" href="css/calendar.css" />
<link rel="stylesheet" href="css/gototop.css" />
<script src="js/calendar_us.js"></script>
<style>
table#PlacementReport {
	margin-left:auto;
	margin-right:auto;
	border: solid thin black;
	text-align: left;
	border-collapse:collapse;
}
table#PlacementReport td {
	padding: 5px;
	border: solid thin black;
}
table#PlacementReport th {
	padding: 5px;
	border: solid thin black;
	text-align: center;
}
table#PlacementReport #SelectBoxes td {
    width: 175px;
    vertical-align: top;
    text-align: center;
}
.selectParent {
	display:inline-block; vertical-align:top; overflow:hidden; border:solid grey 1px;
}
.selectParent select {
    padding:10px; 
    margin:-5px -20px -5px -5px;
}
</style>
</head>

<body>
<div class="center">
<%
  Call MakeHeader("Counts Report")
  if Request.ServerVariables("REQUEST_METHOD") = "POST" then
    if ErrMsg <> "" then
      Response.Write "<h3 class=""red"">Error(s) Occurred in Submission<br />" & ErrMsg & "</h3>" & vbNewLine
    end if
  end if
%>
<p class="bold">
<a style="margin-right: 10px" href="countsReport.asp">Counts Report</a>
<a style="margin-right: 10px" href="utilrept.asp">Utilities &amp; Reports</a>
<a style="margin-right: 10px" href="default.asp?ResetQuery=true">Main Menu</a>
<a href="default.asp?Logout=true">Log Out</a></p>
<hr />
<%
  if Request.ServerVariables("REQUEST_METHOD") = "POST" and ErrMsg = "" then
    if Cond = "" then
      Cond = "NONE"
    end if
    Response.Write "<h3 class=""blue center"">" & "The Number Of Students = " & Count & "<br />" & _
                    "Condition: " & Cond & "</h3>" & vbNewLine
    Response.Write "<hr />" & vbNewLine
  end if
%>
<form method="post">
	<p class="bold">Start Date: <input type="text" id="StartDate" name="StartDate" size="10" maxlength="10" value="<%=StartDate%>" /> : 12:00am
	<script>
	  new tcal ({
	    'controlname': 'StartDate'
	  })
	</script>
	&nbsp;&nbsp;&nbsp; Final Date: <input type="text" id="FinalDate" name="FinalDate" size="10" maxlength="10" value="<%=FinalDate%>" /> : 12:00am
	<script>
	  new tcal ({
	    'controlname': 'FinalDate'
	  })
	</script><br />
	<span style="font-size:x-small">Dates are entered in MM/DD/YY format.<br />
	Empty Dates form an open ended boundary. Dates are inclusive.</span></p>
	<h5>(Hold Control Key While Selecting/Deselecting Options)</h5>
	<table id="PlacementReport">
		<tr>
			<th colspan="5">Placement Levels</th>
		</tr>
		<tr>
			<th>Math</th>
			<th>English</th>
			<th>Reading</th>
			<th>Reading Exit</th>
			<th>Typing</th>
		</tr>
		<tr>
		<td>
		<label><input class="CheckBoxAll" id="MathAll" type="checkbox" name="MathAll" value="MathAll" <%=IsChecked(MathAll, "MathAll")%> data-selectbox="MathPlacement" /> 
		Toggle All Placements</label><br/>
		<label><input class="CheckBoxNone" id="MathNone" type="checkbox" name="MathNone" value="MathNone" <%=IsChecked(MathNone, "MathNone")%> data-selectbox="MathPlacement" /> No Placement Taken</label>
		</td>
		<td>
		<label><input class="CheckBoxAll" id="EnglishAll" type="checkbox" name="EnglishAll" value="EnglishAll" <%=IsChecked(EnglishAll, "EnglishAll")%> data-selectbox="EnglishPlacement" />
		Toggle All Placements</label><br/>
		<label><input class="CheckBoxNone" id="EnglishNone" type="checkbox" name="EnglishNone" value="EnglishNone" <%=IsChecked(EnglishNone, "EnglishNone")%> data-selectbox="EnglishPlacement" /> No Placement Taken</label>
		</td>
		<td>
		<label><input class="CheckBoxAll" id="ReadingAll" type="checkbox" name="ReadingAll" value="ReadingAll" <%=IsChecked(ReadingAll, "ReadingAll")%> data-selectbox="ReadingPlacement" />
		Toggle All Placements</label><br/>
		<label><input class="CheckBoxNone" id="ReadingNone" type="checkbox" name="ReadingNone" value="ReadingNone" <%=IsChecked(ReadingNone, "ReadingNone")%> data-selectbox="ReadingPlacement" /> No Placement Taken</label>
		</td>
		<td>
		<label><input class="CheckBoxAll" id="ExitAll" type="checkbox" name="ExitAll" value="ExitAll" <%=IsChecked(ExitAll, "ExitAll")%> data-selectbox="ExitPlacement" />
		Toggle All Placements</label><br/>
		<label><input class="CheckBoxNone" id="ExitNone" type="checkbox" name="ExitNone" value="ExitNone" <%=IsChecked(ExitNone, "ExitNone")%> data-selectbox="ExitPlacement" /> No Placement Taken</label>
		</td>
		<td>
		<label><input class="CheckBoxAll"t id="TypeAll" type="checkbox" name="TypeAll" value="TypeAll" <%=IsChecked(TypeAll, "TypeAll")%>  data-selectbox="TypePlacement" />
		Toggle All Placements</label><br/>
		<label><input class="CheckBoxNone" id="TypeNone" type="checkbox" name="TypeNone" value="TypeNone" <%=IsChecked(TypeNone, "TypeNone")%> data-selectbox="TypePlacement" /> No Placement Taken</label>
		</td>
		</tr>
		<tr id="SelectBoxes">
			<td>
			<div class="selectParent">
		    <select size="17" id="MathPlacement" name="MathPlacement" multiple="multiple" data-allcheckbox="MathAll" data-nonecheckbox="MathNone">
			<option <%=IsSelected(MathPlacement, "ARTH071")%>>ARTH071</option>
			<option <%=IsSelected(MathPlacement, "MAT071")%>>MAT071</option>
			<option <%=IsSelected(MathPlacement, "ARTH071U")%>>ARTH071U</option>
			<option <%=IsSelected(MathPlacement, "MAT071U")%>>MAT071U</option>
			<option <%=IsSelected(MathPlacement, "ALGB081")%>>ALGB081</option>
			<option <%=IsSelected(MathPlacement, "ALGB081U")%>>ALGB081U</option>
			<option <%=IsSelected(MathPlacement, "MAT081")%>>MAT081</option>
			<option <%=IsSelected(MathPlacement, "ARTH081U")%>>ARTH081U</option>
			<option <%=IsSelected(MathPlacement, "MAT081U")%>>MAT081U</option>
			<option <%=IsSelected(MathPlacement, "ALGB091")%>>ALGB091</option>
			<option <%=IsSelected(MathPlacement, "MAT091")%>>MAT091</option>
			<option <%=IsSelected(MathPlacement, "MATH101")%>>MATH101</option>
			<option <%=IsSelected(MathPlacement, "MAT101")%>>MAT101</option>
			<option <%=IsSelected(MathPlacement, "MATH105")%>>MATH105</option>
			<option <%=IsSelected(MathPlacement, "MAT105")%>>MAT105</option>
			<option <%=IsSelected(MathPlacement, "MATH155")%>>MATH155</option>
			<option <%=IsSelected(MathPlacement, "MAT131")%>>MAT131</option>
			</select>
			</div>
			</td>
			<td>
			<div class="selectParent">
			<select size="8" id="EnglishPlacement" name="EnglishPlacement" multiple="multiple" data-allcheckbox="EnglishAll" data-nonecheckbox="EnglishNone">
			<option <%=IsSelected(EnglishPlacement, "DWT099")%>>DWT099</option>
			<option <%=IsSelected(EnglishPlacement, "DWT099U")%>>DWT099U</option>
			<option <%=IsSelected(EnglishPlacement, "DWT099C")%>>DWT099C</option>
			<option <%=IsSelected(EnglishPlacement, "DWRT099")%>>DWRT099</option>
			<option <%=IsSelected(EnglishPlacement, "ENG101")%>>ENG101</option>
			<option <%=IsSelected(EnglishPlacement, "ENGL100")%>>ENGL100</option>
			<option <%=IsSelected(EnglishPlacement, "ENG101H")%>>ENG101H</option>
			<option <%=IsSelected(EnglishPlacement, "ENGL110")%>>ENGL110</option>
			<option <%=IsSelected(EnglishPlacement, "PARAGRAPH")%>>PARAGRAPH</option>
			<option <%=IsSelected(EnglishPlacement, "ESSAY")%>>ESSAY</option>
			</select>
			</div>
            </td>
			<td>
			<div class="selectParent">
			<select size="5" id="ReadingPlacement" name="ReadingPlacement" multiple="multiple" data-allcheckbox="ReadingAll" data-nonecheckbox="ReadingNone">
			<option <%=IsSelected(ReadingPlacement, "DRG091")%>>DRG091</option>
			<option <%=IsSelected(ReadingPlacement, "DRDG091")%>>DRDG091</option>
			<option <%=IsSelected(ReadingPlacement, "DRG092")%>>DRG092</option>
			<option <%=IsSelected(ReadingPlacement, "DRDG092")%>>DRDG092</option>
			<option <%=IsSelected(ReadingPlacement, "READ105")%>>READ105</option>
			</select>
			</div>
			</td>
			<td>
			<div class="selectParent">
			<select size="5" id="ExitPlacement" name="ExitPlacement" multiple="multiple" data-allcheckbox="ExitAll" data-nonecheckbox="ExitNone">
			<option <%=IsSelected(ExitPlacement, "DRG091E")%>>DRG091E</option>
			<option <%=IsSelected(ExitPlacement, "DRDG091E")%>>DRDG091E</option>
			<option <%=IsSelected(ExitPlacement, "DRG092E")%>>DRG092E</option>
			<option <%=IsSelected(ExitPlacement, "DRDG092E")%>>DRDG092E</option>
			<option <%=IsSelected(ExitPlacement, "READ105E")%>>READ105E</option>
			</select>
			</div>
			</td>
			<td>
			<div class="selectParent">
			<select size="4" id="TypePlacement" name="TypePlacement" multiple="multiple" data-allcheckbox="TypeAll" data-nonecheckbox="TypeNone">
			<option <%=IsSelected(TypePlacement, "OFFS100")%>>OFFS100</option>
			<option <%=IsSelected(TypePlacement, "OIT100")%>>OIT100</option>
			<option <%=IsSelected(TypePlacement, "OFFS110")%>>OFFS110</option>
			<option <%=IsSelected(TypePlacement, "OIT110")%>>OIT110</option>
			</select>
			</div>
			</td>
		</tr>
	</table>
	<p><input name="Action" type="submit" value="Generate Report" /></p>
</form>
</div>
<script src="//ajax.googleapis.com/ajax/libs/jquery/2.2.4/jquery.min.js"></script> 
<script src="js/jquery.gototop.js"></script> 
<script>
$(function(){
  "use strict";
  $("#toTop").gototop({ container: "body" });
 
  $(".CheckBoxAll").each(function() {
     var $sbox = $("#" + $(this).data("selectbox"));
     if ($(this).prop("checked")) {
         $sbox.find("option").prop("selected", true);
         $("#" + $sbox.data("nonecheckbox")).prop("checked", false);
     }
  });
  
  $(".CheckBoxNone").each(function() {
     var $sbox = $("#" + $(this).data("selectbox"));
     if ($(this).prop("checked")) {
         $sbox.prop("disabled", true);
         $sbox.val([]);  
         $("#" + $sbox.data("allcheckbox")).prop("checked", false);
     }
  });
  
  $(".CheckBoxNone").click(function(e) {
     var $sbox = $("#" + $(this).data("selectbox"));
     $sbox.prop("disabled", $(this).prop("checked"));
     $sbox.val([]);  
     $("#" + $sbox.data("allcheckbox")).prop("checked", false);
     e.stopPropagation();   
  });
  
  $(".CheckBoxAll").click(function(e) {
     var $sbox = $("#" + $(this).data("selectbox"));
     $sbox.prop("disabled", false);
     $sbox.find("option").prop("selected", $(this).prop("checked"));
     $("#" + $sbox.data("nonecheckbox")).prop("checked", false);
     e.stopPropagation();   
  });
  
  $("#SelectBoxes select").change(function(e) {
    var $allcheckbox = $("#" + $(this).data("allcheckbox"));
    $allcheckbox.prop("checked", $(this).find("option:not(:selected)").length === 0);
    e.stopPropagation();  
  });

});
</script>
</body>
</html>
<%
  closeConnection(conn)
%>