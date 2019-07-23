<%
  Option Explicit
  On Error Resume Next
%>
<!-- #include file="library/library.asp" -->
<%
  Dim conn, rs, sql, tmprs, rcount
  Dim strQuery, ErrMsg, sqlErr, fvalue, tmpError, RequestSSN
  Dim English
  Dim fsize, isize
  Dim field, Fields(11), SSN, StudentID
  Dim columns
  
  if IsEmpty(Session("User")) then
    Response.Redirect "login.asp"
  elseif not Session("User")("IsWriter") then 
    Response.Redirect "default.asp"
  elseif Request("value") = "" then
    Response.Redirect "default.asp?ResetQuery=true"
  else
    StudentID = Request("value")
  end if

  Set conn = openConnection(Application("ConnectionString"))
  
  Function UpdatePlacementResult(ByVal rsField, ByVal Table, ByVal Placement, ByVal studentField)
    Dim temprs
    Dim sqlErr, ErrMsg
    Dim sqlQuery
    if IsNull(rs(rsField)) and Request(rsField) = "ON" then
      strQuery = "SELECT ID From " & Table & " WHERE SSN = '" & SSN & "'"
      sqlErr = ExecuteSQLForRS(conn, strQuery, tmprs)
      if sqlErr <> "" then
        ErrMsg = BuildErrMsg(ErrMsg, Placement & " Placement Set Error: " & sqlErr)
      else
        rcount = tmprs.RecordCount
        if rcount <> 1 then
          ErrMsg = BuildErrMsg(ErrMsg, "Must have 1 " & Placement & " Placement to Automatically Reinstate, " & rcount & " found")
        else
          strQuery = "UPDATE dbo.Students SET " & studentField & " = " & tmprs("ID") & " WHERE SSN = '" & SSN & "'"
          sqlErr = ExecuteSQL(conn, strQuery)
          if sqlErr <> "" then
            ErrMsg = BuildErrMsg(ErrMsg, Placement & " Placement Set Error: " & sqlErr)
          end if
        end if  
        tmprs.Close
      end if 
      Set tmprs = Nothing
    elseif (not IsNull(rs(rsField))) and Request(rsField) <> "ON" then
      strQuery = "UPDATE dbo.Students SET " & studentField & " = Null WHERE SSN = '" & SSN & "'"
      sqlErr = ExecuteSQL(conn, strQuery)
      if sqlErr <> "" then
        ErrMsg = BuildErrMsg(ErrMsg, Placement & " Placement Reset Error: " & sqlErr)
      end if
    end if
    UpdatePlacementResult = ErrMsg
  End Function
  
  Function UpdateEssaySampleResult()
    Dim temprs
    Dim sqlErr, ErrMsg
    Dim sqlQuery
    if IsNull(rs("EssayID")) and Request("EssayID") = "ON" then
      strQuery = "SELECT ID From EnglishEssays WHERE SSN = '" & SSN & "'"
      sqlErr = ExecuteSQLForRS(conn, strQuery, tmprs)
      if sqlErr <> "" then
        ErrMsg = BuildErrMsg(ErrMsg, "Essay Sample Set Error: " & sqlErr)
      else
        rcount = tmprs.RecordCount
        if rcount <> 1 then
          ErrMsg = BuildErrMsg(ErrMsg, "Must have 1 Essay Sample to Automatically Reinstate, " & rcount & " found")
        else
          strQuery = "UPDATE dbo.[Full Placement Testing Join] SET ESSAYID = " & tmprs("ID") & " WHERE SSN = '" & SSN & "'"
          sqlErr = ExecuteSQL(conn, strQuery)
          if sqlErr <> "" then
            ErrMsg = BuildErrMsg(ErrMsg, "Essay Sample Set Error: " & sqlErr)
          end if
        end if  
        tmprs.Close
      end if 
      Set tmprs = Nothing
    elseif (not IsNull(rs("EssayID"))) and Request("EssayID") <> "ON" then
      strQuery = "UPDATE dbo.[Full Placement Testing Join] SET ESSAYID = Null WHERE SSN = '" & SSN & "'"
      sqlErr = ExecuteSQL(conn, strQuery)
      if sqlErr <> "" then
        ErrMsg = BuildErrMsg(ErrMsg, "Essay Sample Reset Error: " & sqlErr)
      end if
    end if
    UpdateEssaySampleResult = ErrMsg
  End Function
  
  Class FieldClass
    public name
    public caption
    public essential
  End Class
  
  Set field = New FieldClass
  field.name = "LastName"
  field.caption = "Last Name"
  field.essential = true
  Set Fields(0) = field
  
  Set field = New FieldClass
  field.name = "FirstName"
  field.caption = "First Name"
  field.essential = true
  Set Fields(1) = field
  
  Set field = New FieldClass
  field.name = "Street"
  field.caption = ""
  field.essential = false
  Set Fields(2) = field

  Set field = New FieldClass
  field.name = "City"
  field.caption = ""
  field.essential = false
  Set Fields(3) = field

  Set field = New FieldClass
  field.name = "State"
  field.caption = ""
  field.essential = false
  Set Fields(4) = field

  Set field = New FieldClass
  field.name = "Zipcode"
  field.caption = "Zip Code"
  field.essential = false
  Set Fields(5) = field

  Set field = New FieldClass
  field.name = "Phone"
  field.caption = ""
  field.essential = false
  Set Fields(6) = field

  Set field = New FieldClass
  field.name = "DOB"
  field.caption = "Birth Date"
  field.essential = false
  Set Fields(7) = field

  Set field = New FieldClass
  field.name = "Sex"
  field.caption = ""
  field.essential = false
  Set Fields(8) = field

  Set field = New FieldClass
  field.name = "Email"
  field.caption = ""
  field.essential = false
  Set Fields(9) = field
  
  Set field = New FieldClass
  field.name = "ProfDate"
  field.caption = "Record Date"
  field.essential = false
  Set Fields(10) = field

  Set field = New FieldClass
  field.name = "IPAddr"
  field.caption = "IP Address"
  field.essential = false
  Set Fields(11) = field

  Set field = Nothing
 
  SSN = TranslateID2SSN(conn, StudentID)
  if SSN = "" then
    ErrMsg = BuildErrMsg(ErrMsg, "Record Cannot Be Located")
  else
    sql = "SELECT dbo.DecodeSSN(SSN) As SSN, LastName, FirstName, Street, City, State, Zipcode, Phone, " & _
          "Dob, Sex, Email, ProfDate, StudentIPAddr AS IPAddr, EnglPlaceID, MathPlacement, EnglPlacement, EssayID, ReadPlacement, " & _
          "ExitPlacement, TypePlacement FROM dbo.[Full Placement Testing Join] WHERE StudentID = " & StudentID
    sqlErr = ExecuteSQLForRS(conn, sql, rs)
    ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
  end if
  
  if Request.ServerVariables("REQUEST_METHOD") = "POST" and Request("MakeChanges") = "Make Changes" then
    RequestSSN = KeepJustDigits(Request("SSN"))
    if Len(RequestSSN) = 0 Or Len(RequestSSN) = 8 Then
      ErrMsg = BuildErrMsg(ErrMsg, "Illegal ID/SSN")
    else
      RequestSSN = EncodeSSN(RequestSSN)
    End if
    if Request("DOB") <> "" then
      if not IsDate(Request("DOB")) then
        ErrMsg = BuildErrMsg(ErrMsg, "Illegal Birth Date.")
      end if
    end if
    English = UCase(Trim(Request("EnglCourse")))
    if English <> "" And InStr(".DWT099.ENG101.ENG101H.ESSAY", "." & English) = 0 then
      ErrMsg = BuildErrMsg(ErrMsg, "Illegal English Placement")
    end if
     
    if ErrMsg = "" then
      conn.BeginTrans
      ' Phase 1 SSN Change
      if RequestSSN <> SSN then
        strQuery = "UPDATE dbo.STUDENTS SET SSN = '" & RequestSSN & "' WHERE SSN = '" & SSN & "'"
        sqlErr = ExecuteSQL(conn, strQuery)
        ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
        if ErrMsg = "" then
          strQuery = ""
          strQuery = strQuery & "UPDATE dbo.MathPlacement        SET SSN = '" & RequestSSN & "' WHERE SSN = '" & SSN & "' "
          strQuery = strQuery & "UPDATE dbo.EnglishPlacement     SET SSN = '" & RequestSSN & "' WHERE SSN = '" & SSN & "' "
          strQuery = strQuery & "UPDATE dbo.ReadingPlacement     SET SSN = '" & RequestSSN & "' WHERE SSN = '" & SSN & "' " 
          strQuery = strQuery & "UPDATE dbo.EnglishEssays        SET SSN = '" & RequestSSN & "' WHERE SSN = '" & SSN & "' "
          strQuery = strQuery & "UPDATE dbo.TypingPlacement      SET SSN = '" & RequestSSN & "' WHERE SSN = '" & SSN & "' " 
          strQuery = strQuery & "UPDATE dbo.ReadingExitPlacement SET SSN = '" & RequestSSN & "' WHERE SSN = '" & SSN & "' "
          strQuery = strQuery & "UPDATE dbo.ContactReaders       SET SSN = '" & RequestSSN & "' WHERE SSN = '" & SSN & "' "
          sqlErr = ExecuteSQL(conn, strQuery)
          ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
          if ErrMsg = "" then 
            SSN = RequestSSN
          end if
        end if
      end if
      
      ' Phase 2 Normal Field Mods
      if ErrMsg = "" then
        strQuery = ""
        for each field in Fields
          if rs(field.name).type = adDBTimeStamp then
            if rs(field.name) <> CDate(Request(field.name)) then
              if strQuery <> "" then 
                strQuery = strQuery & ", "
              end if
              strQuery = strQuery & "[" & field.name & "] = '" & Request(field.name) & "'"
            end if
          else
            if rs(field.name) <> Request(field.name) then
              if strQuery <> "" then 
                strQuery = strQuery & ", "
              end if
              if field.name = "Phone" then
                strQuery = strQuery & "[" & field.name & "] = '" & KeepJustDigits(Request(field.name)) & "'"
              else
                strQuery = strQuery & "[" & field.name & "] = '" & Replace(Trim(Request(field.name)), "'", "''") & "'"
              end if
            end if
          end if    
        next
        if strQuery <> "" then
          strQuery = "UPDATE dbo.Students SET " & strQuery & " WHERE SSN = '" & SSN & "'"
          sqlErr = ExecuteSQL(conn, strQuery)
          ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
        end if
      end if
      
      ' Phase 3 English Placement ESSAY to Something Else
      if ErrMsg = "" then
        if English <> "" And English <> "ESSAY" And English <> UCase(Trim(rs("EnglPlacement"))) then
          strQuery = "UPDATE dbo.EnglishPlacement SET Placement = '" & English & "', STATUS = 1 WHERE ID = " & rs("EnglPlaceID")
          sqlErr = ExecuteSQL(conn, strQuery)
          ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
        end if
      end if
      
      ' Phase 4 Placement Result Mods
      if ErrMsg = "" then
        sqlErr = UpdatePlacementResult("MathPlacement", "MathPlacement", "Math", "MathPlaceID")
        ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
        sqlErr = UpdatePlacementResult("EnglPlacement", "EnglishPlacement", "English", "EnglPlaceID")
        ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
        sqlErr = UpdateEssaySampleResult()
        ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
        sqlErr = UpdatePlacementResult("ReadPlacement", "ReadingPlacement", "Reading", "ReadPlaceID")
        ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
        sqlErr = UpdatePlacementResult("TypePlacement", "TypingPlacement", "Keyboarding", "TypePlaceID")
        ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
        sqlErr = UpdatePlacementResult("ExitPlacement", "ReadingExitPlacement", "Reading Exit", "ExitPlaceID")
        ErrMsg = BuildErrMsg(ErrMsg, sqlErr)          
      end if
      
      if Err.Number <> 0 then
        ErrMsg = BuildErrMsg(ErrMsg, "ERROR #" & Err.Number & " " & Err.Description & " Source: " & Err.Source)
      end if
      
      if ErrMsg = "" then
        conn.CommitTrans
        sqlErr = ExecuteSQLForRS(conn, sql, rs)
        ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
      else
        conn.RollbackTrans
      end if
    end if
  end if
%>
<!DOCTYPE html>
<html lang="en">

<head>
<meta charset="UTF-8" />
<title><% =SYSTEM_NAME %> - Edit Record</title>
<link rel="stylesheet" href="css/placement.css" />
<link rel="stylesheet" href="css/gototop.css" />
<style>
table {	
	text-align:left;
	margin-left:auto;
	margin-right:auto;
}

th {
	text-align: center;
}

th.headerwidth {
	width: 85px;
}
.detailrow {
	text-align: center !important;
	height: 50px !important;
}
#personal {
	margin-bottom: 20px;
}
#personal td {
	height: 25px;
}

#personal td.firstrow {
	width: 120px;
	font-weight:bold;
}
#personal td.secondrow {
	width: 400px;
}
</style>
</head>

<body>
<div class="center">
<%
  Call MakeHeader("Edit Record")
  if ErrMsg <> "" then
    Response.Write "<h3 class=""red"">Error(s) Occurred in Record Edit<br />" & ErrMsg & "</h3>" & vbNewLine
  elseif Request.ServerVariables("REQUEST_METHOD") = "POST" then
     Response.Write "<h3 class=""green"">Successful Submission!!!" & "</h3>" & vbNewLine
  end if
%>
	<p class="bold">
    <a style="margin-right: 10px" href="addRecord.asp">Add Record</a>
	<a style="margin-right: 10px" href="viewRecord.asp?value=<%=StudentID%>">View Record</a>
    <a style="margin-right: 10px" href="editRecord.asp?value=<%=StudentID%>">Edit Record</a>
    <a style="margin-right: 10px" href="testdetails.asp?value=<%=StudentID%>">Edit Test Details</a>
    <a style="margin-right: 10px" href="delRecord.asp?value=<%=StudentID%>">Delete Record</a>
    <a style="margin-right: 10px" href="default.asp">Query List</a>
    <a style="margin-right: 10px" href="default.asp?RefineSearch=true">Refine Search</a>
    <a style="margin-right: 10px" href="default.asp?ResetQuery=true">New Query</a>
    <a href="default.asp?Logout=true">Log Out</a>
    </p>
	<hr />
<%
    if not rs.eof then
      columns = 6
      if rs("EnglPlacement") = "ESSAY" then columns = columns + 1
%>
<form method="post">
<p class="bold">Student ID/SSN: <input type="text" name="SSN" size="20" value="<%=rs("SSN")%>" />&nbsp; 
Last Name:
<input type="text" name="LastName" size="20" value="<%=rs("LastName")%>" />&nbsp; 
First Name:
<input type="text" name="FirstName" size="20" value="<%=rs("FirstName")%>" /></p>

<table id="examstatus" border="1" cellpadding="2" cellspacing="2">
  <tr>
    <th colspan="<%=columns%>">STCC Placement Exam Status</th>
  </tr>
  <tr>
    <th class="headerwidth">Math</th>
    <th class="headerwidth">English</th>
    <% if rs("EnglPlacement") = "ESSAY" then %>
    <th>English Placement</th>
    <% end if %>
    <th class="headerwidth">Essay</th>
    <th class="headerwidth">Reading</th>
    <th class="headerwidth">Keyboarding</th>
    <th class="headerwidth">Reading Exit</th>
  </tr>
  <tr>
    <td class="center"><input type="checkbox" name="MathPlacement" value="ON" <%=iif(IsNull(rs("MathPlacement")), "", "checked=""checked""")%> /></td>
    <td class="center"><input type="checkbox" name="EnglPlacement" value="ON" <%=iif(IsNull(rs("EnglPlacement")), "", "checked=""checked""")%> /></td>
    <% if rs("EnglPlacement") = "ESSAY" then %>
    <td class="center"><input type="text" name="EnglCourse" size="12" maxlength="10" value="<%=rs("EnglPlacement")%>" /></td>
    <% end if %>
    <td class="center"><input type="checkbox" name="EssayID"       value="ON" <%=iif(IsNull(rs("EssayID")), "", "checked=""checked""")%> /></td>
    <td class="center"><input type="checkbox" name="ReadPlacement" value="ON" <%=iif(IsNull(rs("ReadPlacement")), "", "checked=""checked""")%> /></td>
    <td class="center"><input type="checkbox" name="TypePlacement" value="ON" <%=iif(IsNull(rs("TypePlacement")), "", "checked=""checked""")%> /></td>
    <td class="center"><input type="checkbox" name="ExitPlacement" value="ON" <%=iif(IsNull(rs("ExitPlacement")), "", "checked=""checked""")%> /></td>
  </tr>
  <tr>
    <td colspan="<%=columns%>" class="detailrow">
    <input type="submit" name="MakeChanges" value="Make Changes" /></td>
  </tr>

  </table>

<hr style="margin:20px auto" />
<table id="personal" border="1" cellpadding="2" cellspacing="2">
<%  
  for each field in Fields
    if not field.essential then
      Response.Write "<tr>" & vbNewLine
      if field.caption = "" then
        field.caption = field.name
      end if
      Response.Write "<td class=""firstrow"">" & field.caption & "</td>" & vbNewLine
      fsize = rs(field.name).DefinedSize
      isize = fsize
      if fsize > 60 then
        isize = 60
      end if 
      if field.name = "IPAddr" then
        Response.Write "<td class=""secondrow"">" & "<input type=""hidden"" name=""" & field.name & """ value=""*" & Request.ServerVariables("REMOTE_ADDR") & """ />" & "<input type=""text"" size=""" & isize & """ maxlength=""" & fsize & """ name=""" & field.name & """ value=""" & Trim(rs(field.name)) & """ disabled=""disabled""  />" & "</td>" & vbNewLine
      else
        Response.Write "<td class=""secondrow"">" & "<input type=""text"" size=""" & isize & """ maxlength=""" & fsize & """ name=""" & field.name & """ value=""" & Trim(rs(field.name)) & """ />" & "</td>" & vbNewLine
      end if
        Response.Write "</tr>" & vbNewLine
    end if
  next 
%>
<tr>
<td colspan="2" class="detailrow">
<input type="submit" name="MakeChanges" value="Make Changes" /></td>
</tr>
</table>
</form>
<%
      rs.close
      Set rs = Nothing
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