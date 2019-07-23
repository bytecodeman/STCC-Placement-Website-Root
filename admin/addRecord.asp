<% 
  Option Explicit
  On Error Resume Next
%>
<!-- #include file="library/library.asp" -->
<%
  Dim conn, rs, sql, ErrMsg, sqlErr
  Dim fsize, isize, StudentID

  Dim field, bgc
  Dim SSN, LastName, FirstName, Street, City, State, Zipcode, Phone, DOB, Sex, Email

  if IsEmpty(Session("User")) then
    Response.Redirect "login.asp"
  elseif not Session("User")("IsWriter") then 
    Response.Redirect "default.asp"
  end if

  Set conn = openConnection(Application("ConnectionString"))

  ErrMsg = ""
  if Request.ServerVariables("REQUEST_METHOD") = "POST" and Request("MakeChanges") = "Make Changes" then
    SSN = KeepJustDigits(Request("SSN"))
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
    if Request("DOB") <> "" then
      if not IsDate(Request("DOB")) then
        ErrMsg = BuildErrMsg(ErrMsg, "Illegal Birth Date.")
      end if
    end if

    if ErrMsg = "" then
      conn.BeginTrans
      LastName =  Trim(Request("LastName"))
      FirstName = Trim(Request("FirstName"))
      Street = Trim(Request("Street"))
      City = Trim(Request("City"))
      State = Trim(Request("State"))
      Zipcode = Trim(Request("Zipcode"))
      Phone = KeepJustDigits(Trim(Request("Phone")))
      DOB = Trim(Request("DOB"))
      Sex = Trim(Request("Sex"))
      Email = Trim(Request("Email"))    
    
      sql = "INSERT INTO dbo.STUDENTS " & _
            "(SSN, LastName, FirstName, Street, City, State, Zipcode, Phone, Dob, Sex, Email, ProfDate) " & _
            "VALUES (" & _
            "'" & SSN & "', " & _
            "'" & LastName & "', " & _
            "'" & FirstName & "', " & _
            "'" & Street & "', " & _
            "'" & City & "', " & _
            "'" & State & "', " & _
            "'" & Zipcode & "', " & _
            "'" & Phone & "', " & _
            "'" & DOB & "', " & _
            iif(Sex = "", "Null", "'" & Sex & "'") & ", " & _
            "'" & Email & "', " & _
            "'" & Date & "')"
      sqlErr = ExecuteSQL(conn, sql)
      ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
      sql = "Select StudentID From dbo.STUDENTS Where SSN = '" & SSN & "'"
      sqlErr = ExecuteSQLForRs(conn, sql, rs)
      ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
      if ErrMsg = "" then
        StudentID = rs("StudentID")
        rs.Close
        conn.CommitTrans
      else
        conn.RollbackTrans
      end if
      Set rs = Nothing
    end if    
  end If
%>
<!DOCTYPE html>
<html lang="en">

<head>
<meta charset="UTF-8" />
<title><% =SYSTEM_NAME %> - Add New Placement Record</title>
<link rel="stylesheet" href="css/placement.css" />
<link rel="stylesheet" href="css/gototop.css" />
<style>
.firstcol {
	width: 220px;
	height: 30px;
	text-align: right;
	font-weight: bold;
}
.secondcol {
	width: 20px;
	height: 30px;
}
.thirdcol {
	width: 400px;
	height: 30px;
	text-align: left;
}
.submitrow {
	height: 60px;
	text-align: center;
}
.oddcolor {
	background-color: #E0E0E0;
}
</style>
</head>

<body>
<div class="center">
<%
  Call MakeHeader("Add New Record")
  if Request.ServerVariables("REQUEST_METHOD") = "POST" then
    if ErrMsg <> "" then
      Response.Write "<h3 class=""red"">Error(s) Occurred in Record Addition<br />" & ErrMsg & "</h3>" & vbNewLine
    else
      Response.Write "<h3 class=""green"">Successful Record Addition</h3>" & vbNewLine
    end if
  end if
%>
	<p class="bold">
    <a style="margin-right: 10px" href="addRecord.asp">Add Record</a>
<% if StudentID <> "" then %>
	<a style="margin-right: 10px" href="viewRecord.asp?value=<%=StudentID%>">View Just Entered Record</a>
    <a style="margin-right: 10px" href="editRecord.asp?value=<%=StudentID%>">Edit Just Entered Record</a>
    <a style="margin-right: 10px" href="testdetails.asp?value=<%=StudentID%>">Edit Just Entered Test Details</a>
    <a style="margin-right: 10px" href="delRecord.asp?value=<%=StudentID%>">Delete Just Entered Record</a>
<% end if %>
    <a style="margin-right: 10px" href="default.asp">Query List</a>
    <a style="margin-right: 10px" href="default.asp?RefineSearch=true">Refine Search</a>
    <a style="margin-right: 10px" href="default.asp?ResetQuery=true">New Query</a>
    <a href="default.asp?Logout=true">Log Out</a>
    </p>

<hr />
	<div style="width: 640px; margin-left: auto; margin-right: auto">
		<form method="post" style="margin: 5px 0 5px 0">
			<table border="1" cellpadding="0" cellspacing="0">
				<tr>
					<td>
					<table border="0" cellpadding="0" cellspacing="0">
<% 
  sql = "SELECT TOP 1 SSN, LastName AS LastName, FirstName, Street, City, State, Zipcode, Phone, Dob, Sex, Email FROM STUDENTS"
  Set rs = Server.CreateObject("ADODB.Recordset")
  rs.CursorLocation = adUseClient
  rs.Open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
  Set rs.ActiveConnection = Nothing
   
  for each field in rs.Fields
    if bgc = "" then
      bgc = "#E0E0E0"
      Response.Write "<tr class=""oddcolor"">" & vbNewLine
    else
      bgc = ""
      Response.Write "<tr>" & vbNewLine
    end if
    Response.Write "<td class=""firstcol"">" & field.name & "</td>" & vbNewLine
    fsize = rs(field.name).DefinedSize
    isize = fsize
    if fsize > 60 then
      isize = 60
    end if 
    Response.Write "<td class=""secondcol"">&nbsp;</td>" & vbNewLine
    Response.Write "<td class=""thirdcol"">" & "<input type=""text"" size=""" & isize & """ maxlength=""" & fsize & """ name=""" & field.name & """ />" & "</td>" & vbNewLine
    Response.Write "</tr>" & vbNewLine
  next 
  rs.Close
  Set rs = Nothing
%>
						<tr>
							<td colspan="3" class="submitrow">
							<input type="submit" name="MakeChanges" value="Make Changes" /></td>
						</tr>
					</table>
					</td>
				</tr>
			</table>
		</form>
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
<%
    closeConnection(conn)
%>