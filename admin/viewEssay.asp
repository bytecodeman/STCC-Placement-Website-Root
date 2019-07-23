<%
  Option Explicit
  On Error Resume Next
%>
<!-- #include file="library/library.asp" -->
<%
  Dim conn, rs, sql, sqlErr, ErrMsg
  Dim essayrs, placers
  Dim EssayID
  Dim Phase
  Dim SSN, StudentID
  
  if IsEmpty(Session("User")) then
    Response.Redirect "login.asp"
  end if

  Set conn = openConnection(Application("ConnectionString"))
  
  Function GrabSSN(ByVal conn, ByVal EssayID, ByRef rs)
    Dim sql, sqlErr
    sql = "SELECT SSN, Date, Code, Topic, Essay FROM EnglishEssays WHERE ID = " & EssayID
    sqlErr = ExecuteSQLForRS(conn, sql, rs)
    if sqlErr = "" then
      if rs.eof then
        sqlErr = "Cannot Find Essay Number " & EssayID
      end if
    end if 
    GrabSSN = sqlErr    
  End Function
  
  Function GrabPlacementInfo(ByVal conn, ByVal SSN, ByRef rs)
    Dim sql, sqlErr
    sql = "SELECT * FROM [Full Placement Testing Join] WHERE SSN = '" & SSN & "'"
    sqlErr = ExecuteSQLForRs(conn, sql, rs)
    if sqlErr = "" then
      if rs.EOF Then
        sqlErr = "STCC Placement: Cannot Find SSN " & SSN
      end if
    end if
    GrabPlacementInfo = sqlErr
  End Function
  
  Sub CloseDatabase()
    On Error Resume Next
    essayrs.Close
    Set essayrs = Nothing
    placers.Close
    Set placers = Nothing
  End Sub

  Function TypeOfWritingSample(Byval Code)
    TypeOfWritingSample = iif(UCase(Left(Code, 1)) = "P", "PARAGRAPH", "ESSAY")
  End Function
  
  Sub DisplayEssayInfo()
    if IsObject(placers) then
      Response.Write "<h3 style=""margin-top: 10px"">Author: " & placers("lastname") & ", " & placers("firstname") & "</h3>" & vbNewLine
    else
      Response.Write "<h3>ERROR Retrieving Placement Data</h3>" & vbNewLine
    end if
    Response.Write "<hr />" & vbNewLine
    Response.Write "<h3>" & TypeOfWritingSample(essayrs("Code")) & " Topic</h3>" & vbNewLine
    if IsObject(essayrs) then
      Response.Write "<p class=""bold"">" & essayrs("Topic") & "</p>" & vbNewLine
    else
      Response.Write "<p class=""bold"">ERROR in Retrieving Topic</p>" & vbNewLine
    end if
    Response.Write "<hr />" & vbNewLine
    Response.Write "<h3>Writing Sample</h3>" & vbNewLine
    if IsObject(essayrs) then
      Response.Write "<p class=""bold"">" & Replace(essayrs("Essay"), vbNewLine, "<br />") & "</p>" & vbNewLine
    else
      Response.Write "<p class=""bold"">ERROR in Retrieving Essay</p>" & vbNewLine
    end if  
  End Sub
  
  EssayID = InputFilter(Request("EssayID"))
  Phase = ""
  ErrMsg = ""
  if EssayID = "" then
    ErrMsg = BuildErrMsg(ErrMsg, "No Essay Specified.")
    Phase = "0"
  else
    sqlErr = GrabSSN(conn, EssayID, essayrs)
    if sqlErr <> "" then
      ErrMsg = BuildErrMsg(ErrMsg, "ERROR Retrieving Essay: " & sqlErr)
      Phase = "0"
    else
      SSN = essayrs("SSN")
      sqlErr = GrabPlacementInfo(conn, SSN, placers)
      StudentID = placers("StudentID")
      if sqlErr <> "" then
        ErrMsg = BuildErrMsg(ErrMsg, "ERROR Retrieving Placement Info: " & sqlErr)
        Phase = "0"
      end if
    end if
  end if
%>
<!DOCTYPE html>
<html lang="en">

<head>
<meta charset="UTF-8" />
<title>STCC Writing Sample</title>
<link rel="stylesheet" href="css/placement.css" />
<link rel="stylesheet" href="css/gototop.css" />
</head>

<body class="center">
<%
  Call MakeHeader("Writing Sample")
  if ErrMsg <> "" then
     Response.Write "<h3 class=""red center"">Error(s) Occurred in Essay View<br />" & ErrMsg & "</h3>" & vbNewLine
  end if
%>
<p class="bold hideElement">
<% if Session("User")("IsWriter") and not IsEmpty(StudentID) then %>
  <a style="margin-right: 10px" href="addRecord.asp">Add Record</a>
  <a style="margin-right: 10px" href="viewRecord.asp?value=<%=StudentID%>">View Record</a>
  <a style="margin-right: 10px" href="editRecord.asp?value=<%=StudentID%>">Edit Record</a>
  <a style="margin-right: 10px" href="testdetails.asp?value=<%=StudentID%>">Edit Test Details</a>
  <a style="margin-right: 10px" href="delRecord.asp?value=<%=StudentID%>">Delete Record</a>
<% end if %>
  <a style="margin-right: 10px" href="default.asp">Query List</a>
  <a style="margin-right: 10px" href="default.asp?RefineSearch=true">Refine Search</a>
  <a style="margin-right: 10px" href="default.asp?ResetQuery=true">New Query</a>
  <a href="default.asp?Logout=true">Log Out</a>
</p>
<hr class="hideElement" />

<div style="width: 800px; margin-left: auto; margin-right: auto;">
	<div style="text-align: left">
<%
  if Phase = "" then
    Call DisplayEssayInfo()
  end if
  Call CloseDatabase()
%>
  <form class="hideElement">
  <hr />
	<p><input name="button" type="button" onclick="history.back();" value="Go Back" /></p>
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