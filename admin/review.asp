<%
  Option Explicit
  On Error Resume Next
%>
<!-- #include file="library/library.asp" -->
<%
  Dim conn, rs, sql, sqlErr, ErrMsg
  Dim essayrs, placers
  Dim EssayID, ReaderID, Evaluation, Submit, englPlacement
  Dim PageTitle
  Dim Phase
  Dim SSN
  
  Set conn = openConnection(Application("ConnectionString"))

  Function GrabSSN(ByVal conn, ByVal EssayID, ByRef rs)
    Dim sql, sqlErr, ErrMsg
    sql = "SELECT SSN, Date, Code, Topic, Essay FROM EnglishEssays WHERE ID = " & EssayID
    sqlErr = ExecuteSQLForRS(conn, sql, rs)
    ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
    if sqlErr = "" then
      if rs.eof then
        ErrMsg = BuildErrMsg(ErrMsg, "Cannot Find Essay Number " & EssayID)
      end if
    end if 
    GrabSSN = ErrMsg
  End Function
  
  Function GrabPlacementInfo(ByVal conn, ByVal SSN, ByRef rs)
    Dim sql, sqlErr, ErrMsg
    sql = "SELECT * FROM [Full Placement Testing Join] WHERE SSN = '" & SSN & "'"
    sqlErr = ExecuteSQLForRs(conn, sql, rs)
    ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
    if sqlErr = "" then
      if rs.EOF Then
        ErrMsg = BuildErrMsg(ErrMsg, "STCC Placement: Cannot Find SSN " & SSN)
      end if
    end if
    GrabPlacementInfo = ErrMsg
  End Function
  
  Function ReaderCounts(ByVal conn, ByVal EssayID, ByVal Condition, ByRef sqlErr)
    Dim sql, rs
    sql = "SELECT Count(*) AS Count FROM ContactReaders WHERE EssayID = " & EssayID & " AND " & Condition 
    sqlErr = ExecuteSQLForRS(conn, sql, rs)
    if sqlErr = "" then
      ReaderCounts = rs("Count")
    else
      ReaderCounts = -1
    end if
    rs.Close
    Set rs = Nothing
  End Function
  
  Sub CloseRecordSets()
    On Error Resume Next
    essayrs.Close
    Set essayrs = Nothing
    placers.Close
    Set placers = Nothing
  End Sub
  
  Function EnglishPlacement(ByVal rs)
    if IsNull(rs("EnglPlacement")) then
      EnglishPlacement = "NO PLACEMENT TAKEN"
    else
      EnglishPlacement = FormatPlacement(rs("EnglPlacement")) & " " & FormatDateTime(rs("EnglDate")) & " "  & ZPadStr(CInt(rs("EnglScore")), 3)
    end if
  End Function
  
  Function ReadingPlacement(ByVal rs)
    If IsNull(rs("ReadPlacement")) Then
      ReadingPlacement = "NO PLACEMENT TAKEN"
    Else
      ReadingPlacement = FormatPlacement(rs("ReadPlacement")) & " " & FormatDateTime(rs("ReadDate")) & " " & ZPadStr(CInt(rs("ReadScore")), 3)
    End If
  End Function
  
  Function SetEvaluationRadio(ByVal value)
    SetEvaluationRadio = iif(Request("Evaluation") = value, "checked=""checked""", "")
  End Function
  
  Function CheckIfReaderRecordExists(ByVal conn, ByVal EssayID, ByVal ReaderID)
    Dim sql, sqlErr, rs, ErrMsg
    sql = "SELECT ReaderID FROM ContactReaders "
    sql = sql & "WHERE EssayID = " & EssayID & " AND ReaderID = " & ReaderID 
    sqlErr = ExecuteSQLForRS(conn, sql, rs)
    ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
    if sqlErr = "" then
      if rs.Eof then
        ErrMsg = BuildErrMsg(ErrMsg, "Reader Record Not Found.")
      end if
    end if
    rs.Close
    Set rs = Nothing
    CheckIfReaderRecordExists = ErrMsg
  End Function

  Function EssayEvaluation(ByVal conn, ByVal placers, ByVal Placement, Byval EssayID, ByVal ReaderID)
    Dim sql, sqlErr, ErrMsg, alreadyPlaced
    Dim readersrs
    Dim englPlacement
    Dim ENG101Count, DWT099Count, EvalCount
    
    conn.BeginTrans
    sql = "UPDATE ContactReaders SET "
    sql = sql & "ReaderScoreDate = '" & Date & "', "
    sql = sql & "ReaderPlacement = '" & Placement & "' "
    sql = sql & "WHERE EssayID = " & EssayID & " AND ReaderID = " & ReaderID
    sqlErr = ExecuteSQL(conn, sql)
    ErrMsg = BuildErrMsg(ErrMsg, sqlErr)

    alreadyPlaced = false
    englPlacement = placers("EnglPlacement")
    if englPlacement <> "ESSAY" and englPlacement <> "PARAGRAPH" then
      alreadyPlaced = true
      ErrMsg = BuildErrMsg(ErrMsg, "Placement Accessment has already been made: " & englPlacement)
    end if
    
    if ErrMsg = "" and not alreadyPlaced then
      ENG101Count = ReaderCounts(conn, EssayID, "ReaderPlacement = 'ENG101'", sqlErr)
      ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
      DWT099Count = ReaderCounts(conn, EssayID, "ReaderPlacement = 'DWT099'", sqlErr)
      ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
      EvalCount    = ReaderCounts(conn, EssayID, "ReaderPlacement Is Not Null", sqlErr)
      ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
      if ENG101Count >= 2 then
        sql = "UPDATE [Full Placement Testing Join] SET EnglPlacement = 'ENG101', EnglStatus = 1 WHERE SSN = '" & placers("SSN") & "'"
        sqlErr = ExecuteSQL(conn, sql)
        ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
      elseif DWT099Count >= 2 then
        sql = "UPDATE [Full Placement Testing Join] SET EnglPlacement = 'DWT099', EnglStatus = 1 WHERE SSN = '" & placers("SSN") & "'"
        sqlErr = ExecuteSQL(conn, sql)
        ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
      elseif EvalCount >= 2 then
        sql = "Select ID, FullName, Email FROM (SELECT ReaderID FROM ContactReaders WHERE EssayID = " & EssayID & ") A RIGHT OUTER JOIN " & _
              "(SELECT ID, Code, FullName, Email From EssayReaders where Code is not Null) B ON A.ReaderID = B.ID WHERE A.ReaderID Is Null"
        sqlErr = ExecuteSQLForRs(conn, sql, readersrs)
        ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
        if ErrMsg = "" then
          if not readersrs.eof then
            sql = "INSERT INTO ContactReaders (EssayID, SSN, ReaderID) VALUES (" & EssayID & ", '" & placers("SSN") & "', " & readersrs("ID") & ")"
            sqlErr = ExecuteSQL(conn, sql)
            ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
            if ErrMsg = "" then
              Call ShootEmail(placers, readersrs, EssayID)
            end if
          end if
        end if
      end if    
    end if
    
    if ErrMsg = "" then
      conn.CommitTrans
    else
      conn.RollbackTrans
    end if
    EssayEvaluation = ErrMsg
  End Function
  
  Function TypeOfWritingSample(Byval Code)
    TypeOfWritingSample = iif(UCase(Left(Code, 1)) = "P", "PARAGRAPH", "ESSAY")
  End Function
  
  Sub DisplayEssayInfo()
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
  
  Sub DisplayPlacementInfo()
    Response.Write "<hr />" & vbNewLine
    Response.Write "<h3>English/Reading Placement Testing Results</h3>" & vbNewLine
    if IsObject(placers) then
      Response.Write "<p class=""bold"">NAME: " & placers("lastname") & ", " & placers("firstname") & "<br />" & vbNewLine
      Response.Write "ENGLISH: " & EnglishPlacement(placers) & "<br />" & vbNewLine
      Response.Write "READING: " & ReadingPlacement(placers) & "</p>" & vbNewLine
    else
      Response.Write "<p class=""bold"">Error Retrieving Placement Results</p>" & vbNewLine
    end if
  End Sub

  EssayID = InputFilter(Request("EssayID"))
  ReaderID = InputFilter(Request("ReaderID"))
  Phase = ""
  PageTitle = "STCC " & iif(ReaderID = "", "Writing Sample", "Writing Sample Evaluation")
  
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
      if sqlErr <> "" then
        ErrMsg = BuildErrMsg(ErrMsg, "ERROR Retrieving Placement Info: " & sqlErr)
        Phase = "0"
      end if
    end if
    if ErrMsg = "" then
      if ReaderID = "" then
        Phase = "3"
      else
        sqlErr = CheckIfReaderRecordExists(conn, EssayID, ReaderID)
        if sqlErr <> "" then
          ErrMsg = BuildErrMsg(ErrMsg, sqlErr)
          Phase = "4"
        end if
        englPlacement = placers("EnglPlacement")
        if englPlacement <> "ESSAY" and englPlacement <> "PARAGRAPH" then
          ErrMsg = BuildErrMsg(ErrMsg, "Writing Sample Already Evaluated as " & englPlacement)
          Phase = "4"
        end if
      end if
    end if 
  end if
  if ErrMsg = "" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
    Evaluation = InputFilter(Request("Evaluation"))
    Submit = InputFilter(Request("Submit"))
    if Submit = "Submit" then
      if Evaluation = "" then
        ErrMsg = BuildErrMsg(ErrMsg, "You Must Select a Placement Level.")
        Phase = ""
      else
        Phase = "1"
      end if
    elseif Submit = "Yes" then
      sqlErr = EssayEvaluation(conn, placers, Evaluation, EssayID, ReaderID)
      if sqlErr <> "" then
        ErrMsg = BuildErrMsg(ErrMsg, "ERROR Posting Essay Evaluation: " & sqlErr)
        Phase = "1"
      else
        sqlErr = GrabPlacementInfo(conn, SSN, placers)
        if sqlErr <> "" then
          ErrMsg = BuildErrMsg(ErrMsg, "ERROR Posting Essay Evaluation: " & sqlErr)
          Phase = "1"
        else
          Phase = "2"
        end if
      end if
    end if
  end if
%>
<!DOCTYPE html>
<html lang="en">

<head>
<meta charset="UTF-8" />
<title><%=PageTitle%></title>
<link rel="stylesheet" href="css/placement.css" />
<link rel="stylesheet" href="css/gototop.css" />
</head>

<body class="center">

<div style="width: 800px; margin-left: auto; margin-right: auto;">
	<div class="hideElement">
	<img alt="<%=PageTitle%>" src="img/<%=iif(ReaderID = "", "wrtngsampl.jpg", "wrtngsampeval.jpg")%>" width="505" height="54" />
	</div>
<%
  if ErrMsg <> "" then
    Response.Write "<h3 class=""red center"">Error(s) Occurred in Writing Sample Evaluation<br />" & ErrMsg & "</h3>" & vbNewLine
  end if
%>
	<div style="text-align: left">
	<form method="post">
<%
  if Phase = "" then
  	Call DisplayPlacementInfo()
    Call DisplayEssayInfo()
%>
  <hr />
  <h3>Select Placement Level</h3>
  <p class="bold">
	<input type="radio" name="Evaluation" id="DWT099" accesskey="d" tabindex="1" value="DWT099" <%=SetEvaluationRadio("DWT099")%> /><label for="DWT099" style="margin-right: 25px">DWT099</label>
	<!-- <input type="radio" name="Evaluation" id="REFERRAL" accesskey="r" value="REFERRAL" <%=SetEvaluationRadio("REFERRAL")%> /><label for="REFERRAL" style="margin-right: 25px">REFERRAL</label> -->
	<input type="radio" name="Evaluation" id="ENG101" accesskey="e" value="ENG101" <%=SetEvaluationRadio("ENG101")%> /><label for="ENG101">ENG101</label></p>
	<p>
	<input type="submit" name="Submit" accesskey="s" tabindex="2" value="Submit" />
	<input type="hidden" name="EssayID"  value="<%=EssayID%>" />
	<input type="hidden" name="ReaderID" value="<%=ReaderID%>" /></p>
<%
  elseif Phase = "0" then
    ' Show Nothing -- Error State
  elseif Phase = "1" then
%>
	<h3>Are you sure you want to score this Writing Sample as: <%=Evaluation%></h3>
	<p>If you click <strong>Yes</strong>, you will not be able to change your evaluation later.</p>
	<p>
	<input type="submit" name="Submit" value="No" style="margin-right: 25px" accesskey="n" tabindex="1" />
	<input type="submit" name="Submit" value="Yes" accesskey="y" tabindex="2" />
	<input type="hidden" name="EssayID"  value="<%=EssayID%>"  />
	<input type="hidden" name="ReaderID" value="<%=ReaderID%>" />
	<input type="hidden" name="Evaluation" value="<%=Evaluation%>" /></p>
<% 
	Call DisplayPlacementInfo()
	Call DisplayEssayInfo()
  elseif Phase = "2" then
    Response.Write "<h3 class=""green center"">Writing Sample Evaluation of " & Evaluation & " Successfully Submitted" & "</h3>" & vbNewLine
    Call DisplayPlacementInfo()
    Call DisplayEssayInfo()
  elseif Phase = "3" then
    Call DisplayEssayInfo()
  elseif Phase = "4" then
    Call DisplayPlacementInfo()
    Call DisplayEssayInfo()
  end if
  Call CloseRecordSets()
%>
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
