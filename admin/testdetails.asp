<%
  Option Explicit
  'On Error Resume Next
%>
<!-- #include file="library/library.asp" -->
<!-- #include file="library/processXML.asp" -->
<!-- #include file="library/calculatePlacements.asp" -->
<%  
  Dim conn, rs, sql
  Dim strQuery, eStr
  Dim SSN, StudentID
  Dim readers, temprs
  Dim TestType, Operation, ID, EssayID
  Dim WPScore
  
  Const INIFileLocation = "/GlobalTestingShell.xml"

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
  
  SSN = TranslateID2SSN(conn, StudentID)
  eStr = ""
  if SSN = "" then
    eStr = "Record Cannot Be Located"
  elseif Request.ServerVariables("REQUEST_METHOD") = "POST" then
    TestType = Request("TestType")
    Operation = UCase(Trim(Request("Operation")))
    ID = UCase(Trim(Request("ID")))
    if ID = "" then
      eStr = "No Item Specified for " & Operation & " in " & TestType 
    else
      conn.BeginTrans
      if TestType = "Math Placement" then
        eStr = MathPlacement(Operation, ID)
      elseif TestType = "English Placement" then
        eStr = EnglishPlacement(Operation, ID)
      elseif TestType = "Essay Samples" then
        eStr = EssaySamples(Operation, ID)
      elseif TestType = "Contact Readers" then
        eStr = ContactReaders(Operation, ID)
      elseif TestType = "Reading Placement" then
        eStr = ReadingPlacement(Operation, ID)
      elseif TestType = "Keyboarding Placement" then
        eStr = TypingPlacement(Operation, ID)
      elseif TestType = "Reading Exit Placement" then
        eStr = ExitPlacement(Operation, ID)
      end if
      if Err.Number <> 0 then
        eStr = BuildErrMsg(estr, "ERROR #" & Err.Number & " " & Err.Description & " Source: " & Err.Source)
      end if
      if eStr = "" then
        conn.CommitTrans
      else
        conn.RollbackTrans
      end if
    end if    
  end if
  
  '=================================================================================================================

  Function MathPlacement(ByVal Operation, ByVal ID)
    Dim rs, takeid
    Dim sql, estr, sqlErr
    Dim TestDate, ArithScore, AlgScore, CollegeScore, Placement, IPAddr
    estr = ""
    if Operation = "SELECT" then
      if ID <> "NEW" then
        sql = "SET NOCOUNT ON "
        sql = sql & "UPDATE STUDENTS SET MathPlaceID = " & ID & " WHERE SSN = '" & SSN & "' "
        sql = sql & "UPDATE MATHPLACEMENT SET Status = 1 WHERE ID = " & ID
        sqlErr = ExecuteSQL(conn, sql)
        if sqlErr <> "" then
          estr = "Math Placement Test Detail Set Error: " & sqlErr
        end if
      else
        TestDate = Trim(Request("Date"))
        IPAddr = Trim(Request("IPAddr"))
        ArithScore = Trim(Request("ArithScore"))
        if isBlank(ArithScore) then
          ArithScore = 0
        end if
        AlgScore = Trim(Request("AlgScore"))
        if isBlank(AlgScore) then
          AlgScore = 0
        end if
        CollegeScore = Trim(Request("CollegeScore"))
        if isBlank(CollegeScore) then
          CollegeScore = 0
        end if
        if IsNumeric(ArithScore) and IsNumeric(AlgScore) and IsNumeric(CollegeScore) then
          if ArithScore = 0 and AlgScore = 0 and CollegeScore = 0 then
            sql = "SET NOCOUNT ON " 
            sql = sql & "UPDATE STUDENTS Set MathPlaceID = Null WHERE SSN = '" & SSN & "'"
            sqlErr = ExecuteSQL(conn, sql)
            if sqlErr <> "" then
              estr = "Math Placement Test Detail NULL Set Error: " & sqlErr
            end if
          else
            Placement = CalculateMathPlacement(CDbl(ArithScore), CDbl(AlgScore), CDbl(CollegeScore))
            sql = "SET NOCOUNT ON " 
            sql = sql & "INSERT INTO MATHPLACEMENT (SSN, Date, ArithScore, AlgScore, CollegeScore, Placement, Status, IPAddr) "
            sql = sql & "VALUES ('" & SSN & "', '" & TestDate & "', " & ArithScore & ", " & AlgScore & ", " & CollegeScore & ", '" & Placement & "', 1, '" & IPAddr & "') "
            sql = sql & "SELECT SCOPE_IDENTITY() AS NewID"
            sqlErr = ExecuteSQLForRs(conn, sql, rs)
            if sqlErr <> "" then
              estr = "Math Placement Test Detail Insert Error: " & sqlErr
            else
              takeid = rs("NewID")
              rs.Close
              sql = "SET NOCOUNT ON " 
              sql = sql & "UPDATE STUDENTS SET MathPlaceID = " & takeid & " WHERE SSN = '" & SSN & "'"
              sqlErr = ExecuteSQL(conn, sql)
              if sqlErr <> "" then
                estr = "Math Placement Test Detail Insert Error: " & sqlErr
              end if
            end if
            Set rs = Nothing
          end if
        else
          estr = "Math Placement Test Detail Insert Error: Bad Value(s) Specified"
        end if
      end if
    elseif Operation = "DELETE" then
      if ID <> "NEW" then
        sql = "SET NOCOUNT ON " 
        sql = sql & "SELECT MathPlaceID FROM STUDENTS WHERE SSN = '" & SSN & "'"
        sqlErr = ExecuteSQLForRs(conn, sql, rs)
        if sqlErr <> "" then
          estr = "Math Placement Test Delete Error: " & sqlErr
        else
          takeid = rs("MathPlaceID")
          rs.Close
          if CLng(ID) = takeid then
            sql = "SET NOCOUNT ON " 
            sql = sql & "UPDATE STUDENTS SET MathPlaceID = NULL WHERE SSN = '" & SSN & "'"
            sqlErr = ExecuteSQL(conn, sql)
            if sqlErr <> "" then
              estr = "Math Placement Test Delete Error: " & sqlErr
            end if
          end if
          if sqlErr = "" then
            sql = "SET NOCOUNT ON " 
            sql = sql & "DELETE FROM MATHPLACEMENT WHERE ID = " & ID
            sqlErr = ExecuteSQL(conn, sql)
            if sqlErr <> "" then
              estr = "Math Placement Test Delete Error: " & sqlErr
            end if
          end if
        end if
        Set rs = Nothing
      end if
    end if
    MathPlacement = estr
  End Function
  
  Function EnglishPlacement(ByVal Operation, ByVal ID)
    Dim rs, takeid
    Dim sql, estr, sqlErr
    Dim TestDate, Score, WPScore, Placement, IPAddr
    Dim ReadingScore
    estr = ""
    if Operation = "SELECT" then
      if ID <> "NEW" then
        sql = "SET NOCOUNT ON " 
        sql = sql & "UPDATE STUDENTS SET EnglPlaceID = " & ID & " WHERE SSN = '" & SSN & "' "
        sql = sql & "UPDATE ENGLISHPLACEMENT SET Status = 1 WHERE ID = " & ID & " "
        sqlErr = ExecuteSQL(conn, sql)
        if sqlErr <> "" then
          estr = "English Placement Test Detail Set Error: " & sqlErr
        end if
      else
        TestDate = Trim(Request("Date"))
        Score = 0
        ReadingScore = 0
        WPScore = Trim(Request("WPScore"))
        IPAddr = Trim(Request("IPAddr"))
        Placement = CalculateEnglishPlacement(Score, ReadingScore, CInt(WPScore))
        if Placement <> "NULL" then
          sql = "SET NOCOUNT ON " 
          sql = sql & "INSERT INTO ENGLISHPLACEMENT (SSN, Date, Score, WPScore, Placement, Status, IPAddr) "
          sql = sql & "VALUES ('" & SSN & "', '" & TestDate & "', " & Score & ", " & WPScore & ", '" & Placement & "', 1, '" & IPAddr & "') "
          sql = sql & "SELECT SCOPE_IDENTITY() AS NewID"
          sqlErr = ExecuteSQLForRs(conn, sql, rs)
          if sqlErr <> "" then
            estr = "English Placement Test Detail Insert Error: " & sqlErr
          else
            takeid = rs("NewID")
            rs.Close
            sql = "SET NOCOUNT ON " 
            sql = sql & "UPDATE STUDENTS SET EnglPlaceID = " & takeid & " WHERE SSN = '" & SSN & "'"
            sqlErr = ExecuteSQL(conn, sql)
            if sqlErr <> "" then
              estr = "English Placement Test Detail Insert Error: " & sqlErr
            end if
          end if
        else
          sql = "SET NOCOUNT ON " 
          sql = sql & "UPDATE STUDENTS SET EnglPlaceID = NULL WHERE SSN = '" & SSN & "'"
          sqlErr = ExecuteSQL(conn, sql)
          if sqlErr <> "" then
            estr = "English Placement Test Null Out Error: " & sqlErr
          end if
        end if
        Set rs = Nothing
      end if
    elseif Operation = "DELETE" then
      if ID <> "NEW" then
        sql = "SET NOCOUNT ON " 
        sql = sql & "SELECT EnglPlaceID FROM STUDENTS WHERE SSN = '" & SSN & "'"
        sqlErr = ExecuteSQLForRs(conn, sql, rs)
        if sqlErr <> "" then
          estr = "English Placement Test Delete Error: " & sqlErr
        else
          takeid = rs("EnglPlaceID")
          rs.Close
          if CLng(ID) = takeid then
            sql = "SET NOCOUNT ON " 
            sql = sql & "UPDATE STUDENTS SET EnglPlaceID = NULL WHERE SSN = '" & SSN & "'"
            sqlErr = ExecuteSQL(conn, sql)
            if sqlErr <> "" then
              estr = "English Placement Test Delete Error: " & sqlErr
            end if
          end if
          if sqlErr = "" then
            sql = "SET NOCOUNT ON " 
            sql = sql & "DELETE FROM ENGLISHPLACEMENT WHERE ID = " & ID
            sqlErr = ExecuteSQL(conn, sql)
            if sqlErr <> "" then
              estr = "English Placement Test Delete Error: " & sqlErr
            end if
          end if
        end if
        Set rs = Nothing
      end if
    end if
    EnglishPlacement = estr
  End Function 
  
  Function EssaySamples(ByVal Operation, ByVal ID)
    Dim rs, takeid
    Dim sql, estr, sqlErr
    Dim RedoTIme
    estr = ""
    if Operation = "SELECT" then
      if ID <> "NEW" then
        RedoTime = Trim(Request("RedoTime" & ID))
        if RedoTime = "" then
          sql = "UPDATE EnglishEssays SET RedoTime = Null WHERE ID = " & ID
          sqlErr = ExecuteSQL(conn, sql)
          estr = BuildErrMsg(estr, sqlErr)
        elseif IsNumeric(RedoTime) then
          sql = "UPDATE EnglishEssays SET RedoTime = " & CInt(RedoTime) & " WHERE ID = " & ID
          sqlErr = ExecuteSQL(conn, sql)
          estr = BuildErrMsg(estr, sqlErr)
        else
          estr = BuildErrMsg(estr, "Specified Time is not Numeric")
        end if
        if estr = "" then       
          sql = "UPDATE [Full Placement Testing Join] SET EssayID = " & ID & " WHERE SSN = '" & SSN & "'"
          sqlErr = ExecuteSQL(conn, sql)
          estr = BuildErrMsg(estr, sqlErr)
        end if
      else
        sql = "UPDATE [Full Placement Testing Join] Set EssayID = Null WHERE SSN = '" & SSN & "'"
        sqlErr = ExecuteSQL(conn, sql)
        estr = BuildErrMsg(estr, sqlErr)
      end if
    elseif Operation = "DELETE" then
      if ID <> "NEW" then
        sql = "SELECT EssayID FROM [Full Placement Testing Join] WHERE SSN = '" & SSN & "'"
        sqlErr = ExecuteSQLForRs(conn, sql, rs)
        estr = BuildErrMsg(estr, sqlErr)
        takeid = rs("EssayID")
        rs.Close
        Set rs = Nothing
        if CLng(ID) = takeid then
          sql = "UPDATE [Full Placement Testing Join] SET EssayID = NULL WHERE SSN = '" & SSN & "'"
          sqlErr = ExecuteSQL(conn, sql)
          estr = BuildErrMsg(estr, sqlErr)
        end if
        sql = "SET NOCOUNT ON "
        sql = sql & "DELETE FROM ContactReaders WHERE EssayID = " & ID & " "
        sql = sql & "DELETE FROM EnglishEssays WHERE ID = " & ID & " " 
        sql = sql & "SET NOCOUNT OFF"
        sqlErr = ExecuteSQL(conn, sql)
        estr = BuildErrMsg(estr, sqlErr)
      end if
    end if
    if estr <> "" then
      estr = "Writing Sample Change Error: " & estr
    end if
    EssaySamples = estr
  End Function
  
  Function ContactReaders(ByVal Operation, ByVal ID)
    Dim sql, estr, sqlerr
    Dim placers, readersrs
    Dim ReaderID, EssayID
    estr = ""
    if Operation = "INSERT" then
      if ID = "NEW" then
        ReaderID = Trim(Request("ReaderID"))
        if ReaderID = "" then
          estr = "Specified Reader Is Not Valid"
        else
          EssayID = Request("EssayID")
          sql = "INSERT INTO ContactReaders (EssayID, SSN, ReaderID) VALUES (" & EssayID & ", '" & SSN & "', " & ReaderID & ")"
          sqlerr = ExecuteSQL(conn, sql)
          estr = BuildErrMsg(eStr, sqlerr)
          sql = "SELECT * FROM [Full Placement Testing Join] WHERE SSN = '" & SSN & "'"
          sqlerr = ExecuteSQLForRs(conn, sql, placers)
          estr = BuildErrMsg(eStr, sqlerr)
          sql = "Select ID, FullName, Email FROM EssayReaders WHERE ID = " & ReaderID
          sqlerr = ExecuteSQLForRs(conn, sql, readersrs)
          estr = BuildErrMsg(eStr, sqlerr)
          if estr = "" then
            Call ShootEmail(placers, readersrs, EssayID)
          end if
        end if
      else
        estr = "Select the New Reader Option"
      end if
    elseif Operation = "DELETE" then
      if ID <> "NEW" then
        EssayID = Request("EssayID")
        sql = "DELETE FROM ContactReaders WHERE ReaderID = " & ID & " AND EssayID = " & EssayID
        sqlerr = ExecuteSQL(conn, sql)
        estr = BuildErrMsg(eStr, sqlerr)
      else
        estr = "Select an Existing Reader"
      end if
    end if
    if estr <> "" then
      estr = "Write Sample Readers Change Error: " & estr
    end if
    ContactReaders = estr
  End Function

  Function ReadingPlacement(ByVal Operation, ByVal ID)
    Dim rs, takeid
    Dim sql, estr, sqlErr
    Dim TestDate, Score, Placement, IPAddr
    estr = ""
    if Operation = "SELECT" then
      if ID <> "NEW" then
        sql = "SET NOCOUNT ON " 
        sql = sql & "UPDATE STUDENTS SET ReadPlaceID = " & ID & " WHERE SSN = '" & SSN & "' "
        sql = sql & "UPDATE READINGPLACEMENT SET Status = 1 WHERE ID = " & ID & " "
        sqlErr = ExecuteSQL(conn, sql)
        if sqlErr <> "" then
          estr = "Reading Placement Test Detail Set Error: " & sqlErr
        end if
      else
        TestDate = Trim(Request("Date"))
        Score = Trim(Request("Score"))
        IPAddr = Trim(Request("IPAddr"))
        if IsNumeric(Score) then
          Placement = CalculateReadingPlacement(CDbl(Score))
          sql = "SET NOCOUNT ON " 
          sql = sql & "INSERT INTO READINGPLACEMENT (SSN, Date, Score, Placement, Status, IPAddr) "
          sql = sql & "VALUES ('" & SSN & "', '" & TestDate & "', " & Score & ", '" & Placement & "', 1, '" & IPAddr & "') "
          sql = sql & "SELECT SCOPE_IDENTITY() AS NewID"
          sqlErr = ExecuteSQLForRs(conn, sql, rs)
          if sqlErr <> "" then
            estr = "Reading Placement Test Detail Insert Error: " & sqlErr
          else
            takeid = rs("NewID")
            rs.Close
            sql = "SET NOCOUNT ON " 
            sql = sql & "UPDATE STUDENTS SET ReadPlaceID = " & takeid & " WHERE SSN = '" & SSN & "'"
            sqlErr = ExecuteSQL(conn, sql)
            if sqlErr <> "" then
              estr = "Reading Placement Test Detail Insert Error: " & sqlErr
            end if
          end if
          Set rs = Nothing
        elseif Score <> "" then
          estr = "Reading Placement Test Detail Insert Error: Bad Value(s) Specified"
        else
          sql = "SET NOCOUNT ON "
          sql = sql & "UPDATE STUDENTS Set ReadPlaceID = Null WHERE SSN = '" & SSN & "'"
          sqlErr = ExecuteSQL(conn, sql)
          if sqlErr <> "" then
            estr = "Reading Placement Test Detail NULL Set Error: " & sqlErr
          end if
        end if
      end if
    elseif Operation = "DELETE" then
      if ID <> "NEW" then
        sql = "SET NOCOUNT ON "
        sql = sql & "SELECT ReadPlaceID FROM STUDENTS WHERE SSN = '" & SSN & "'"
        sqlErr = ExecuteSQLForRs(conn, sql, rs)
        if sqlErr <> "" then
          estr = "Reading Placement Test Delete Error: " & sqlErr
        else
          takeid = rs("ReadPlaceID")
          rs.Close
          if CLng(ID) = takeid then
            sql = "SET NOCOUNT ON "
            sql = sql & "UPDATE STUDENTS SET ReadPlaceID = NULL WHERE SSN = '" & SSN & "'"
            sqlErr = ExecuteSQL(conn, sql)
            if sqlErr <> "" then
              estr = "Reading Placement Test Delete Error: " & sqlErr
            end if
          end if
          if sqlErr = "" then
            sql = "SET NOCOUNT ON "
            sql = sql & "DELETE FROM READINGPLACEMENT WHERE ID = " & ID
            sqlErr = ExecuteSQL(conn, sql)
            if sqlErr <> "" then
              estr = "Reading Placement Test Delete Error: " & sqlErr
            end if
          end if
        end if
        Set rs = Nothing
      end if
    end if
    ReadingPlacement = estr
  End Function 

  Function TypingPlacement(ByVal Operation, ByVal ID)
    Dim rs, takeid
    Dim sql, estr, sqlErr
    Dim TestDate, PassageNo, Words, Errors, Placement, IPAddr  
    estr = ""
    if Operation = "SELECT" then
      if ID <> "NEW" then
        sql = "SET NOCOUNT ON " 
        sql = sql & "UPDATE STUDENTS SET TypePlaceID = " & ID & " WHERE SSN = '" & SSN & "' "
        sql = sql & "UPDATE TYPINGPLACEMENT SET Status = 1 WHERE ID = " & ID & " "
        sqlErr = ExecuteSQL(conn, sql)
        if sqlErr <> "" then
          estr = "Keyboarding Placement Test Detail Set Error: " & sqlErr
        end if
      else
        TestDate = Trim(Request("Date"))
        PassageNo = Trim(Request("PassageNo"))
        Words = Trim(Request("WordsPerMin"))
        Errors = Trim(Request("Errors"))
        IPAddr = Trim(Request("IPAddr"))
        if IsNumeric(PassageNo) and IsNumeric(Words) and IsNumeric(Errors) then
          Placement = CalculateTypingPlacement(CInt(Words), CInt(Errors))
          sql = "SET NOCOUNT ON " 
          sql = sql & "INSERT INTO TYPINGPLACEMENT (SSN, Date, PassageNo, WordsPerMin, Errors, Placement, Status, IPAddr) "
          sql = sql & "VALUES ('" & SSN & "', '" & TestDate & "', " & PassageNo & ", " & Words & ", " & Errors & ", '" & Placement & "', 1, '" & IPAddr & "') "
          sql = sql & "SELECT SCOPE_IDENTITY() AS NewID"
          sqlErr = ExecuteSQLForRs(conn, sql, rs)
          if sqlErr <> "" then
            estr = "Keyboarding Placement Test Detail Insert Error: " & sqlErr
          else
            takeid = rs("NewID")
            rs.Close
            sql = "SET NOCOUNT ON "
            sql = sql & "UPDATE STUDENTS SET TypePlaceID = " & takeid & " WHERE SSN = '" & SSN & "'"
            sqlErr = ExecuteSQL(conn, sql)
            if sqlErr <> "" then
              estr = "Keyboarding Placement Test Detail Insert Error: " & sqlErr
            end if
          end if
          Set rs = Nothing
        elseif PassageNo & Words & Errors <> "" then
          estr = "Keyboarding Placement Test Detail Insert Error: Bad Value(s) Specified"
        else
          sql = "SET NOCOUNT ON "
          sql = sql & "UPDATE STUDENTS Set TypePlaceID = Null WHERE SSN = '" & SSN & "'"
          sqlErr = ExecuteSQL(conn, sql)
          if sqlErr <> "" then
            estr = "Keyboarding Placement Test Detail NULL Set Error: " & sqlErr
          end if
        end if
      end if
    elseif Operation = "DELETE" then
      if ID <> "NEW" then
        sql = "SET NOCOUNT ON "
        sql = sql & "SELECT TypePlaceID FROM STUDENTS WHERE SSN = '" & SSN & "'"
        sqlErr = ExecuteSQLForRs(conn, sql, rs)
        if sqlErr <> "" then
          estr = "Keyboarding Placement Test Delete Error: " & sqlErr
        else
          takeid = rs("TypePlaceID")
          rs.Close
          if CLng(ID) = takeid then
            sql = "SET NOCOUNT ON "
            sql = sql & "UPDATE STUDENTS SET TypePlaceID = NULL WHERE SSN = '" & SSN & "'"
            sqlErr = ExecuteSQL(conn, sql)
            if sqlErr <> "" then
              estr = "Keyboarding Placement Test Delete Error: " & sqlErr
            end if
          end if
          if sqlErr = "" then
            sql = "SET NOCOUNT ON "
            sql = sql & "DELETE FROM TYPINGPLACEMENT WHERE ID = " & ID
            sqlErr = ExecuteSQL(conn, sql)
            if sqlErr <> "" then
              estr = "Keyboarding Placement Test Delete Error: " & sqlErr
            end if
          end if
        end if
        Set rs = Nothing
      end if
    end if
    TypingPlacement = estr
  End Function 

  Function ExitPlacement(ByVal Operation, ByVal ID)
    Dim rs, takeid
    Dim sql, estr, sqlErr
    Dim TestDate, Score, Placement, IPAddr  
    estr = ""
    if Operation = "SELECT" then
      if ID <> "NEW" then
        sql = "SET NOCOUNT ON " 
        sql = sql & "UPDATE STUDENTS SET ExitPlaceID = " & ID & " WHERE SSN = '" & SSN & "' "
        sql = sql & "UPDATE READINGEXITPLACEMENT SET Status = 1 WHERE ID = " & ID & " "
        sqlErr = ExecuteSQL(conn, sql)
        if sqlErr <> "" then
          estr = "Reading Exit Placement Test Detail Set Error: " & sqlErr
        end if
      else
        TestDate = Trim(Request("Date"))
        Score = Trim(Request("Score"))
        IPAddr = Trim(Request("IPAddr"))
        if IsNumeric(Score) then
          Placement = CalculateExitPlacement(CDbl(Score))
          sql = "SET NOCOUNT ON " 
          sql = sql  & "INSERT INTO READINGEXITPLACEMENT (SSN, Date, Score, Placement, Status, IPAddr) "
          sql = sql  & "  VALUES ('" & SSN & "', '" & TestDate & "', " & Score & ", '" & Placement & "', 1, '" & IPAddr & "') "
          sql = sql  & "SELECT SCOPE_IDENTITY() AS NewID "
          sqlErr = ExecuteSQLForRs(conn, sql, rs)
          if sqlErr <> "" then
            estr = "Reading Exit Placement Test Detail Insert Error: " & sqlErr
          else
            takeid = rs("NewID")
            rs.Close
            sql = "SET NOCOUNT ON " 
            sql = sql & "UPDATE STUDENTS SET ExitPlaceID = " & takeid & " WHERE SSN = '" & SSN & "'"
            sqlErr = ExecuteSQL(conn, sql)
            if sqlErr <> "" then
              estr = "Reading Exit Placement Test Detail Insert Error: " & sqlErr
            end if
          end if
          Set rs = Nothing
        elseif Score <> "" then
          estr = "Reading Exit Placement Test Detail Insert Error: Bad Value(s) Specified"
        else
          sql = "SET NOCOUNT ON " 
          sql = sql & "UPDATE STUDENTS Set ExitPlaceID = Null WHERE SSN = '" & SSN & "'"
          sqlErr = ExecuteSQL(conn, sql)
          if sqlErr <> "" then
            estr = "Reading Exit Placement Test Detail NULL Set Error: " & sqlErr
          end if
        end if
      end if
    elseif Operation = "DELETE" then
      if ID <> "NEW" then
        sql = "SET NOCOUNT ON " 
        sql = sql & "SELECT ExitPlaceID FROM STUDENTS WHERE SSN = '" & SSN & "'"
        sqlErr = ExecuteSQLForRs(conn, sql, rs)
        if sqlErr <> "" then
          estr = "Reading Exit Placement Test Delete Error: " & sqlErr
        else
          takeid = rs("ExitPlaceID")
          rs.Close
          if CLng(ID) = takeid then
            sql = "SET NOCOUNT ON " 
            sql = sql & "UPDATE STUDENTS SET ExitPlaceID = NULL WHERE SSN = '" & SSN & "'"
            sqlErr = ExecuteSQL(conn, sql)
            if sqlErr <> "" then
              estr = "Reading Exit Placement Test Delete Error: " & sqlErr
            end if
          end if
          if sqlErr = "" then
            sql = "SET NOCOUNT ON " 
            sql = sql & "DELETE FROM READINGEXITPLACEMENT WHERE ID = " & ID
            sqlErr = ExecuteSQL(conn, sql)
            if sqlErr <> "" then
              estr = "Reading Exit Placement Test Delete Error: " & sqlErr
            end if
          end if
        end if
        Set rs = Nothing
      end if
    end if
    ExitPlacement = estr
  End Function 
  
  Sub ShowMessage(ByVal TestType, ByVal TestValue, ByVal eStr)
    if TestType = TestValue then
      if eStr <> "" then
        Response.Write "<h3 class=""red center"">Error(s) Occurred in Record Test Details Edit<br />Check Your Input Values<br/>" & _
                       eStr & "</h3>" & vbNewLine
      elseif Request.ServerVariables("REQUEST_METHOD") = "POST" then
        Response.Write "<h3 class=""green"">Successful Submission!!!</h3>" & vbNewLine
      end if
    End if
  End Sub

%>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<title><% =SYSTEM_NAME %> - Edit Test Details</title>
<link rel="stylesheet" href="css/placement.css" />
<link rel="stylesheet" href="css/gototop.css" />
<style>
table {
  text-align: left;
  margin-right: auto;
  margin-left: auto;
}
th {
	text-align: center;
}
.detailRow {
	height: 50px;
}
hr {
	margin-top: 15px;
}
table#menu {
  border: none;
  margin-top: 10px;	
}
table#menu td {
  vertical-align: top;
}
table#samples {
  border: none;
  margin-top: 10px;	
}
table#samples td {
  vertical-align: top;
}
table.scores {
	border: thin black outset;
	border-collapse: separate;
	border-spacing: 2px;
}
table.scores td, table.scores th {
	border: thin black inset;
	padding: 5px;
}

</style>
</head>

<body>
  <div class="center">
    <%
      Call MakeHeader("Edit Test Details")
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
  strQuery = "SELECT FirstName, LastName FROM STUDENTS WHERE StudentID = " & StudentID
  Set rs = conn.Execute(strQuery)
%>
<form method="post">
<p class="bold">Last Name: <input type="text" size="20" readonly="readonly" value="<%=rs("LastName")%>" />&nbsp; 
First Name: <input type="text" size="20" readonly="readonly" value="<%=rs("FirstName")%>" /></p>
</form>
<%  
  rs.Close
  Set rs = Nothing
%>
<hr />
<table id="menu">
	<tr>
		<td style="vertical-align:top">
<ul>
  <li><a href="#math">Math Placement Tests</a></li>
  <li><a href="#english">English Placement Tests</a></li>
  <li><a href="#essay">Writing Samples</a></li>
  <li><a href="#essayreaders">Writing Sample Readers</a></li>
</ul>
    	</td>
		<td style="vertical-align:top">
<ul>
  <li><a href="#reading">Reading Placement Tests</a></li>
  <li><a href="#keyboard">Keyboarding Placement Tests</a></li>
  <li><a href="#exit">Reading Exit Placement Tests</a></li>
</ul>
    	</td>
	</tr>
</table>

<hr />
<h3 id="math">Math Placement Tests</h3>
<%
  Call ShowMessage(TestType, "Math Placement", eStr)
%>
<form action="#math" method="post">
<table class="scores">
<%
  strQuery = "SELECT S.MathPlaceID, M.ID, M.Date, M.ArithScore, M.AlgScore, M.CollegeScore, M.Placement, M.IPAddr FROM " & _
             "STUDENTS S RIGHT OUTER JOIN MathPlacement M ON S.MathPlaceID = M.ID WHERE M.SSN = '" & SSN & "' ORDER BY DATE, ID" 
  Set rs = conn.Execute(strQuery)
  Response.Write "<tr>" & vbNewLine
  Response.Write "<th rowspan=""2"">&nbsp;</th>" & vbNewLine
  Response.Write "<th rowspan=""2"">Date</th>" & vbNewLine
  Response.Write "<th colspan=""3"">Scores</th>" & vbNewLine
  Response.Write "<th rowspan=""2"">Placement</th>" & vbNewLine
  Response.Write "<th rowspan=""2"">IP Address</th>" & vbNewLine
  Response.Write "</tr>" & vbNewLine
  Response.Write "<tr>" & vbNewLine
  Response.Write "<th>Arithmetic</th>" & vbNewLine
  Response.Write "<th>Quant-Alg-Stats</th>" & vbNewLine
  Response.Write "<th>Adv Alg-Func</th>" & vbNewLine
  Response.Write "</tr>" & vbNewLine
  do while not rs.Eof
    Response.Write "<tr>" & vbNewLine
    Response.Write "<td><input name=""ID"" type=""radio"" value=""" & rs("ID").value & """" & iif(rs("MathPlaceID").value = rs("ID").value, " checked=""checked""", "") & " /></td>" & vbNewLine
    Response.Write "<td>" & FixNull(rs("Date").value) & "</td>" & vbNewLine
    Response.Write "<td>" & FixNull(rs("ArithScore").value) & "</td>" & vbNewLine
    Response.Write "<td>" & FixNull(rs("AlgScore").value) & "</td>" & vbNewLine
    Response.Write "<td>" & FixNull(rs("CollegeScore").value) & "</td>" & vbNewLine
    Response.Write "<td>" & FixNull(rs("Placement").value) & "</td>" & vbNewLine
    Response.Write "<td>" & FixNull(rs("IPAddr").value) & "</td>" & vbNewLine
    Response.Write "</tr>" & vbNewLine
    rs.MoveNext
  Loop
  Response.Write "<tr>" & vbNewLine
  Response.Write "<td><input name=""ID"" type=""radio"" value=""NEW"" /></td>" & vbNewLine
  Response.Write "<td><input type=""text"" size=""20"" maxlength=""20"" name=""Date"" value=""" & Now() & """ /></td>" & vbNewLine
  Response.Write "<td><input type=""text"" size=""5"" maxlength=""10"" name=""ArithScore"" /></td>" & vbNewLine
  Response.Write "<td><input type=""text"" size=""5"" maxlength=""10"" name=""AlgScore"" /></td>" & vbNewLine
  Response.Write "<td><input type=""text"" size=""5"" maxlength=""10"" name=""CollegeScore"" /></td>" & vbNewLine
  Response.Write "<td>&nbsp;</td>" & vbNewLine
  Response.Write "<td><input type=""hidden"" value=""*" & Request.ServerVariables("REMOTE_ADDR") & """ name=""IPAddr"" />&nbsp;</td>" & vbNewLine
  Response.Write "</tr>" & vbNewLine

  Response.Write "<tr>" & vbNewLine
  Response.Write "<td class=""center"" colspan=""7"">" 
  Response.Write "<input type=""submit"" name=""Operation"" value=""Select"" style=""margin-right: 25px"" />"
  Response.Write "<input type=""submit"" name=""Operation"" value=""Delete"" style=""margin-right: 25px"" />"
  Response.Write "<input type=""reset"" name=""Reset"" />"
  Response.Write "<input type=""hidden"" name=""TestType"" value=""Math Placement"" />" 
  Response.Write "</td>" & vbNewLine
  Response.Write "</tr>" & vbNewLine  
  rs.Close
  Set rs = Nothing
%>
</table>
</form>

<hr />
<h3 id="english">English Placement Tests</h3>
<%
  Call ShowMessage(TestType, "English Placement", eStr)
%>
<form action="#english" method="post">
<table class="scores">
<%
  strQuery = "SELECT S.EnglPlaceID, E.ID, E.Date, E.Score, E.WPScore, E.Placement, E.IPAddr FROM " & _
             "STUDENTS S RIGHT OUTER JOIN EnglishPlacement E ON S.EnglPlaceID = E.ID WHERE E.SSN = '" & SSN & "' ORDER BY DATE, ID"
  Set rs = conn.Execute(strQuery)
  Response.Write "<tr>" & vbNewLine
  Response.Write "<th>&nbsp;</th>" & vbNewLine
  Response.Write "<th>Date</th>" & vbNewLine
  Response.Write "<th>SS Score</th>" & vbNewLine
  Response.Write "<th>WritePlacer Score</th>" & vbNewLine
  Response.Write "<th>Placement</th>" & vbNewLine
  Response.Write "<th>IP Address</th>" & vbNewLine
  Response.Write "</tr>" & vbNewLine
  do while not rs.Eof
    Response.Write "<tr>" & vbNewLine
    Response.Write "<td><input name=""ID"" type=""radio"" value=""" & rs("ID").value & """" & iif(rs("EnglPlaceID").value = rs("ID").value, " checked=""checked""", "") & " /></td>" & vbNewLine
    Response.Write "<td>" & FixNull(rs("Date").value) & "</td>" & vbNewLine
    Response.Write "<td>" & rs("Score") & "</td>" & vbNewLine
    WPScore = rs("WPScore").value
    if IsNull(WPScore) then
      WPScore = "Essay Not Taken"
    elseif WPScore = -2 then
      WPScore = "Essay Not Taken"
    else
      WPScore = CStr(WPScore)
    end if
    Response.Write "<td>" & WPScore & "</td>" & vbNewLine
    Response.Write "<td>" & FixNull(rs("Placement").value) & "</td>" & vbNewLine
    Response.Write "<td>" & FixNull(rs("IPAddr").value) & "</td>" & vbNewLine
    Response.Write "</tr>" & vbNewLine
    rs.MoveNext
  Loop
  Response.Write "<tr>" & vbNewLine
  Response.Write "<td><input name=""ID"" type=""radio"" value=""NEW"" /></td>" & vbNewLine
  Response.Write "<td><input type=""text"" size=""20"" maxlength=""20"" name=""Date"" value=""" & Now() & """ /></td>" & vbNewLine
  Response.Write "<td>&nbsp;</td>" & vbNewLine
  Response.Write "<td><select name=""WPScore"">" &vbNewLine
  Response.Write "<option value=""-2"">Essay Not Taken</option>" &vbNewLine
  Response.Write "<option value=""0"">0</option>" &vbNewLine
  Response.Write "<option value=""1"">1</option>" &vbNewLine
  Response.Write "<option value=""2"">2</option>" &vbNewLine
  Response.Write "<option value=""3"">3</option>" &vbNewLine
  Response.Write "<option value=""4"">4</option>" &vbNewLine
  Response.Write "<option value=""5"">5</option>" &vbNewLine
  Response.Write "<option value=""6"">6</option>" &vbNewLine
  Response.Write "<option value=""7"">7</option>" &vbNewLine
  Response.Write "<option value=""8"">8</option>" &vbNewLine
  Response.Write "</select></td>" & vbNewLine
  Response.Write "<td>&nbsp;</td>" & vbNewLine
  Response.Write "<td><input type=""hidden"" value=""*" & Request.ServerVariables("REMOTE_ADDR") & """ name=""IPAddr"" />&nbsp;</td>" & vbNewLine
  Response.Write "</tr>" & vbNewLine

  Response.Write "<tr>" & vbNewLine
  Response.Write "<td class=""center"" colspan=""6"">" 
  Response.Write "<input type=""submit"" name=""Operation"" value=""Select"" style=""margin-right: 25px"" />"
  Response.Write "<input type=""submit"" name=""Operation"" value=""Delete"" style=""margin-right: 25px"" />"
  Response.Write "<input type=""reset"" name=""Reset"" />"
  Response.Write "<input type=""hidden"" name=""TestType"" value=""English Placement"" />" 
  Response.Write "</td>" & vbNewLine
  Response.Write "</tr>" & vbNewLine  
  rs.Close
  Set rs = Nothing
%>
</table>
</form>

<hr />

<h3 id="essay" class="center">Writing Samples</h3>
<%
  Call ShowMessage(TestType, "Writing Samples", eStr)
%>
<form action="#essay" method="post">
<table class="scores">
<%
  strQuery = "SELECT EP.EssayID, E.ID, E.Date, E.RedoTime, E.IPAddr FROM " & _
             "[Full Placement Testing Join] EP RIGHT OUTER JOIN EnglishEssays E ON EP.EssayID = E.ID WHERE E.SSN = '" & SSN & "' ORDER BY DATE, ID"
  Set rs = conn.Execute(strQuery)
  Response.Write "<tr>" & vbNewLine
  Response.Write "<th>&nbsp;</th>" & vbNewLine
  Response.Write "<th>Date</th>" & vbNewLine
  Response.Write "<th>Redo Time</th>" & vbNewLine
  Response.Write "<th>Essay</th>" & vbNewLine
  Response.Write "<th>IP Address</th>" & vbNewLine
  Response.Write "</tr>" & vbNewLine
  do while not rs.Eof
    Response.Write "<tr>" & vbNewLine
    Response.Write "<td><input name=""ID"" type=""radio"" value=""" & rs("ID").value & """" & iif(rs("EssayID").value = rs("ID").value, " checked=""checked""", "") &" /></td>" & vbNewLine
    Response.Write "<td>" & FixNull(rs("Date").value) & "</td>" & vbNewLine
    Response.Write "<td><input name=""RedoTime" & rs("ID") & """ type=""text"" value=""" & rs("RedoTime") & """ size=""3"" maxlength=""3"" /></td>" & vbNewLine
    Response.Write "<td><a href=""viewEssay.asp?EssayID=" & rs("ID") & """>Click to View</a></td>" & vbNewLine
    Response.Write "<td>" & FixNull(rs("IPAddr")) & "</td>" & vbNewLine
    Response.Write "</tr>" & vbNewLine
    rs.MoveNext
  Loop
  Response.Write "<tr>" & vbNewLine
  Response.Write "<td><input name=""ID"" type=""radio"" value=""NEW"" /></td>" & vbNewLine
  Response.Write "<td colspan=""4"">Unassign Essay</td>" & vbNewLine
  Response.Write "</tr>" & vbNewLine

  Response.Write "<tr>" & vbNewLine
  Response.Write "<td class=""center"" colspan=""5"">" & vbNewline 
  Response.Write "<input type=""submit"" name=""Operation"" value=""Select"" style=""margin-right: 25px"" />"
  Response.Write "<input type=""submit"" name=""Operation"" value=""Delete"" style=""margin-right: 25px"" />"
  Response.Write "<input type=""reset"" name=""Reset"" />"
  Response.Write "<input type=""hidden"" name=""TestType"" value=""Writing Samples"" />" 
  Response.Write "</td>" & vbNewLine
  Response.Write "</tr>" & vbNewLine  
  rs.Close
  Set rs = Nothing
%>
</table>
</form>

<%
  strQuery = "SELECT EssayID From [Full Placement Testing Join] WHERE SSN = '" & SSN & "'"
  Call ExecuteSQLForRs(conn, strQuery, temprs)
  EssayID = temprs("EssayID")
  temprs.Close
  Set temprs = Nothing
  if not IsNull(EssayID) then
%>
<hr/>
<h3 id="essayreaders" class="center">Writing Sample Readers</h3>
<%
  Call ShowMessage(TestType, "Contact Readers", eStr)
%>
<form action="#essayreaders" method="post">
<table class="scores">
<%
  strQuery = "SELECT B.ReaderID, A.FullName, B.ReaderScoreDate, B.ReaderPlacement "
  strQuery = strQuery & "FROM EssayReaders A INNER JOIN (SELECT ReaderID, ReaderScoreDate, ReaderPlacement, FPJ.SSN FROM ContactReaders CR INNER JOIN [Full Placement Testing Join] FPJ ON CR.EssayID = FPJ.EssayID) B ON B.ReaderID = A.ID "
  strQuery = strQuery & "WHERE B.SSN = '" & SSN & "'"
  Set rs = conn.Execute(strQuery)
  Response.Write "<tr>" & vbNewLine
  Response.Write "<th>&nbsp;</th>" & vbNewLine
  Response.Write "<th>Name</th>" & vbNewLine
  Response.Write "<th>Date</th>" & vbNewLine
  Response.Write "<th>Placement</th>" & vbNewLine
  Response.Write "</tr>" & vbNewLine
  do while not rs.Eof
    Response.Write "<tr>" & vbNewLine
    Response.Write "<td><input name=""ID"" type=""radio"" value=""" & rs("ReaderID") & """ /></td>" & vbNewLine
    Response.Write "<td>" & rs("FullName").value & "</td>" & vbNewLine
    Response.Write "<td>" & rs("ReaderScoreDate").value & "</td>" & vbNewLine
    Response.Write "<td>" & rs("ReaderPlacement").value & "</td>" & vbNewLine
    Response.Write "</tr>" & vbNewLine
    rs.MoveNext
  Loop
  Response.Write "<tr>" & vbNewLine
  Response.Write "<td><input name=""ID"" type=""radio"" value=""NEW"" /></td>" & vbNewLine
  Response.Write "<td><select name=""ReaderID"">" & vbNewLine
  Response.Write "<option value="""">Choose A Reader</option>" & vbNewLine
  strQuery = "SELECT ID, FullName FROM EssayReaders Order By FullName"
  Set readers = conn.Execute(strQuery)
  do while not readers.EOF
    Response.Write "<option value=""" & readers("ID") & """>" & readers("FullName") & "</option>" & vbNewLine
    readers.MoveNext
  Loop
  readers.Close
  Set readers = Nothing
  Response.Write "</select></td>" & vbNewLine
  Response.Write "<td>&nbsp;</td>" & vbNewLine
  Response.Write "<td>&nbsp;</td>" & vbNewLine
  Response.Write "</tr>" & vbNewLine

  Response.Write "<tr>" & vbNewLine
  Response.Write "<td class=""center"" colspan=""4"">" & vbNewLine
  Response.Write "<input type=""submit"" name=""Operation"" value=""Delete"" style=""margin-right: 25px"" />"
  Response.Write "<input type=""submit"" name=""Operation"" value=""Insert"" style=""margin-right: 25px"" />"
  Response.Write "<input type=""reset"" name=""Reset"" />"
  Response.Write "<input type=""hidden"" name=""TestType"" value=""Contact Readers"" />" 
  Response.Write "<input type=""hidden"" name=""EssayID"" value=""" & EssayID & """ />"
  Response.Write "</td>" & vbNewLine
  Response.Write "</tr>" & vbNewLine  
  rs.Close
  Set rs = Nothing
%>
</table>
</form>
<%
  end if
%>

<hr />
<h3 id="reading">Reading Placement Tests</h3>
<%
  Call ShowMessage(TestType, "Reading Placement", eStr)
%>
<form action="#reading" method="post">
<table class="scores">
<%
  strQuery = "SELECT S.ReadPlaceID, R.ID, R.Date, R.Score, R.Placement, R.IPAddr FROM " & _
             "STUDENTS S RIGHT OUTER JOIN ReadingPlacement R ON S.ReadPlaceID = R.ID WHERE R.SSN = '" & SSN & "' ORDER BY DATE, ID"
  Set rs = conn.Execute(strQuery)
  Response.Write "<tr>" & vbNewLine
  Response.Write "<th>&nbsp;</th>" & vbNewLine
  Response.Write "<th>Date</th>" & vbNewLine
  Response.Write "<th>Score</th>" & vbNewLine
  Response.Write "<th>Placement</th>" & vbNewLine
  Response.Write "<th>IP Address</th>" & vbNewLine
  Response.Write "</tr>" & vbNewLine
  do while not rs.Eof
    Response.Write "<tr>" & vbNewLine
    Response.Write "<td><input name=""ID"" type=""radio"" value=""" & rs("ID").value & """" & iif(rs("ReadPlaceID").value = rs("ID").value, " checked=""checked""", "") & " /></td>" & vbNewLine
    Response.Write "<td>" & FixNull(rs("Date").value) & "</td>" & vbNewLine
    Response.Write "<td>" & FixNull(rs("Score").value) & "</td>" & vbNewLine
    Response.Write "<td>" & FixNull(rs("Placement").value) & "</td>" & vbNewLine
    Response.Write "<td>" & FixNull(rs("IPAddr").value) & "</td>" & vbNewLine
    Response.Write "</tr>" & vbNewLine
    rs.MoveNext
  Loop
  Response.Write "<tr>" & vbNewLine
  Response.Write "<td><input name=""ID"" type=""radio"" value=""NEW"" /></td>" & vbNewLine
  Response.Write "<td><input type=""text"" size=""20"" maxlength=""20"" name=""Date"" value=""" & Now() & """ /></td>" & vbNewLine
  Response.Write "<td><input type=""text"" size=""5"" maxlength=""10"" name=""Score"" /></td>" & vbNewLine
  Response.Write "<td>&nbsp;</td>" & vbNewLine
  Response.Write "<td><input type=""hidden"" value=""*" & Request.ServerVariables("REMOTE_ADDR") & """ name=""IPAddr"" />&nbsp;</td>" & vbNewLine
  Response.Write "</tr>" & vbNewLine

  Response.Write "<tr>" & vbNewLine
  Response.Write "<td class=""center"" colspan=""5"">" 
  Response.Write "<input type=""submit"" name=""Operation"" value=""Select"" style=""margin-right: 25px"" />"
  Response.Write "<input type=""submit"" name=""Operation"" value=""Delete"" style=""margin-right: 25px"" />"
  Response.Write "<input type=""reset"" name=""Reset"" />"
  Response.Write "<input type=""hidden"" name=""TestType"" value=""Reading Placement"" />" 
  Response.Write "</td>" & vbNewLine
  Response.Write "</tr>" & vbNewLine  
  rs.Close
  Set rs = Nothing
%>
</table>
</form>

<hr />
<h3 id="keyboard">Keyboarding Placement Tests</h3>
<%
  Call ShowMessage(TestType, "Keyboarding Placement", eStr)
%>
<form action="#keyboard" method="post">
<table class="scores">
<%
  strQuery = "SELECT S.TypePlaceID, T.ID, T.Date, T.PassageNo, T.WordsPerMin, T.Errors, T.Placement, T.IPAddr FROM " & _
             "STUDENTS S RIGHT OUTER JOIN TypingPlacement T ON S.TypePlaceID = T.ID WHERE T.SSN = '" & SSN & "' ORDER BY DATE, ID"
  Set rs = conn.Execute(strQuery)
  Response.Write "<tr>" & vbNewLine
  Response.Write "<th>&nbsp;</th>" & vbNewLine
  Response.Write "<th>Date</th>" & vbNewLine
  Response.Write "<th>Passage</th>" & vbNewLine
  Response.Write "<th>Words/Min</th>" & vbNewLine
  Response.Write "<th>Errors</th>" & vbNewLine
  Response.Write "<th>Placement</th>" & vbNewLine
  Response.Write "<th>IP Address</th>" & vbNewLine
  Response.Write "</tr>" & vbNewLine
  do while not rs.Eof
    Response.Write "<tr>" & vbNewLine
    Response.Write "<td><input name=""ID"" type=""radio"" value=""" & rs("ID").value & """" & iif(rs("TypePlaceID").value = rs("ID").value, " checked=""checked""", "") & " /></td>" & vbNewLine
    Response.Write "<td>" & FixNull(rs("Date").value) & "</td>" & vbNewLine
    Response.Write "<td>" & FixNull(rs("PassageNo").value) & "</td>" & vbNewLine
    Response.Write "<td>" & FixNull(rs("WordsPerMin").value) & "</td>" & vbNewLine
    Response.Write "<td>" & FixNull(rs("Errors").value) & "</td>" & vbNewLine
    Response.Write "<td>" & FixNull(rs("Placement").value) & "</td>" & vbNewLine
    Response.Write "<td>" & FixNull(rs("IPAddr").value) & "</td>" & vbNewLine
    Response.Write "</tr>" & vbNewLine
    rs.MoveNext
  Loop
  Response.Write "<tr>" & vbNewLine
  Response.Write "<td><input name=""ID"" type=""radio"" value=""NEW"" /></td>" & vbNewLine
  Response.Write "<td><input type=""text"" size=""20"" maxlength=""20"" name=""Date"" value=""" & Now() & """ /></td>" & vbNewLine
  Response.Write "<td><input type=""text"" size=""5"" maxlength=""10"" name=""PassageNo"" /></td>" & vbNewLine
  Response.Write "<td><input type=""text"" size=""5"" maxlength=""10"" name=""WordsPerMin"" /></td>" & vbNewLine
  Response.Write "<td><input type=""text"" size=""5"" maxlength=""10"" name=""Errors"" /></td>" & vbNewLine
  Response.Write "<td>&nbsp;</td>" & vbNewLine
  Response.Write "<td><input type=""hidden"" value=""*" & Request.ServerVariables("REMOTE_ADDR") & """ name=""IPAddr"" />&nbsp;</td>" & vbNewLine
  Response.Write "</tr>" & vbNewLine
  
  Response.Write "<tr>" & vbNewLine
  Response.Write "<td class=""center"" colspan=""7"">" 
  Response.Write "<input type=""submit"" name=""Operation"" value=""Select"" style=""margin-right: 25px"" />"
  Response.Write "<input type=""submit"" name=""Operation"" value=""Delete"" style=""margin-right: 25px"" />"
  Response.Write "<input type=""reset"" name=""Reset"" />"
  Response.Write "<input type=""hidden"" name=""TestType"" value=""Keyboarding Placement"" />" 
  Response.Write "</td>" & vbNewLine
  Response.Write "</tr>" & vbNewLine
  rs.Close
  Set rs = Nothing
%>
</table>
</form>

<hr />
<h3 id="exit">Reading Exit Placement Tests</h3>
<%
  Call ShowMessage(TestType, "Reading Exit Placement", eStr)
%>
<form action="#exit" method="post">
<table class="scores">
<%
  strQuery = "SELECT S.ExitPlaceID, E.ID, E.Date, E.Score, E.Placement, E.IPAddr FROM " & _
             "STUDENTS S RIGHT OUTER JOIN ReadingExitPlacement E ON S.ExitPlaceID = E.ID WHERE E.SSN = '" & SSN & "' ORDER BY DATE, ID"
  Set rs = conn.Execute(strQuery)
  Response.Write "<tr>" & vbNewLine
  Response.Write "<th>&nbsp;</th>" & vbNewLine
  Response.Write "<th>Date</th>" & vbNewLine
  Response.Write "<th>Score</th>" & vbNewLine
  Response.Write "<th>Placement</th>" & vbNewLine
  Response.Write "<th>IP Address</th>" & vbNewLine
  Response.Write "</tr>" & vbNewLine
  do while not rs.Eof
    Response.Write "<tr>" & vbNewLine
    Response.Write "<td><input name=""ID"" type=""radio"" value=""" & rs("ID").value & """" & iif(rs("ExitPlaceID").value = rs("ID").value, " checked=""checked""", "") & " /></td>" & vbNewLine
    Response.Write "<td>" & FixNull(rs("Date").value) & "</td>" & vbNewLine
    Response.Write "<td>" & FixNull(rs("Score").value) & "</td>" & vbNewLine
    Response.Write "<td>" & FixNull(rs("Placement").value) & "</td>" & vbNewLine
    Response.Write "<td>" & FixNull(rs("IPAddr").value) & "</td>" & vbNewLine
    Response.Write "</tr>" & vbNewLine
    rs.MoveNext
  Loop
  Response.Write "<tr>" & vbNewLine
  Response.Write "<td><input name=""ID"" type=""radio"" value=""NEW"" /></td>" & vbNewLine
  Response.Write "<td><input type=""text"" size=""20"" maxlength=""20"" name=""Date"" value=""" & Now() & """ /></td>" & vbNewLine
  Response.Write "<td><input type=""text"" size=""5"" maxlength=""10"" name=""Score"" /></td>" & vbNewLine
  Response.Write "<td>&nbsp;</td>" & vbNewLine
  Response.Write "<td><input type=""hidden"" value=""*" & Request.ServerVariables("REMOTE_ADDR") & """ name=""IPAddr"" />&nbsp;</td>" & vbNewLine
  Response.Write "</tr>" & vbNewLine

  Response.Write "<tr>" & vbNewLine
  Response.Write "<td class=""center"" colspan=""5"">" 
  Response.Write "<input type=""submit"" name=""Operation"" value=""Select"" style=""margin-right: 25px"" />"
  Response.Write "<input type=""submit"" name=""Operation"" value=""Delete"" style=""margin-right: 25px"" />"
  Response.Write "<input type=""reset"" name=""Reset"" />"
  Response.Write "<input type=""hidden"" name=""TestType"" value=""Reading Exit Placement"" />" 
  Response.Write "</td>" & vbNewLine
  Response.Write "</tr>" & vbNewLine  
  rs.Close
  Set rs = Nothing
%>
</table>
</form>
</div>
<div id="VerifyDeleteBox"></div>
<script src="//ajax.googleapis.com/ajax/libs/jquery/2.2.4/jquery.min.js"></script> 
<script src="js/jquery.gototop.js"></script> 
<script>
$(function(){
  "use strict";
  $("#toTop").gototop({ container: "body" });
  
  $("input[type=submit]").click(function(e) {
    var btnVal = $(this).val();
    return window.confirm("OK to perform this " + btnVal + " operation?");
  });
});
</script>
</body>
</html>
<%
  closeConnection(conn)
%>
