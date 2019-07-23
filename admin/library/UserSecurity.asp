<%
  function isUser(ByVal aDictionary)
    if isObject(aDictionary) then
      if typename(aDictionary) = "Dictionary" then
        isUser= aDictionary.exists("username") and aDictionary.exists("password")
      end if
    end if
  end function

  function recordToDictionary(ByVal R)
    dim d, F
    set d = Server.CreateObject("Scripting.Dictionary")
    for each F in R.Fields
      d.Add F.Name, F.Value
    next
    set recordToDictionary = d
  end function

  function getUserRecord(ByVal conn, ByVal aSignon, ByVal aPassword)
    Dim cmd
    set cmd = Server.CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    cmd.CommandText = "PlacementTesting_Get_Users('" & aSignon & "', '" & aPassword & "')"
    cmd.CommandType = adCmdStoredProc
    set getUserRecord = cmd.Execute()
    set cmd = Nothing
  end function

  function signUserOn(ByVal aSignon, ByVal aPassword)
    dim dict, conn, R
    aSignon = trim(aSignon)
    aPassword = trim(aPassword)
    set conn = openConnection(Application("UserDirectoryConnectionString"))
    set R = getUserRecord(conn, aSignon, aPassword)  
    if not R.EOF then
      set dict = recordToDictionary(R)
    else
      set dict = nothing
    end if
    R.Close
    set R = nothing
    closeConnection(conn)
    if isUser(dict) then
      signUserOn = true
      set Session("User") = dict
    else
      signUserOn = false
      set Session("User") = Nothing
    end if
  end function
%>