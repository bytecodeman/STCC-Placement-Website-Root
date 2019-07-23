<%
  Const DEFAULTRETURNVALUE = False
 
  Private Function AmIOnCampus(ByVal addr)
    Dim i, bytes, subnets
    subnets = Array(61, 62, 63, 64, 96, 97, 98, 99, 100, 101, 176, 177, 178, 183)
    AmIOnCampus = False
    bytes = Split(addr, ".")
    if CInt(bytes(0)) = 127 And CInt(bytes(1)) = 0 And CInt(bytes(2)) = 0 And CInt(bytes(3)) = 1 Then
       AmIOnCampus = True
    elseif CInt(bytes(0)) = 10 then
      AmIOnCampus = true
    elseif CInt(bytes(0)) = 192 and CInt(bytes(1)) = 168 then
      AmIOnCampus = true
    elseif CInt(bytes(0)) = 172 and CInt(bytes(1)) >= 16 And CInt(bytes(1)) <= 31 then
      AmIOnCampus = true
    elseif Cint(bytes(0)) = 134 and Cint(bytes(1)) = 241 then
      for i = 0 to UBound(subnets)
        if Cint(bytes(2)) = subnets(i) then
          exit for
        end if
      next
      AmIOnCampus = i <= UBound(subnets)
    end if
  End Function
  
  Public Function NorthAmericanVisitor(ByVal IPAddress)
    if AmIOnCampus(IPAddress) then
      NorthAmericanVisitor = true
      Exit Function
    end If
	NorthAmericanVisitor = DEFAULTRETURNVALUE
  end function
%>
