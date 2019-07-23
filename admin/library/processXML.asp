<%
  Function GetXMLString(Section, KeyName, Default, FileName)
    Dim objXML, objLst, i, objSection, objSetting
    
    GetXMLString = Default
    Set objXML = Server.CreateObject("Microsoft.XMLDOM")
    objXML.async = False
    objXML.Load (Server.MapPath(Filename))
    If objXML.parseError.errorCode <> 0 Then
	  Err.Raise 9000, "GetXMLString", "ERROR Parsing XML File: " & Filename
    End If
 
    'Find section
    Set objLst = objXML.getElementsByTagName("section")
    Set objSection = Nothing
    for i = 0 to objLst.length - 1
      If objLst.item(i).getAttribute("name") = Section Then
        Set objSection = objLst.item(i)
        Exit For
      End If
    Next
    if objSection Is Nothing then
      Exit Function
    end if
  
    'Find Setting
    Set objLst = objSection.getElementsByTagName("setting")
    Set objSetting = Nothing
    for i = 0 to objLst.length - 1
      If objLst.item(i).getAttribute("name") = KeyName Then
        Set objSetting = objLst.item(i)
        Exit For
      End If
    Next
    if objSetting Is Nothing then
      Exit Function
    end if
    
    if IsNull(objSetting.getAttribute("value")) then
      GetXMLString = objSetting.text
    else
      GetXMLString = objSetting.getAttribute("value")
    end if
    Set objXML = Nothing
  End Function
%>