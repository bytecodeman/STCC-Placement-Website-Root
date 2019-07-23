<%  
  Function CalculateMathPlacement(ByVal dblArithScore, ByVal dblAlgScore, ByVal dblCollegeScore)
    Dim Placement : Placement = "ERROR"
    if dblAlgScore > 0 then
      If dblAlgScore < 230 Then
        if dblArithScore > 0 then
          If dblArithScore < 230 Then
            Placement = "MAT071" '5 Day Pre Algebra ???
          ElseIf dblArithScore < 250 Then
            Placement = "MAT071" 'Pre Algebra Pre Stats???
          ElseIf dblArithScore < 270 Then
            Placement = "MAT071U" '5 Day Extended Algebra 1
          Else
            Placement = "MAT081" 'Regular Algebra 1
          End If
        end if
      ElseIf dblAlgScore >= 270 Then
        if dblCollegeScore > 0 then
          If dblCollegeScore < 243 Then
            If dblAlgScore < 250 Then
              Placement = "MAT081"
            ElseIf dblAlgScore < 268 Then
              Placement = "MAT081U"
            ElseIf dblAlgScore < 290 Then
              Placement = "MAT091"
            Else
              Placement = "MAT101"
            End If
          ElseIf dblCollegeScore < 263 Then
            Placement = "MAT105"
          Else
            Placement = "MAT131"
          End If
        end if
      Else
        If dblAlgScore < 250 Then
          Placement = "MAT081"
        ElseIf dblAlgScore < 268 Then
          Placement = "MAT081U"
        ElseIf dblAlgScore < 290 Then
          Placement = "MAT091"
        Else
          Placement = "MAT101"
        End If
      End If
    End if
    CalculateMathPlacement = Placement
  End Function
    
  Function CalculateReadingPlacement(ByVal dblrcscore)
    Dim Placement : Placement = "ERROR"
    If dblrcscore < 245 then
      Placement = "DRG091"
    ElseIf dblrcscore < 259 Then
      Placement = "DRG092"
    Else
      Placement = "READ105"
    End If       
    CalculateReadingPlacement = Placement
  End Function 
   
  Function CalculateEnglishPlacement(ByVal EnglScore, ByVal ReadScore, ByVal WPScore)
    Dim Placement : Placement = "ERROR"
    Select Case WPScore
      Case 7, 8
        Placement = "ENG101H"
      Case 5, 6
		Placement = "ENG101"
      Case 4
		Placement = "DWT099U"
	  Case 3
		Placement = "DWT099C"
	  Case 0, 1, 2
		Placement = "DWT099"
	  Case -1
        Placement = "ESSAY"
      Case -2
        Placement = "NULL"
    End Select
    CalculateEnglishPlacement = Placement
  End Function
  
  Function CalculateTypingPlacement(ByVal mWords, ByVal mErrors)
    Dim Placement
    if mWords < 20 then
      Placement = "OIT100"
    elseif mWords < 45 then
      Placement = iif(mErrors <= 2, "OIT110", "OIT100")
    else
      Placement = iif(mErrors <= 3, "OIT110", "OIT100")
    end if
    CalculateTypingPlacement = Placement
  End Function  
  
  Function CalculateExitPlacement(ByVal dblrcscore)
    Dim Placement : Placement = "ERROR"
    If dblrcscore < 245 then
      Placement = "DRG091E"
    ElseIf dblrcscore < 259 Then
      Placement = "DRG092E"
    Else
      Placement = "READ105E"
    End if
    CalculateExitPlacement = Placement
  End Function

%>
