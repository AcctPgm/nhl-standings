Attribute VB_Name = "mCalculations"
Option Explicit

Sub Calculations(x)
'
' Add calculations used in the worksheet
'
    Dim c As Range
    Dim nme As String
    Dim sColumn As String
    
    ' Status update
    Call ufmStatusMessage.ShowMessage("Adding Calculations")
    
    ' Remove place numbers and other indicators from the team names
    For Each c In wsStandings.Range("B2:B32")
        ' Replace any non-breaking space characters (160) with a regular space
        c.Value = Replace(c.Text, Chr(160), " ")
        nme = Mid(c.Text, InStr(c.Text, " ") + 1, Len(c.Text))
        If Left(nme, 2) = "x-" Or Left(nme, 2) = "y-" Or Left(nme, 2) = "z-" Then
            nme = Mid(nme, 3, Len(nme))
        ElseIf Left(nme, 4) = "xyz-" Then
            nme = Mid(nme, 5, Len(nme))
        End If
        c.Value = nme
    Next c
    
    ' Add an overall points standing counter
    wsStandings.Range("A1").Formula = "#"
    wsStandings.Columns("A:A").ColumnWidth = 4
    
    ' Delete the Streak column
'    wsStandings.Columns("M:M").Delete Shift:=xlToLeft
    
    ' New columns
    wsStandings.Columns("M:M").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    wsStandings.Columns("L:L").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    wsStandings.Columns("K:M").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    wsStandings.Columns("J:J").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    wsStandings.Columns("H:K").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    ' Add range names
    Call RangeNames(0)
    
    ' Add the conferences and divisions from the macro workbook
        ' Conference
    wsStandings.Cells(1, wbStandings.Names("Conf").RefersToRange.Column).Formula = "Conf"
    wsStandings.Range("Conf").Formula = _
        "=INDEX('" & ThisWorkbook.Name & "'!TeamConf,MATCH(B2,'" & ThisWorkbook.Name & "'!TeamName,0))"
        ' Static values
    wsStandings.Range("Conf").Copy
    wsStandings.Range("Conf").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
        ' Division
    wsStandings.Cells(1, wbStandings.Names("Div").RefersToRange.Column).Formula = "Div"
    wsStandings.Range("Div").Formula = _
        "=INDEX('" & ThisWorkbook.Name & "'!TeamDiv,MATCH(B2,'" & ThisWorkbook.Name & "'!TeamName,0))"
         ' Static values
    wsStandings.Range("Div").Copy
    wsStandings.Range("Div").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    ' Add conference standings
    wsStandings.Range("ConfOrder").Formula = "=IF(" & ColLetter("Conf", wsStandings) & "2=" & ColLetter("Conf", wsStandings) & "1," & _
        ColLetter("ConfOrder", wsStandings) & "1+1,1)"
    
    ' Add formulas
        ' Games played and Pts, in case user changes W, L, or OT
    wsStandings.Range("GP_").Formula = "=Wins+Losses+OT_"
    wsStandings.Range("Points").Formula = "=Wins*2+OT_"
        ' PPG
    wsStandings.Cells(1, wbStandings.Names("PPG_").RefersToRange.Column).Formula = "Win%"
    wsStandings.Range("PPG_").Formula = "=Points/GP_/2"
    wsStandings.Range("PPG_").NumberFormat = "0.000"
        ' Projected points
    wsStandings.Cells(1, wbStandings.Names("Proj").RefersToRange.Column).Formula = "Proj"
    wsStandings.Range("Proj").Formula = "=ROUND(PPG_*164,0)"
        ' League rank
    wsStandings.Cells(1, wbStandings.Names("League").RefersToRange.Column).Formula = "League"
    wsStandings.Range("League").Formula = "=COUNTIFS(PPG_,"">""&PPG_)+COUNTIFS(PPG_,PPG_,ROW_PG,"">""&ROW_PG)+" & _
        "COUNTIFS(PPG_,PPG_,ROW_PG,ROW_PG,Diff_PG,"">""&Diff_PG)+COUNTIFS(PPG_,PPG_,ROW_PG,ROW_PG,Diff_PG,Diff_PG,Teams,"">""&Teams)+1"
    With wsStandings.Range(ColLetter("League", wsStandings) & ":" & ColLetter("League", wsStandings))
        .HorizontalAlignment = xlCenter
        .Font.Size = 9
        .Font.Italic = True
        .ColumnWidth = 6
    End With
        ' Goal differential
    wsStandings.Cells(1, wbStandings.Names("Diff").RefersToRange.Column).Formula = "Diff"
    wsStandings.Range("Diff").Formula = "=GF_-GA_"

        ' Projected playoff requirement
    wsStandings.Cells(1, wbStandings.Names("Playoffs").RefersToRange.Column).Formula = "Playoffs"
    wsStandings.Range("Playoffs").Formula = "=IFS(ClinchIn,""* IN *"",ClinchOut,""out"",PPct_Calc>1,IF(PlayoffsPoints,164,1),PPct_Calc<0,0,TRUE,PPct_Calc*IF(PlayoffsPoints,164,1))"
    With wsStandings.Range(ColLetter("Playoffs", wsStandings) & ":" & ColLetter("Playoffs", wsStandings))
        .HorizontalAlignment = xlRight
        .ColumnWidth = 8
    End With
    
        ' various rankings
    wsStandings.Cells(1, wbStandings.Names("GF_Rank").RefersToRange.Column).Formula = "#"
    wsStandings.Range("GF_Rank").Formula = "=IF(LeagueWide,RANK.EQ(GF_PG,GF_PG),COUNTIFS(Conf,Conf,GF_PG,"">""&GF_PG)+1)"
    wsStandings.Cells(1, wbStandings.Names("GA_Rank").RefersToRange.Column).Formula = "#"
    wsStandings.Range("GA_Rank").Formula = "=IF(LeagueWide,RANK.EQ(GA_PG,GA_PG,1),COUNTIFS(Conf,Conf,GA_PG,""<""&GA_PG)+1)"
    wsStandings.Cells(1, wbStandings.Names("Diff_Rank").RefersToRange.Column).Formula = "#"
    wsStandings.Range("Diff_Rank").Formula = "=IF(LeagueWide,RANK.EQ(Diff_PG,Diff_PG,0),COUNTIFS(Conf,Conf,Diff_PG,"">""&Diff_PG)+1)"
    wsStandings.Cells(1, wbStandings.Names("Home_Rank").RefersToRange.Column).Formula = "#"
    wsStandings.Range("Home_Rank").Formula = "=IF(LeagueWide,RANK.EQ(H_PG,H_PG,0),COUNTIFS(Conf,Conf,H_PG,"">""&H_PG)+1)"
    wsStandings.Cells(1, wbStandings.Names("Away_Rank").RefersToRange.Column).Formula = "#"
    wsStandings.Range("Away_Rank").Formula = "=IF(LeagueWide,RANK.EQ(A_PG,A_PG,0),COUNTIFS(Conf,Conf,A_PG,"">""&A_PG)+1)"
        
    wsStandings.Cells(1, wbStandings.Names("L10Change").RefersToRange.Column).Formula = "L10 Chg"
    wsStandings.Range("L10Change").Formula = "=IF(LeagueWide,RANK.EQ(L10_PPG,L10_PPG)-League," & _
        "COUNTIFS(Conf,Conf,L10_PPG,"">""&L10_PPG)-COUNTIFS(Conf,Conf,PPG_,"">""&PPG_))"
        
        ' New column headings & Formulas used to check in/out of playoffs and other stuff
    wsStandings.Cells(1, wbStandings.Names("Teams").RefersToRange.Column).Formula = "Team"
    wsStandings.Cells(1, wbStandings.Names("L10Change").RefersToRange.Column).Formula = "L10 Chg"
    wsStandings.Cells(1, wbStandings.Names("W_PG").RefersToRange.Column).Formula = "W_PG"
    wsStandings.Range("W_PG").Formula = "=IF(GP_=0,0,Wins/GP_)"
    wsStandings.Cells(1, wbStandings.Names("ROW_PG").RefersToRange.Column).Formula = "ROW_PG"
    wsStandings.Range("ROW_PG").Formula = "=IF(ROW_=0,0,ROW_/GP_)"
    wsStandings.Cells(1, wbStandings.Names("GF_PG").RefersToRange.Column).Formula = "GF_PG"
    wsStandings.Range("GF_PG").Formula = "=IF(GP_=0,0,GF_/GP_)"
    wsStandings.Cells(1, wbStandings.Names("GA_PG").RefersToRange.Column).Formula = "GA_PG"
    wsStandings.Range("GA_PG").Formula = "=IF(GP_=0,0,GA_/GP_)"
    wsStandings.Cells(1, wbStandings.Names("Diff_PG").RefersToRange.Column).Formula = "Diff_PG"
    wsStandings.Range("Diff_PG").Formula = "=IF(GP_=0,0,Diff/GP_)"
    wsStandings.Cells(1, wbStandings.Names("H_PG").RefersToRange.Column).Formula = "H_PG"
    wsStandings.Range("H_PG").Formula = "=IFERROR((LEFT(Home,FIND(""-"",Home)-1)*2+MID(Home,FIND(""-"",Home,FIND(""-"",Home)+1)+1,LEN(Home)))/(LEFT(Home,FIND(""-"",Home)-1)+MID(Home,FIND(""-"",Home)+1,FIND(""-"",Home,FIND(""-"",Home)+1)-FIND(""-"",Home)-1)+MID(Home,FIND(""-"",Home,FIND(""-"",Home)+1)+1,LEN(Home))),0)"
    wsStandings.Cells(1, wbStandings.Names("A_PG").RefersToRange.Column).Formula = "A_PG"
    wsStandings.Range("A_PG").Formula = "=IFERROR((LEFT(Away,FIND(""-"",Away)-1)*2+MID(Away,FIND(""-"",Away,FIND(""-"",Away)+1)+1,LEN(Away)))/(LEFT(Away,FIND(""-"",Away)-1)+MID(Away,FIND(""-"",Away)+1,FIND(""-"",Away,FIND(""-"",Away)+1)-FIND(""-"",Away)-1)+MID(Away,FIND(""-"",Away,FIND(""-"",Away)+1)+1,LEN(Away))),0)"
    wsStandings.Cells(1, wbStandings.Names("Div_Top3").RefersToRange.Column).Formula = "Div_Top3"
    wsStandings.Range("Div_Top3").Formula = "=IF(COUNTIFS(Div,Div,League,""<""&League)<3,1,0)"
    
    wsStandings.Cells(1, wbStandings.Names("Max_ROW").RefersToRange.Column).Formula = "Max_ROW"
    wsStandings.Range("Max_ROW").Formula = "=ROW_+(" & SeasonGames & "-GP_)"
    wsStandings.Cells(1, wbStandings.Names("Max_Pts").RefersToRange.Column).Formula = "Max_Pts"
    wsStandings.Range("Max_Pts").Formula = "=Points+(" & SeasonGames & "-GP_)*2"
        
        ' PPG of 8th and 9th in conference, and 3rd and 4th in division have to be entered as array formula
    wsStandings.Cells(1, wbStandings.Names("Conf8th").RefersToRange.Column).Formula = "Conf8th"
    wsStandings.Cells(2, wbStandings.Names("Conf8th").RefersToRange.Column).FormulaArray = _
        "=LARGE(IF(Conf=" & ColLetter("Conf", wsStandings) & "2,PPG_),8)"
    wsStandings.Cells(2, wbStandings.Names("Conf8th").RefersToRange.Column).Copy
    wsStandings.Range(wsStandings.Cells(3, wbStandings.Names("Conf8th").RefersToRange.Column), _
        wsStandings.Cells(32, wbStandings.Names("Conf8th").RefersToRange.Column)).PasteSpecial Paste:=xlPasteFormulas
    Application.CutCopyMode = False
    
    wsStandings.Cells(1, wbStandings.Names("Conf9th").RefersToRange.Column).Formula = "Conf9th"
    wsStandings.Cells(2, wbStandings.Names("Conf9th").RefersToRange.Column).FormulaArray = _
        "=LARGE(IF(Conf=" & ColLetter("Conf", wsStandings) & "2,PPG_),9)"
    wsStandings.Cells(2, wbStandings.Names("Conf9th").RefersToRange.Column).Copy
    wsStandings.Range(wsStandings.Cells(3, wbStandings.Names("Conf9th").RefersToRange.Column), _
        wsStandings.Cells(32, wbStandings.Names("Conf9th").RefersToRange.Column)).PasteSpecial Paste:=xlPasteFormulas
    Application.CutCopyMode = False

    wsStandings.Cells(1, wbStandings.Names("Div3rd").RefersToRange.Column).Formula = "Div3rd"
    wsStandings.Cells(2, wbStandings.Names("Div3rd").RefersToRange.Column).FormulaArray = _
        "=LARGE(IF(Div=" & ColLetter("Div", wsStandings) & "2,PPG_),3)"
    wsStandings.Cells(2, wbStandings.Names("Div3rd").RefersToRange.Column).Copy
    wsStandings.Range(wsStandings.Cells(3, wbStandings.Names("Div3rd").RefersToRange.Column), _
        wsStandings.Cells(32, wbStandings.Names("Div3rd").RefersToRange.Column)).PasteSpecial Paste:=xlPasteFormulas
    Application.CutCopyMode = False
    
    wsStandings.Cells(1, wbStandings.Names("Div4th").RefersToRange.Column).Formula = "Div4th"
    wsStandings.Cells(2, wbStandings.Names("Div4th").RefersToRange.Column).FormulaArray = _
        "=LARGE(IF(Div=" & ColLetter("Div", wsStandings) & "2,PPG_),4)"
    wsStandings.Cells(2, wbStandings.Names("Div4th").RefersToRange.Column).Copy
    wsStandings.Range(wsStandings.Cells(3, wbStandings.Names("Div4th").RefersToRange.Column), _
        wsStandings.Cells(32, wbStandings.Names("Div4th").RefersToRange.Column)).PasteSpecial Paste:=xlPasteFormulas
    Application.CutCopyMode = False
    
    wsStandings.Cells(1, wbStandings.Names("InPlayoffs").RefersToRange.Column).Formula = "InPlayoffs"
    wsStandings.Range("InPlayoffs").Formula = "=OR(Div_Top3=1,COUNTIFS(Div_Top3,""<>1"",Conf,Conf,League,""<""&League)<2)"
    wsStandings.Cells(1, wbStandings.Names("PPct_Calc").RefersToRange.Column).Formula = "PPct_Calc"
    wsStandings.Range("PPct_Calc").Formula = "=IF(GP_=82,0,((MIN(MINIFS(IF(InPlayoffs,Conf9th,Conf8th),Conf,Conf),MINIFS(IF(InPlayoffs,Div4th,Div3rd),Div,Div))+0.001)*164-Points)/(82-GP_)/2)"
    
    wsStandings.Cells(1, wbStandings.Names("ClinchIn").RefersToRange.Column).Formula = "ClinchIn"
    wsStandings.Range("ClinchIn").Formula = "=OR((COUNTIFS(Conf,Conf,Max_Pts,"">""&Points)+COUNTIFS(Conf,Conf,Max_Pts,Points,ROW_,"">""&Max_ROW)+COUNTIFS(Conf,Conf,Max_Pts,Points,ROW_,Max_ROW,GP_,82,Diff_PG,"">""&Diff_PG))<=7," & _
        "(COUNTIFS(Div,Div,Max_Pts,"">""&Points)+COUNTIFS(Div,Div,Max_Pts,Points,ROW_,"">""&Max_ROW)+COUNTIFS(Div,Div,Max_Pts,Points,ROW_,Max_ROW,GP_,82,Diff_PG,"">""&Diff_PG))<3)"
    wsStandings.Cells(1, wbStandings.Names("ClinchOut").RefersToRange.Column).Formula = "ClinchOut"
    wsStandings.Range("ClinchOut").Formula = "=AND((COUNTIFS(Conf,Conf,Points,"">""&Max_Pts)+COUNTIFS(Conf,Conf,Points,Max_Pts,ROW_,"">""&Max_ROW)+COUNTIFS(Conf,Conf,Points,Max_Pts,ROW_,Max_ROW,GP_,82,Diff_PG,"">""&Diff_PG))>7," & _
        "(COUNTIFS(Div,Div,Points,"">""&Max_Pts)+COUNTIFS(Div,Div,Points,Max_Pts,ROW_,"">""&Max_ROW)+COUNTIFS(Div,Div,Points,Max_Pts,ROW_,Max_ROW,GP_,82,Diff_PG,"">""&Diff_PG))>3)"

    ' Calculate PPG before the previous 10 games
    wsStandings.Cells(1, wbStandings.Names("L10_PPG").RefersToRange.Column).Formula = "L10_PPG"
    wsStandings.Range("L10_PPG").Formula = "=IFERROR((Points-(LEFT(Last10,FIND(""-"",Last10)-1)*2+MID(Last10,FIND(""-""," & _
        "Last10,FIND(""-"",Last10)+1)+1,LEN(Last10))))/(GP_-10),0)"

    ' Check whether the teams are in sorted order
    wsStandings.Cells(1, wbStandings.Names("NeedSort").RefersToRange.Column).Formula = "NeedSort"
    wsStandings.Range("NeedSort").Formula = "=IF(ISBLANK(AQ3),0,IF(IF(ConfSort,W2<W3,W2>W3),0,IF(AF2>AF3,0,IF(J2<J3,0,1))))"
    
End Sub
