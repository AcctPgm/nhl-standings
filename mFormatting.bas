Attribute VB_Name = "mFormatting"
Option Explicit

Global clrBlue As Long
Global clrGreen As Long
Global clrPink As Long
Global clrReddish As Long
Global clrYellow As Long
    
Sub Formatting(x)
'
    Dim c As Range
    Dim s
    Dim sColumn As String
    
    ' Status update
    Call ufmStatusMessage.ShowMessage("Formatting")
    
    ' Initialize colors
    clrBlue = RGB(167, 211, 255)
    clrGreen = RGB(226, 240, 218)
    clrPink = RGB(255, 203, 255)
    clrReddish = RGB(255, 51, 153)
    clrYellow = RGB(255, 255, 102)
    
    ' Freeze the title row
    wsStandings.Range("C2").Select
    ActiveWindow.FreezePanes = True
    
    ' Align the column headings
    wsStandings.Range("C1:I1").HorizontalAlignment = xlRight
    wsStandings.Columns("J:L").HorizontalAlignment = xlCenter
    wsStandings.Range("M1,O1,Q1").HorizontalAlignment = xlRight
    
    ' Centre the Conference & Division columns
    wsStandings.Columns(ColLetter("Conf", wsStandings) & ":" & ColLetter("Conf", wsStandings)).HorizontalAlignment = xlCenter
    wsStandings.Columns(ColLetter("Div", wsStandings) & ":" & ColLetter("Div", wsStandings)).HorizontalAlignment = xlCenter
    
    ' Winning % projected to be required to be in playoffs
    wsStandings.Range("Playoffs").NumberFormat = "0.000"
    
    ' Delete any existing conditional formatting (not that there should be any)
    wsStandings.Cells.FormatConditions.Delete
    
    With wsStandings.Range("FormatRange")
        ' Default to pink rows
        .Interior.Color = clrPink
        
        ' Playoff teams in green, conditionally
        .FormatConditions.Add Type:=xlExpression, Formula1:="=$" & ColLetter("InPlayoffs", wsStandings) & "2"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).Interior.Color = clrGreen
        .FormatConditions(1).StopIfTrue = False
    End With

    ' Mark teams that *should* be out of or in the playoffs
    With wsStandings.Range("League")
        ' In playoffs but out of top 16
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND($" & ColLetter("InPlayoffs", wsStandings) & "2,$" & ColLetter("League", wsStandings) & "2>16)"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
            .Pattern = xlPatternRectangularGradient
            .Gradient.RectangleLeft = 0.5
            .Gradient.RectangleRight = 0.5
            .Gradient.RectangleTop = 0.5
            .Gradient.RectangleBottom = 0.5
            .Gradient.ColorStops.Clear
        End With
        .FormatConditions(1).Interior.Gradient.ColorStops.Add(0).Color = clrPink
        .FormatConditions(1).Interior.Gradient.ColorStops.Add(1).Color = clrGreen
        .FormatConditions(1).StopIfTrue = False
    
        ' Out of playoffs but in top 16
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND(NOT($" & ColLetter("InPlayoffs", wsStandings) & "2),$" & ColLetter("League", wsStandings) & "2<=16)"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
            .Pattern = xlPatternRectangularGradient
            .Gradient.RectangleLeft = 0.2
            .Gradient.RectangleRight = 0.1
            .Gradient.RectangleTop = 0.2
            .Gradient.RectangleBottom = 0.1
            .Gradient.ColorStops.Clear
        End With
        .FormatConditions(1).Interior.Gradient.ColorStops.Add(0).Color = clrGreen
        .FormatConditions(1).Interior.Gradient.ColorStops.Add(1).Color = clrPink
        .FormatConditions(1).StopIfTrue = False
    End With

    ' Teams not in top 6 for their conference but that are top 3 of their division
    With wsStandings.Range("Teams")
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND(" & ColLetter("Div_Top3", wsStandings) & "2=1," & _
            "MINIFS(League,Conf," & ColLetter("Conf", wsStandings) & "2,Div_Top3,0)<" & ColLetter("League", wsStandings) & "2)"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
            .Pattern = xlPatternRectangularGradient
            .Gradient.RectangleLeft = 0.2
            .Gradient.RectangleRight = 0.5
            .Gradient.RectangleTop = 0.5
            .Gradient.RectangleBottom = 0.5
            .Gradient.ColorStops.Clear
        End With
        .FormatConditions(1).Interior.Gradient.ColorStops.Add(0).Color = clrBlue
        .FormatConditions(1).Interior.Gradient.ColorStops.Add(1).Color = clrGreen
        .FormatConditions(1).StopIfTrue = False
    End With

    ' Teams not in top 8 for their conference but that are top 3 of their division
    With wsStandings.Range("Teams")
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND(" & ColLetter("Div_Top3", wsStandings) & "2=1," & _
            "MINIFS(League,Conf," & ColLetter("Conf", wsStandings) & "2,InPlayoffs,FALSE)<" & ColLetter("League", wsStandings) & "2)"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
            .Pattern = xlPatternRectangularGradient
            .Gradient.RectangleLeft = 0.2
            .Gradient.RectangleRight = 0.5
            .Gradient.RectangleTop = 0.5
            .Gradient.RectangleBottom = 0.5
            .Gradient.ColorStops.Clear
        End With
        .FormatConditions(1).Interior.Gradient.ColorStops.Add(0).Color = clrPink
        .FormatConditions(1).Interior.Gradient.ColorStops.Add(1).Color = clrGreen
        .FormatConditions(1).StopIfTrue = False
    End With
    
    ' In/out of playoffs
    With wsStandings.Range("Playoffs")
        ' Clinched Playoff Spot
        .FormatConditions.Add Type:=xlExpression, Formula1:="=" & ColLetter("ClinchIn", wsStandings) & "2"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
            .Pattern = xlPatternRectangularGradient
            .Gradient.RectangleLeft = 0.2
            .Gradient.RectangleRight = 0.2
            .Gradient.RectangleTop = 0.3
            .Gradient.RectangleBottom = 0.2
            .Gradient.ColorStops.Clear
        End With
        .FormatConditions(1).Interior.Gradient.ColorStops.Add(0).Color = vbGreen 'RGB(0, 255, 0)
        .FormatConditions(1).Interior.Gradient.ColorStops.Add(1).Color = clrGreen
        .FormatConditions(1).StopIfTrue = False
        
        ' Eliminated from Playoffs
        .FormatConditions.Add Type:=xlExpression, Formula1:="=" & ColLetter("ClinchOut", wsStandings) & "2"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
            .Pattern = xlPatternRectangularGradient
            .Gradient.RectangleLeft = 0.6
            .Gradient.RectangleRight = 0
            .Gradient.RectangleTop = 0.2
            .Gradient.RectangleBottom = 0.2
            .Gradient.ColorStops.Clear
        End With
        .FormatConditions(1).Interior.Gradient.ColorStops.Add(0).Color = clrReddish
        .FormatConditions(1).Interior.Gradient.ColorStops.Add(1).Color = clrPink
        .FormatConditions(1).StopIfTrue = False
        
        ' Ahead of 9th team's projected total - Playoff pace of zero
        .FormatConditions.Add Type:=xlExpression, Formula1:="=(" & ColLetter("Playoffs", wsStandings) & "2=0)"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
            .Pattern = xlPatternRectangularGradient
            .Gradient.RectangleLeft = 0.5
            .Gradient.RectangleRight = 0.2
            .Gradient.RectangleTop = 0.3
            .Gradient.RectangleBottom = 0.2
            .Gradient.ColorStops.Clear
        End With
        .FormatConditions(1).Interior.Gradient.ColorStops.Add(0).Color = RGB(138, 255, 79)
        .FormatConditions(1).Interior.Gradient.ColorStops.Add(1).Color = clrGreen
        .FormatConditions(1).StopIfTrue = False
    
        ' Not able to match 8th team's projected points - Playoff pace of 1.0
        .FormatConditions.Add Type:=xlExpression, Formula1:="=(" & ColLetter("Playoffs", wsStandings) & "2=1)"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
            .Pattern = xlPatternRectangularGradient
            .Gradient.RectangleLeft = 0.5
            .Gradient.RectangleRight = 0.2
            .Gradient.RectangleTop = 0.3
            .Gradient.RectangleBottom = 0.2
            .Gradient.ColorStops.Clear
        End With
        .FormatConditions(1).Interior.Gradient.ColorStops.Add(0).Color = RGB(255, 101, 255)
        .FormatConditions(1).Interior.Gradient.ColorStops.Add(1).Color = clrPink
        .FormatConditions(1).StopIfTrue = False
    End With

    ' Best/worst in rankings columns
    Call BestWorst("GF_Rank", wsStandings)
    Call BestWorst("GA_Rank", wsStandings)
    Call BestWorst("Diff_Rank", wsStandings)
    Call BestWorst("Home_Rank", wsStandings)
    Call BestWorst("Away_Rank", wsStandings)
    
    ' Toggle playoff pace between percentage and points
    With wsStandings.Range("Playoffs")
        .FormatConditions.Add Type:=xlExpression, Formula1:="=PlayoffsPoints"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).NumberFormat = "0"
        .FormatConditions(1).StopIfTrue = False
    End With
    
    ' Hightlight selected teams
    With wsStandings.Range("FormatRange")
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=(INDEX(TeamHighlight,MATCH($B2,TeamName,0))=""y"")"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Bold = True
            .Italic = True
        End With
        .FormatConditions(1).StopIfTrue = False
    End With
    
    ' Format rankings columns
    For Each c In wsStandings.Range("Headings")
        If c.Text = "#" Then
            sColumn = Left(c.Address(True, False), InStr(c.Address(True, False), "$") - 1)
            With wsStandings.Columns(sColumn & ":" & sColumn)
                .HorizontalAlignment = xlCenter
                .Font.Size = 9
                .Font.Italic = True
                .ColumnWidth = 5
            End With
        End If
    Next c
    
    ' Vertical lines between column groups
    For Each s In Array("League", "Playoffs", "ROW_", "GF_Rank", "GA_Rank", "Diff_Rank", "Home_Rank", "Away_Rank")
        With Range(s).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlHairline
        End With
    Next s
    
    ' Horizontal lines
    With wsStandings.Range("FormatRange")
        ' Light line after every four teams (for ease of following stats across the screen)
        .FormatConditions.Add Type:=xlExpression, Formula1:="=(MOD($A2,4)=0)"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Borders(xlBottom)
            .LineStyle = xlContinuous
            .Color = vbRed
            .TintAndShade = 0
            .Weight = xlHairline
        End With
        .FormatConditions(1).StopIfTrue = False
    
        ' Heavier line between conferences
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=($" & ColLetter("Conf", wsStandings) & "2<>$" & ColLetter("Conf", wsStandings) & "3)"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Borders(xlBottom)
            .LineStyle = xlContinuous
            .TintAndShade = 0
            .Weight = xlThin
        End With
        .FormatConditions(1).StopIfTrue = False
    End With
    
    ' Change numbers-as-text into regular numbers
    For Each s In Array("Wins", "Losses", "OT_", "ROW_", "GF_", "GA_")
        For Each c In wsStandings.Range(s)
            c.Value = Val(c.Text)
        Next c
    Next s

    ' Hide calculation columns
    wsStandings.Range("HideRange").EntireColumn.Hidden = True
End Sub
