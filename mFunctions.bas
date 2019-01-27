Attribute VB_Name = "mFunctions"
Option Explicit

Function ColLetter(sRngName As String, Optional ws As Worksheet) As String
'
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    ColLetter = ws.Range(sRngName).Address(True, False)
    ColLetter = Left(ColLetter, InStr(ColLetter, "$") - 1)
End Function

Sub AddCheckBox(rng As Range, Optional v = xlOff)
'
    Dim chk As CheckBox
    
    ' Default state to off if it is provided but neither on nor off
    If (v <> xlOn) And (v <> xlOff) Then
        v = xlOff
    End If
    
    ' Add the checkbox
    Set chk = rng.Parent.CheckBoxes.Add(rng.Left, rng.Top, rng.Width / 2, rng.Height / 2) ' .Select
    chk.Characters.Text = ""
    chk.Value = v
    chk.LinkedCell = rng.Address(True, True, xlA1)
    chk.Display3DShading = False
    
    ' Unlock range cell so that checkbox changes will work
    rng.Locked = False
    
    ' Set the text colour to the cell background colour so that it isn't visible
    rng.Font.Color = rng.Interior.Color
End Sub

Sub BestWorst(sRankRng As String, Optional wsStandings As Worksheet)
'
    If wsStandings Is Nothing Then
        Set wsStandings = ActiveSheet
    End If
    
    With wsStandings.Range(sRankRng)
        ' Best, team in playoff position
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND(" & ColLetter("InPlayoffs", wsStandings) & "2,IF(LeagueWide,MIN(" & sRankRng & ")=" & ColLetter(sRankRng, wsStandings) & "2," & _
            "MINIFS(" & sRankRng & ",Conf," & ColLetter("Conf", wsStandings) & "2)=" & ColLetter(sRankRng, wsStandings) & "2))"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
            .Pattern = xlPatternRectangularGradient
            .Gradient.RectangleLeft = 0.5
            .Gradient.RectangleRight = 0.1
            .Gradient.RectangleTop = 0.3
            .Gradient.RectangleBottom = 0.2
            .Gradient.ColorStops.Clear
        End With
        .FormatConditions(1).Interior.Gradient.ColorStops.Add(0).Color = clrYellow
        .FormatConditions(1).Interior.Gradient.ColorStops.Add(1).Color = clrGreen
        .FormatConditions(1).StopIfTrue = False

        ' Best, team not in playoff position
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND(NOT(" & ColLetter("InPlayoffs", wsStandings) & "2),IF(LeagueWide,MIN(" & sRankRng & ")=" & ColLetter(sRankRng, wsStandings) & "2," & _
            "MINIFS(" & sRankRng & ",Conf," & ColLetter("Conf", wsStandings) & "2)=" & ColLetter(sRankRng, wsStandings) & "2))"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
            .Pattern = xlPatternRectangularGradient
            .Gradient.RectangleLeft = 0.5
            .Gradient.RectangleRight = 0.1
            .Gradient.RectangleTop = 0.3
            .Gradient.RectangleBottom = 0.2
            .Gradient.ColorStops.Clear
        End With
        .FormatConditions(1).Interior.Gradient.ColorStops.Add(0).Color = clrYellow
        .FormatConditions(1).Interior.Gradient.ColorStops.Add(1).Color = clrPink
        .FormatConditions(1).StopIfTrue = False

        ' Worst, team in playoff position
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND(" & ColLetter("InPlayoffs", wsStandings) & "2,IF(LeagueWide,MAX(" & sRankRng & ")=" & ColLetter(sRankRng, wsStandings) & "2," & _
            "MAXIFS(" & sRankRng & ",Conf," & ColLetter("Conf", wsStandings) & "2)=" & ColLetter(sRankRng, wsStandings) & "2))"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
            .Pattern = xlPatternRectangularGradient
            .Gradient.RectangleLeft = 0.5
            .Gradient.RectangleRight = 0.1
            .Gradient.RectangleTop = 0.3
            .Gradient.RectangleBottom = 0.2
            .Gradient.ColorStops.Clear
        End With
        .FormatConditions(1).Interior.Gradient.ColorStops.Add(0).Color = clrBlue
        .FormatConditions(1).Interior.Gradient.ColorStops.Add(1).Color = clrGreen
        .FormatConditions(1).StopIfTrue = False

        ' Worst, team not in playoff position
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND(NOT(" & ColLetter("InPlayoffs", wsStandings) & "2),IF(LeagueWide,MAX(" & sRankRng & ")=" & ColLetter(sRankRng, wsStandings) & "2," & _
            "MAXIFS(" & sRankRng & ",Conf," & ColLetter("Conf", wsStandings) & "2)=" & ColLetter(sRankRng, wsStandings) & "2))"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
            .Pattern = xlPatternRectangularGradient
            .Gradient.RectangleLeft = 0.5
            .Gradient.RectangleRight = 0.1
            .Gradient.RectangleTop = 0.3
            .Gradient.RectangleBottom = 0.2
            .Gradient.ColorStops.Clear
        End With
        .FormatConditions(1).Interior.Gradient.ColorStops.Add(0).Color = clrBlue
        .FormatConditions(1).Interior.Gradient.ColorStops.Add(1).Color = clrPink
        .FormatConditions(1).StopIfTrue = False
    End With
End Sub
