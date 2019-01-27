Attribute VB_Name = "mFormControls"
Option Explicit

Sub FormControls(x)
'
' Add form controls to standings worksheet
'
    Dim c As Range
    Dim btnSort As Button

    ' Add checkboxes
    Call AddCheckBox(wsStandings.Range("LeagueWide"), xlOff)
    wsStandings.Range("LeagueWide").Offset(0, 1).Formula = "Check: show league-wide; Uncheck: show by conference"
    
    Call AddCheckBox(wsStandings.Range("PlayoffsPoints"), xlOff)
    wsStandings.Range("PlayoffsPoints").Offset(0, 1).Formula = "Check: Playoff pace as points; Uncheck: Playoff pace as winning%"
    
    Call AddCheckBox(wsStandings.Range("ConfSort"), xlOff)
    wsStandings.Range("ConfSort").Offset(0, 1).Formula = "Check: East at top; Uncheck: West at top"
    
    ' Add button to re-sort the standings
        ' Reference cell for sort button
    Set c = wsStandings.Range("LeagueWide").Offset(0, 9)
        ' Add button, 2 cells wide & 2 cells high
    Set btnSort = wsStandings.Buttons.Add(c.Left, c.Top, _
        c.Offset(0, 2).Left - c.Left, c.Offset(2, 0).Top - c.Top)
    btnSort.OnAction = "'" & ThisWorkbook.Name & "'!SortStandings"
    btnSort.Characters.Text = "Sort Standings"
End Sub

