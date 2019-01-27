Attribute VB_Name = "mStandings"
Option Explicit

Global Const SeasonGames = 82
    
Global wbStandings As Workbook
Global wsStandings As Worksheet

Sub Standings()
'
' Main routine to get NHL standings & reformat them
'
    ' Turn off screen updating for speed & aesthetics
    Application.ScreenUpdating = False
    
    ' Retrieve current standings from website
    If Not GetFromWebSun(0) Then
        MsgBox "Error retrieving data from thehockeynews.com"
        Exit Sub
    End If
    
    ' Add the calculations used
    Call Calculations(0)
    
    ' Copy the teams list from the macro workbook to use for selecting teams to highlight
    ThisWorkbook.Sheets("Teams").Copy After:=wsStandings
    wsStandings.Select
    
    ' Add formatting to the worksheet
    Call Formatting(0)
    
    ' Add form controls
    Call FormControls(0)
        ' Initial values for check boxes
    wsStandings.Range("LeagueWide").Formula = "=FALSE"
    wsStandings.Range("ConfSort").Formula = ThisWorkbook.Names("ConfSort").RefersToRange.Value
    
    ' Sort the list
    Call SortStandings(0)
    
    ' Finish
    wsStandings.Range("C2").Select
    Call ufmStatusMessage.ShowMessage("", , True)
    Application.ScreenUpdating = True
End Sub

