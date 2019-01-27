Attribute VB_Name = "mGetFromWebTHN"
Option Explicit

Function GetFromWebTHN(ByVal x) As Boolean
'
    Dim wbQuery As Workbook
    Dim wsQuery As Worksheet
    
    ' Create new workbook to hold initial query result
    Set wbQuery = Workbooks.Add
    Set wsQuery = wbQuery.ActiveSheet
    
    ' Status update
    Call ufmStatusMessage.ShowMessage("Refreshing data from thehockeynews.com")
    
    On Error GoTo ErrorPoint
    
    ' Create query for NHL standings from TheHockeyNews.com
    wbQuery.Queries.Add Name:="Table 0", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Web.Page(Web.Contents(""http://forecaster.thehockeynews.com/standings/overall""))," & Chr(13) & "" & Chr(10) & "    Data0 = Source{0}[Data]," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(Data0,{{""Team"", type text}, {""GP"", Int64.Type}, {""W"", Int64.Type}, {""L"", Int64.Type}, {""OT"", Int64.Type}, {""PTS"", Int64.Type}, {""ROW"", Int64.Type}, {""GF"", Int64.Type}, {" & _
        """GA"", Int64.Type}, {""Home"", type text}, {""Away"", type text}, {""STREAK"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
    With wsQuery.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""Table 0"";Extended Properties=""""" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [Table 0]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "Table_0"
        .Refresh BackgroundQuery:=False
    End With
    
    ' Copy the query results into a new workbook, to avoid disruptive query refreshes
    Set wbStandings = Workbooks.Add
    Set wsStandings = wbStandings.ActiveSheet
    wsStandings.Name = Format(Int(Now()), "mmm_d")
    
    wsQuery.Range("Table_0[#All]").Copy
    wsStandings.Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
'    wsStandings.Range("B1").PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    ' Close the query workbook since it isn't needed further
    wbQuery.Close SaveChanges:=False
    
    ' Exit
    GetFromWebTHN = True

ExitPoint:
    Exit Function
    
ErrorPoint:
    GetFromWebTHN = False
    
    ' Delete the query workbook, if it had been created at the time of the error
    If Not wbQuery Is Nothing Then
        wbQuery.Close SaveChanges:=False
    End If
    
    Resume ExitPoint
End Function
