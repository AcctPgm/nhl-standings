Attribute VB_Name = "mGetFromWebSun"
Option Explicit

Function GetFromWebSun(x) As Boolean
'
    Dim c As Range
    
    Dim wbQuery As Workbook
    Dim wsQuery As Worksheet
    
    ' Create new workbook to hold initial query result
    Set wbQuery = Workbooks.Add
    Set wsQuery = wbQuery.ActiveSheet
    
    ' Status update
    Call ufmStatusMessage.ShowMessage("Refreshing data from edmontonsun.com")
    
    On Error GoTo ErrorPoint
    
    ' Create query for NHL standings from TheHockeyNews.com
    wbQuery.Queries.Add Name:="Table 0", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Web.Page(Web.Contents(""http://scores.edmontonsun.com/nhl/standings_conference.asp""))," & Chr(13) & "" & Chr(10) & "    Data0 = Source{0}[Data]," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(Data0,{{""Column1"", type text}, {""Column2"", type text}, {""Column3"", type text}, {""Column4"", type text}, {""Column5"", type text}, {""Column6"", type text}, {""Column7"", typ" & _
        "e text}, {""Column8"", type text}, {""Column9"", type text}, {""Column10"", type text}, {""Column11"", type text}, {""Column12"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
        
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
    wsStandings.Range("B1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    ' Delete unneeded rows
    wsStandings.Rows("1:2").Delete
    wsStandings.Rows("18:19").Delete
    
    ' Close the query workbook since it isn't needed further
    wbQuery.Close SaveChanges:=False
    
    ' Exit
    GetFromWebSun = True

ExitPoint:
    Exit Function
    
ErrorPoint:
    GetFromWebSun = False
    
    ' Delete the query workbook, if it had been created at the time of the error
    If Not wbQuery Is Nothing Then
        wbQuery.Close SaveChanges:=False
    End If
    
    Resume ExitPoint
End Function

