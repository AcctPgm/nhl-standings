Attribute VB_Name = "mSortStandings"
Option Explicit

Sub SortStandings(Optional bUpating As Boolean = True, Optional wsStandings As Worksheet)
Attribute SortStandings.VB_ProcData.VB_Invoke_Func = " \n14"
'
    Dim ConfSort
    Dim ScrnUpdating
    
    ScrnUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    If wsStandings Is Nothing Then
        Set wsStandings = ActiveSheet
    End If
    
    If wsStandings.Range("ConfSort").Value Then
        ConfSort = xlAscending
    Else
        ConfSort = xlDescending
    End If
    wsStandings.Sort.SortFields.Clear
    wsStandings.Sort.SortFields.Add2 Key:=Range("Conf") _
        , SortOn:=xlSortOnValues, Order:=ConfSort, DataOption:=xlSortNormal
    wsStandings.Sort.SortFields.Add2 Key:=Range("Div_Top3") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    wsStandings.Sort.SortFields.Add2 Key:=Range("League") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With wsStandings.Sort
        .SetRange Range("SortRange")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Application.ScreenUpdating = ScrnUpdating
End Sub
