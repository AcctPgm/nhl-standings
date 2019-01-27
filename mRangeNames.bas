Attribute VB_Name = "mRangeNames"
Option Explicit

Sub RangeNames(x)
'
' Add range names
'
    ' Checkboxes
    wbStandings.Names.Add Name:="LeagueWide", RefersTo:="=$A$34"
    wbStandings.Names.Add Name:="PlayoffsPoints", RefersTo:="=$A$35"
    wbStandings.Names.Add Name:="ConfSort", RefersTo:="=$A$36"
    
    ' Columns
    wbStandings.Names.Add Name:="ConfOrder", RefersTo:="=$A$2:$A$32"
    wbStandings.Names.Add Name:="Teams", RefersTo:="=$B$2:$B$32"
    wbStandings.Names.Add Name:="GP_", RefersTo:="=$C$2:$C$32"
    wbStandings.Names.Add Name:="Wins", RefersTo:="=$D$2:$D$32"
    wbStandings.Names.Add Name:="Losses", RefersTo:="=$E$2:$E$32"
    wbStandings.Names.Add Name:="OT_", RefersTo:="=$F$2:$F$32"
    wbStandings.Names.Add Name:="Points", RefersTo:="=$G$2:$G$32"
    wbStandings.Names.Add Name:="PPG_", RefersTo:="=$H$2:$H$32"
    wbStandings.Names.Add Name:="Proj", RefersTo:="=$I$2:$I$32"
    wbStandings.Names.Add Name:="League", RefersTo:="=$J$2:$J$32"
    wbStandings.Names.Add Name:="Playoffs", RefersTo:="=$K$2:$K$32"
    wbStandings.Names.Add Name:="ROW_", RefersTo:="=$L$2:$L$32"
    wbStandings.Names.Add Name:="GF_", RefersTo:="=$M$2:$M$32"
    wbStandings.Names.Add Name:="GF_Rank", RefersTo:="=$N$2:$N$32"
    wbStandings.Names.Add Name:="GA_", RefersTo:="=$O$2:$O$32"
    wbStandings.Names.Add Name:="GA_Rank", RefersTo:="=$P$2:$P$32"
    wbStandings.Names.Add Name:="Diff", RefersTo:="=$Q$2:$Q$32"
    wbStandings.Names.Add Name:="Diff_Rank", RefersTo:="=$R$2:$R$32"
    wbStandings.Names.Add Name:="Home", RefersTo:="=$S$2:$S$32"
    wbStandings.Names.Add Name:="Home_Rank", RefersTo:="=$T$2:$T$32"
    wbStandings.Names.Add Name:="Away", RefersTo:="=$U$2:$U$32"
    wbStandings.Names.Add Name:="Away_Rank", RefersTo:="=$V$2:$V$32"
    wbStandings.Names.Add Name:="Conf", RefersTo:="=$W$2:$W$32"
    wbStandings.Names.Add Name:="Div", RefersTo:="=$X$2:$X$32"
    wbStandings.Names.Add Name:="W_PG", RefersTo:="=$Y$2:$Y$32"
    wbStandings.Names.Add Name:="ROW_PG", RefersTo:="=$Z$2:$Z$32"
    wbStandings.Names.Add Name:="GF_PG", RefersTo:="=$AA$2:$AA$32"
    wbStandings.Names.Add Name:="GA_PG", RefersTo:="=$AB$2:$AB$32"
    wbStandings.Names.Add Name:="Diff_PG", RefersTo:="=$AC$2:$AC$32"
    wbStandings.Names.Add Name:="H_PG", RefersTo:="=$AD$2:$AD$32"
    wbStandings.Names.Add Name:="A_PG", RefersTo:="=$AE$2:$AE$32"
    wbStandings.Names.Add Name:="Div_Top3", RefersTo:="=$AF$2:$AF$32"
    wbStandings.Names.Add Name:="Max_ROW", RefersTo:="=$AG$2:$AG$32"
    wbStandings.Names.Add Name:="Max_Pts", RefersTo:="=$AH$2:$AH$32"
    wbStandings.Names.Add Name:="Conf8th", RefersTo:="=$AI$2:$AI$32"
    wbStandings.Names.Add Name:="Conf9th", RefersTo:="=$AJ$2:$AJ$32"
    wbStandings.Names.Add Name:="Div3rd", RefersTo:="=$AK$2:$AK$32"
    wbStandings.Names.Add Name:="Div4th", RefersTo:="=$AL$2:$AL$32"
    wbStandings.Names.Add Name:="InPlayoffs", RefersTo:="=$AM$2:$AM$32"
    wbStandings.Names.Add Name:="PPct_Calc", RefersTo:="=$AN$2:$AN$32"
    wbStandings.Names.Add Name:="ClinchIn", RefersTo:="=$AO$2:$AO$32"
    wbStandings.Names.Add Name:="ClinchOut", RefersTo:="=$AP$2:$AP$32"
    wbStandings.Names.Add Name:="NeedSort", RefersTo:="=$AQ$2:$AQ$32"

    ' Other ranges and named cells
    wbStandings.Names.Add Name:="FormatRange", RefersTo:="=$A$2:$AQ$32"
    wbStandings.Names.Add Name:="SortRange", RefersTo:="=$B$1:$AQ$32"
    wbStandings.Names.Add Name:="HideRange", RefersTo:="=W:AQ"
    wbStandings.Names.Add Name:="Headings", RefersTo:="=$A$1:$AQ$1"
    wbStandings.Names.Add Name:="LastColumn", RefersTo:="=$AQ$1"
End Sub
