Attribute VB_Name = "Module1"
Dim CurrentMatch As Double
Dim SheetName As String, HomeTeamSheet As String, AwayTeamSheet As String, Home5Sheet As String, Away5Sheet As String
Dim PointsDifferenceRow As Double
Dim Oddscheck As Boolean, LeagueCheck As Boolean, InterestCheck As Boolean, InplayCheck As Boolean, MinimumOddsCheck As Boolean, MinimumELOCheck As Boolean, MinimumOddsVarCheck As Boolean, GoodMatch As Boolean, PredictionCheck As Boolean
Dim CurrentTime As Date, StartTimeCheck As Date, EndTimeCheck As Date, MatchTime As Date

Option Explicit

Sub Next_Button()
Dim League As String

CurrentMatch = Sheets("Dashboard").Range("CurrentMatchRow")
Application.Calculation = xlManual 'used to speed up the calculation
CurrentMatch = CurrentMatch + 1
copy_values
Sheets("Dashboard").Range("CurrentMatchRow") = CurrentMatch
Application.Calculation = xlAutomatic 'used to speed up the calculation
End Sub

Sub Previous_Button()
Dim League As String

CurrentMatch = Sheets("Dashboard").Range("CurrentMatchRow")
Application.Calculation = xlManual 'used to speed up the calculation
CurrentMatch = CurrentMatch - 1
copy_values
Sheets("Dashboard").Range("CurrentMatchRow") = CurrentMatch
Application.Calculation = xlAutomatic 'used to speed up the calculation
End Sub

Public Sub copy_values()
SheetName = "Data"

Sheets("Dashboard").Range("MatchDate") = Sheets(SheetName).Range("S_Date")(CurrentMatch)
Sheets("Dashboard").Range("SeasonName") = Sheets(SheetName).Range("S_Season")(CurrentMatch)
Sheets("Dashboard").Range("LeagueName") = Sheets(SheetName).Range("S_League")(CurrentMatch)
Sheets("Dashboard").Range("MatchId") = Sheets(SheetName).Range("S_MatchId")(CurrentMatch)
Sheets("Dashboard").Range("HomeTeam") = Sheets(SheetName).Range("S_HomeTeam")(CurrentMatch)
Sheets("Dashboard").Range("AwayTeam") = Sheets(SheetName).Range("S_AwayTeam")(CurrentMatch)
Sheets("Dashboard").Range("HomeProb") = Sheets(SheetName).Range("S_prob1")(CurrentMatch)
Sheets("Dashboard").Range("DrawProb") = Sheets(SheetName).Range("S_probtie")(CurrentMatch)
Sheets("Dashboard").Range("AwayProb") = Sheets(SheetName).Range("S_prob2")(CurrentMatch)

Dim LeagueURL As String, LeagueName As String
Dim PvtTblHG  As PivotTable, PvtTblAG As PivotTable, PvtTbl6HG As PivotTable, PvtTbl6AG As PivotTable, PvtTblHGH As PivotTable, PvtTblAGA As PivotTable
Dim PvtTbl6SHG As PivotTable, PvtTbl6SAG As PivotTable
Dim PvtItm      As PivotItem



HomeTeamSheet = "Home"
AwayTeamSheet = "Away"
Home5Sheet = "Home 6"
Away5Sheet = "Away 6"

' set the Pivot Table
Set PvtTblHG = Worksheets(HomeTeamSheet).PivotTables("Games")
Set PvtTblHGH = Worksheets(HomeTeamSheet).PivotTables("Games_H")
Set PvtTblAG = Worksheets(AwayTeamSheet).PivotTables("Games")
Set PvtTblAGA = Worksheets(AwayTeamSheet).PivotTables("Games_A")
Set PvtTbl6HG = Worksheets(Home5Sheet).PivotTables("OverallGames")
Set PvtTbl6AG = Worksheets(Away5Sheet).PivotTables("OverallGames")
Set PvtTbl6SHG = Worksheets(Home5Sheet).PivotTables("SideGames")
Set PvtTbl6SAG = Worksheets(Away5Sheet).PivotTables("SideGames")

Dim rng1 As Range, rng2 As Range
Dim dblMin1 As Double, dblMax1 As Double, dblMin2 As Double, dblMax2 As Double

'Set range from which to determine smallest value
Set rng1 = Sheets(HomeTeamSheet).Range("B8:B13")
Set rng2 = Sheets(HomeTeamSheet).Range("S7:S12")



PvtTblHG.ManualUpdate = True
PvtTblAG.ManualUpdate = True
PvtTblHGH.ManualUpdate = True
PvtTblAGA.ManualUpdate = True
PvtTbl6HG.ManualUpdate = True
PvtTbl6AG.ManualUpdate = True
PvtTbl6SHG.ManualUpdate = True
PvtTbl6SAG.ManualUpdate = True

With PvtTblHG
    .PivotFields("season").ClearAllFilters ' <-- clear all filters to "season"
    .PivotFields("date").ClearAllFilters ' <-- clear all filters to "opp_team"
    .PivotFields("league").ClearAllFilters ' <-- clear all filters to "opp_team"
    .PivotFields("team").ClearAllFilters ' <-- clear all filters to "team"
    .PivotFields("date").PivotFilters.Add Type:=xlBefore, Value1:=CLng(Sheets("Dashboard").Range("MatchDate").Value)
    .PivotFields("league").PivotFilters.Add Type:=xlCaptionEquals, Value1:=Sheets("Dashboard").Range("LeagueName").Value
    .PivotFields("team").PivotFilters.Add Type:=xlCaptionEquals, Value1:=Sheets("Dashboard").Range("HomeTeam").Value

    For Each PvtItm In .PivotFields("season").PivotItems
            If PvtItm.Name = Sheets("Dashboard").Range("SeasonName").Value Then
                PvtItm.Visible = True
            Else
                PvtItm.Visible = False
            End If
        Next PvtItm
End With

With PvtTblHGH
    .PivotFields("season").ClearAllFilters ' <-- clear all filters to "season"
    .PivotFields("date").ClearAllFilters ' <-- clear all filters to "opp_team"
    .PivotFields("league").ClearAllFilters ' <-- clear all filters to "opp_team"
    .PivotFields("team").ClearAllFilters ' <-- clear all filters to "team"
    .PivotFields("date").PivotFilters.Add Type:=xlBefore, Value1:=CLng(Sheets("Dashboard").Range("MatchDate").Value)
    .PivotFields("league").PivotFilters.Add Type:=xlCaptionEquals, Value1:=Sheets("Dashboard").Range("LeagueName").Value
    .PivotFields("team").PivotFilters.Add Type:=xlCaptionEquals, Value1:=Sheets("Dashboard").Range("HomeTeam").Value

    For Each PvtItm In .PivotFields("season").PivotItems
            If PvtItm.Name = Sheets("Dashboard").Range("SeasonName").Value Then
                PvtItm.Visible = True
            Else
                PvtItm.Visible = False
            End If
        Next PvtItm
End With

With PvtTblAG
    .PivotFields("season").ClearAllFilters ' <-- clear all filters to "season"
    .PivotFields("date").ClearAllFilters ' <-- clear all filters to "opp_team"
    .PivotFields("league").ClearAllFilters ' <-- clear all filters to "opp_team"
    .PivotFields("team").ClearAllFilters ' <-- clear all filters to "team"
    .PivotFields("date").PivotFilters.Add Type:=xlBefore, Value1:=CLng(Sheets("Dashboard").Range("MatchDate").Value)
    .PivotFields("league").PivotFilters.Add Type:=xlCaptionEquals, Value1:=Sheets("Dashboard").Range("LeagueName").Value
    .PivotFields("team").PivotFilters.Add Type:=xlCaptionEquals, Value1:=Sheets("Dashboard").Range("AwayTeam").Value
        
        For Each PvtItm In .PivotFields("season").PivotItems
            If PvtItm.Name = Sheets("Dashboard").Range("SeasonName").Value Then
                PvtItm.Visible = True
            Else
                PvtItm.Visible = False
            End If
        Next PvtItm
End With
With PvtTblAGA
    .PivotFields("season").ClearAllFilters ' <-- clear all filters to "season"
    .PivotFields("date").ClearAllFilters ' <-- clear all filters to "opp_team"
    .PivotFields("league").ClearAllFilters ' <-- clear all filters to "opp_team"
    .PivotFields("team").ClearAllFilters ' <-- clear all filters to "team"
    .PivotFields("date").PivotFilters.Add Type:=xlBefore, Value1:=CLng(Sheets("Dashboard").Range("MatchDate").Value)
    .PivotFields("league").PivotFilters.Add Type:=xlCaptionEquals, Value1:=Sheets("Dashboard").Range("LeagueName").Value
    .PivotFields("team").PivotFilters.Add Type:=xlCaptionEquals, Value1:=Sheets("Dashboard").Range("AwayTeam").Value
        
        For Each PvtItm In .PivotFields("season").PivotItems
            If PvtItm.Name = Sheets("Dashboard").Range("SeasonName").Value Then
                PvtItm.Visible = True
            Else
                PvtItm.Visible = False
            End If
        Next PvtItm
End With
PvtTblHG.ManualUpdate = False
PvtTblAG.ManualUpdate = False

'Worksheet function MIN returns the smallest value in a range

PvtTblHGH.ManualUpdate = False
PvtTblAGA.ManualUpdate = False
dblMin1 = Application.WorksheetFunction.Min(rng1)
dblMax1 = Application.WorksheetFunction.Max(rng1)
dblMin2 = Application.WorksheetFunction.Min(rng2)
dblMax2 = Application.WorksheetFunction.Max(rng2)




With PvtTbl6HG
    .PivotFields("date").ClearAllFilters ' <-- clear all filters to "opp_team"
    .PivotFields("league").ClearAllFilters ' <-- clear all filters to "opp_team"
    .PivotFields("team").ClearAllFilters ' <-- clear all filters to "team"
    .PivotFields("match_id").ClearAllFilters ' <-- clear all filters to "team"
    .PivotFields("match_id").PivotFilters.Add Type:=xlCaptionIsBetween, Value1:=CLng(dblMin1), Value2:=CLng(dblMax1)
    .PivotFields("league").PivotFilters.Add Type:=xlCaptionEquals, Value1:=Sheets("Dashboard").Range("LeagueName").Value
    .PivotFields("team").PivotFilters.Add Type:=xlCaptionEquals, Value1:=Sheets("Dashboard").Range("HomeTeam").Value
End With
With PvtTbl6SHG
    .PivotFields("date").ClearAllFilters ' <-- clear all filters to "opp_team"
    .PivotFields("league").ClearAllFilters ' <-- clear all filters to "opp_team"
    .PivotFields("team").ClearAllFilters ' <-- clear all filters to "team"
    .PivotFields("match_id").ClearAllFilters ' <-- clear all filters to "team"
    .PivotFields("match_id").PivotFilters.Add Type:=xlCaptionIsBetween, Value1:=CLng(dblMin2), Value2:=CLng(dblMax2)
    .PivotFields("league").PivotFilters.Add Type:=xlCaptionEquals, Value1:=Sheets("Dashboard").Range("LeagueName").Value
    .PivotFields("team").PivotFilters.Add Type:=xlCaptionEquals, Value1:=Sheets("Dashboard").Range("HomeTeam").Value
End With

PvtTbl6HG.ManualUpdate = False
PvtTbl6SHG.ManualUpdate = False

'Set range from which to determine smallest value
Set rng1 = Sheets(AwayTeamSheet).Range("B8:B13")
Set rng2 = Sheets(AwayTeamSheet).Range("S7:S12")
dblMin1 = Application.WorksheetFunction.Min(rng1)
dblMax1 = Application.WorksheetFunction.Max(rng1)
dblMin2 = Application.WorksheetFunction.Min(rng2)
dblMax2 = Application.WorksheetFunction.Max(rng2)
With PvtTbl6AG
    .PivotFields("date").ClearAllFilters ' <-- clear all filters to "opp_team"
    .PivotFields("league").ClearAllFilters ' <-- clear all filters to "opp_team"
    .PivotFields("team").ClearAllFilters ' <-- clear all filters to "team"
    .PivotFields("match_id").ClearAllFilters ' <-- clear all filters to "team"
    .PivotFields("match_id").PivotFilters.Add Type:=xlCaptionIsBetween, Value1:=CLng(dblMin1), Value2:=CLng(dblMax1)
    .PivotFields("league").PivotFilters.Add Type:=xlCaptionEquals, Value1:=Sheets("Dashboard").Range("LeagueName").Value
    .PivotFields("team").PivotFilters.Add Type:=xlCaptionEquals, Value1:=Sheets("Dashboard").Range("AwayTeam").Value
End With
With PvtTbl6SAG
    .PivotFields("date").ClearAllFilters ' <-- clear all filters to "opp_team"
    .PivotFields("league").ClearAllFilters ' <-- clear all filters to "opp_team"
    .PivotFields("team").ClearAllFilters ' <-- clear all filters to "team"
    .PivotFields("match_id").ClearAllFilters ' <-- clear all filters to "team"
    .PivotFields("match_id").PivotFilters.Add Type:=xlCaptionIsBetween, Value1:=CLng(dblMin2), Value2:=CLng(dblMax2)
    .PivotFields("league").PivotFilters.Add Type:=xlCaptionEquals, Value1:=Sheets("Dashboard").Range("LeagueName").Value
    .PivotFields("team").PivotFilters.Add Type:=xlCaptionEquals, Value1:=Sheets("Dashboard").Range("AwayTeam").Value
End With
PvtTbl6AG.ManualUpdate = False
PvtTbl6SAG.ManualUpdate = False
Sheets("Dashboard").Range("HomeTotalHomeGames") = PvtTblHG.GetPivotData("Games", "side", "home") 'Total Home Games Home Team
Sheets("Dashboard").Range("HomeTotalAwayGames") = PvtTblHG.GetPivotData("Games", "side", "away")         'Total Away Games Home Team
Sheets("Dashboard").Range("AwayTotalHomeGames") = PvtTblAG.GetPivotData("Games", "side", "home")         'Total Home Games Away Team
Sheets("Dashboard").Range("AwayTotalAwayGames") = PvtTblAG.GetPivotData("Games", "side", "away")          'Total Away Games Away Team

Sheets("Dashboard").Range("HomeTotalHomeGoals") = PvtTblHG.GetPivotData("Scores", "side", "home")         'Total Home Goals Home Team
Sheets("Dashboard").Range("HomeTotalAwayGoals") = PvtTblHG.GetPivotData("Scores", "side", "away")           'naming error, ignore.Total Away Goals Home Team
Sheets("Dashboard").Range("AwayTotalHomeGoals") = PvtTblAG.GetPivotData("Scores", "side", "home")            'Total Home Goals Away Team
Sheets("Dashboard").Range("AwayTotalAwayGoals") = PvtTblAG.GetPivotData("Scores", "side", "away")           'Total Away Goals Away Team

Sheets("Dashboard").Range("HomeTotalSideFTS") = PvtTblHG.GetPivotData("FTS", "side", "home") / PvtTblHG.GetPivotData("Games", "side", "home")    'Total Home FTS % Home Team
Sheets("Dashboard").Range("AwayTotalSideFTS") = PvtTblAG.GetPivotData("FTS", "side", "away") / PvtTblAG.GetPivotData("Games", "side", "away") 'Total Away FTS % Away Team

Sheets("Dashboard").Range("HomeTotalSideCS") = PvtTblHG.GetPivotData("CS", "side", "home") / PvtTblHG.GetPivotData("Games", "side", "home")    'Total Home CS % Home Team
Sheets("Dashboard").Range("AwayTotalSideCS") = PvtTblAG.GetPivotData("CS", "side", "away") / PvtTblAG.GetPivotData("Games", "side", "away") 'Total Away CS % Away Team

Sheets("Dashboard").Range("HomeTotalConcedeHome") = PvtTblHG.GetPivotData("Concede", "side", "home")           'Total Home Conceded Home Team
Sheets("Dashboard").Range("HomeTotalConcedeAway") = PvtTblHG.GetPivotData("Concede", "side", "away")            'Total Away Conceded Home Team
Sheets("Dashboard").Range("AwayTotalConcedeHome") = PvtTblAG.GetPivotData("Concede", "side", "home")            'Total Home Conceded Away Team
Sheets("Dashboard").Range("AwayTotalConcedeAway") = PvtTblAG.GetPivotData("Concede", "side", "away")           'Total Away Conceded Away Team
Sheets("Dashboard").Range("Home5TotalGoals") = PvtTbl6HG.GetPivotData("Scores")           'Last 5 Goals Scored - All - Home Team
Sheets("Dashboard").Range("Away5TotalGoals") = PvtTbl6AG.GetPivotData("Scores")          'Last 5 Goals Scored - All - Away Team

Sheets("Dashboard").Range("Home5TotalConcede") = PvtTbl6HG.GetPivotData("Concede")          'Last 5 Goals Conceded - All - Home Team
Sheets("Dashboard").Range("Away5TotalConcede") = PvtTbl6AG.GetPivotData("Concede")          'Last 5 Goals Conceded - All - Away Team

Sheets("Dashboard").Range("Home5TotalFTS") = PvtTbl6HG.GetPivotData("FTS") / PvtTbl6HG.GetPivotData("Games")        'Last 5 Home FTS % - All - Home Team
Sheets("Dashboard").Range("Away5TotalFTS") = PvtTbl6AG.GetPivotData("FTS") / PvtTbl6AG.GetPivotData("Games")      'Last 5 Away FTS % Away Team

Sheets("Dashboard").Range("Home5TotalCS") = PvtTbl6HG.GetPivotData("CS") / PvtTbl6HG.GetPivotData("Games")       'Last 5 Home CS % - All - Home Team
Sheets("Dashboard").Range("Away5TotalCS") = PvtTbl6AG.GetPivotData("CS") / PvtTbl6AG.GetPivotData("Games")      'Last 5 Away CS % - All - Away Team

Sheets("Dashboard").Range("Home5SideGoals") = PvtTbl6SHG.GetPivotData("Scores", "side", "home")           'Last 5 Goals Scored - Home - Home Team
Sheets("Dashboard").Range("Away5SideGoals") = PvtTbl6SAG.GetPivotData("Scores", "side", "away")           'Last 5 Goals Scored - Away - Away Team

Sheets("Dashboard").Range("Home5SideConcede") = PvtTbl6SHG.GetPivotData("Concede", "side", "home")           'Last 5 Goals Conceded - Home - Home Team
Sheets("Dashboard").Range("Away5SideConcede") = PvtTbl6SAG.GetPivotData("Concede", "side", "away")           'Last 5 Goals Conceded - Away - Away Team

Sheets("Dashboard").Range("Home5SideFTS") = PvtTbl6SHG.GetPivotData("FTS", "side", "home") / PvtTbl6SHG.GetPivotData("Games", "side", "home")     'Last 5 Home FTS % - All - Home Team
Sheets("Dashboard").Range("Away5SideFTS") = PvtTbl6SAG.GetPivotData("FTS", "side", "away") / PvtTbl6SAG.GetPivotData("Games", "side", "away")     'Last 5 Away FTS % Away Team

Sheets("Dashboard").Range("Home5SideCS") = PvtTbl6SHG.GetPivotData("CS", "side", "home") / PvtTbl6SHG.GetPivotData("Games", "side", "home")        'Last 5 Home CS % - All - Home Team
Sheets("Dashboard").Range("Away5SideCS") = PvtTbl6SAG.GetPivotData("CS", "side", "away") / PvtTbl6SAG.GetPivotData("Games", "side", "away")       'Last 5 Away CS % - All - Away Team

'forms
Sheets("Dashboard").Range("HomeGlobalForm") = PvtTbl6HG.GetPivotData("GForm")
Sheets("Dashboard").Range("AwayGlobalForm") = PvtTbl6AG.GetPivotData("GForm")
Sheets("Dashboard").Range("HomeSideForm") = PvtTbl6SHG.GetPivotData("SForm", "side", "home")
Sheets("Dashboard").Range("AwaySideForm") = PvtTbl6SAG.GetPivotData("SForm", "side", "away")
'Actual Odds
'Sheets("Dashboard").Range("StartHomeOdds") = Sheets(SheetName).Range("S_HomeOdds")(CurrentMatch)
'Sheets("Dashboard").Range("StartDrawOdds") = Sheets(SheetName).Range("S_DrawOdds")(CurrentMatch)
'Sheets("Dashboard").Range("StartAwayOdds") = Sheets(SheetName).Range("S_AwayOdds")(CurrentMatch)
'Score result
Sheets("Dashboard").Range("ScoreResult") = Sheets(SheetName).Range("S_Score1")(CurrentMatch) & "-" & Sheets(SheetName).Range("S_Score2")(CurrentMatch)
'Sheets("Dashboard").Range("ScoreExtra") = Sheets(SheetName).Range("S_Extra")(CurrentMatch)
''urls
On Error Resume Next
Dim urlarray() As String
'LeagueURL = Split(Split(Sheets(SheetName).Range("S_HomeTeam")(CurrentMatch).Formula, "=HYPERLINK(""")(1), """,""")(0)
'Sheets("Dashboard").Range("LeagueURL").Formula = "=HYPERLINK(""" & LeagueURL & """,""" & "League" & """)"
On Error GoTo 0
'league finding
Dim lRow As Long
On Error Resume Next
lRow = Application.WorksheetFunction.Match(Sheets(SheetName).Range("S_League")(CurrentMatch), Sheets("League stats").Range("A2:A5000"), 0)
On Error GoTo 0
If lRow > 0 Then
    Sheets("Dashboard").Range("LeagueRow") = lRow + 2
End If

End Sub


