clear
$xl = New-Object -comobject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $False
$pathWorking = "D:\CloudStorage\Dropbox\craighirata@gmail.com\Dropbox\Tennis\ScoreCards"
$pathWorking ="D:\CloudStorage\Dropbox\craighirata@gmail.com\Dropbox\Tennis\ScoreCards"

$xlAutomatic=-4105
$xlBottom = -4107
$xlCenter = -4108
$xlContext = -5002
$xlContinuous=1
$xlDiagonalDown=5
$xlDiagonalUp=6
$xlEdgeBottom=9
$xlEdgeLeft=7
$xlEdgeRight=10
$xlEdgeTop=8
$xlInsideHorizontal=12
$xlInsideVertical=11
$xlNone=-4142
$xlThin=2
$xlMedium = -4138
$xlThick = 4
$xlShiftDown	= -4121

function RGB ($r, $g, $b ){
    $color = [Long]($r + $g * 256 + $b * 256 * 256)
    $color.GetType()
  return $color
}

$red =       RGB 255 199 206  
$darkRed =   RGB 156 0 6
$green =     RGB 198 239 206
$darKGreen = RGB 0 97 0
$intRed =       $red[1]       #13551615
$intDarkRed =   $darkRed[1]   #393372
$intGreen =     $green[1]     #13561798
$intDarkGreen = $darkGreen[1] #24832


$wbDump = $xl.Workbooks.Open("$pathWorking\Fall2018Dump.xlsx")
$wbOut  = $xl.Workbooks.Open("$pathWorking\Fall2018Out.xlsm")

$wsPlayers
$wsScores
$wsFinalScores

$wsFinalScores = $wbDump.worksheets("FinalScores")
$wsPlayers = $wbDump.worksheets("Players")
$wsScores = $wbDump.worksheets("Scores")
$wsDumper  = $wbDump.worksheets("Dumper")
$wsTeams = $wbDump.worksheets("Teams")

$wsPlayersRng = $wsPlayers.UsedRange.Cells 
$wsPlayersRowCount = $wsPlayersRng.Rows.Count
$wsPlayersRowCount

$xlCellTypeLastCell = 11

$rangeUsedTeams = $wsTeams.UsedRange()
$lastTeamRow = $rangeUsedTeams.SpecialCells($xlCellTypeLastCell).row



#create team tabs
for ($row = 2; $row -le $lastTeamRow; $row++)
{
    $teamName = $wsTeams.Cells($row, 2).Text
    $wsNew = $wbOut.worksheets.add()
    $wsNew.activate() 
    $wsNew.Name = $teamName
}

#add players to team tabs
for ($row = 2; $row -le $wsPlayersRowCount ; $row++)
{
    $teamName   = $wsPlayers.Cells.Item($row, 1).Text
    #$teamName
    $wsTeam = $wbOut.worksheets.item($teamName)
    $getRange = $wsPlayers.Range("D$($row):P$($row)")
    $getRange.copy()
    $wsTeam.activate()
    $usedRanger = $wsTeam.UsedRange()
    $lastCell = $usedRanger.SpecialCells($xlCellTypeLastCell)
    $newRow = $lastCell.row + 1
    #$newRow = $usedRange.Rows.count + 1
    $newRow
    $pasteRange = $wsTeam.Range("A$($newRow):M$($newRow)")
    $wsTeam.Paste($pasteRange)    
}
#$eRow = $worksheet.cells.item(1,1).entireRow
#$active = $eRow.activate()
#$active = $eRow.insert($xlShiftDown)

#colorize cells
for ($row = 2; $row -le $lastTeamRow; $row++) #iterate through each team in the Teams sheet
{
    $teamName = $wsTeams.Cells($row, 2).Text
    $wsTeam = $wbOut.worksheets.item($teamName)
    $rangeUsed = $wsTeam.UsedRange()
    $firstCell = $rangeUsed.Cells(1,1)
    $firstRow = $firstCell.EntireRow
    $firstRow.insert($xlShiftDown)
    $lastRow = $rangeUsed.SpecialCells($xlCellTypeLastCell).row
    $lastCol = $rangeUsed.SpecialCells($xlCellTypeLastCell).column
    $lastRow
    $lastCol

    for ($col = 2; $col -le 10; $col+=2)
    {
        $MergeCells = $wsTeam.Range($wsTeam.Cells(1,$col), $wsTeam.Cells(1,$col + 1))
        $MergeCells.Select
        $MergeCells.MergeCells = $true
        $MergeCells = "Test"
    }

    for ($col = 2; $col -lt $lastCol; $col+=2)
    {
        $rangeColor1 = $wsTeam.Range($rangeUsed.Cells(1, $col), $rangeUsed.Cells($lastRow - 1, $col))
        $rangeColor2 = $rangeColor1.Offset(0,1)

        $rangeColor1.Interior.Color =     [Long]$intGreen
        $rangeColor1.Font.Color =         [Long]$intDarkGreen
        $rangeColor1.Borders.LineStyle =  -4119
        $rangeColor1.Borders.ColorIndex = 16
        $rangeColor2.Interior.Color =     [Long]$intRed
        $rangeColor2.Font.Color =         [Long]$intDarkRed
        $rangeColor2.Borders.LineStyle =  -4119
        $rangeColor2.Borders.ColorIndex = 16
    }



    $teamName
}

exit

























$used = $ws4.UsedRange()
$lastRow = $ws4.Range("Number").SpecialCells(11).row
$lastCol = $ws4.Range("Number").SpecialCells(11).column



for ($col = 1; $col -lt $lastCol; $col+=2)
{
    $rangeColor1 = $ws4.Range($used.Cells(1,$col), $used.Cells($lastRow,$col))
    $rangeColor2 = $rangeColor1.Offset(0,1)

    $rangeColor1.Interior.Color =     [Long]$intRed
    $rangeColor1.Font.Color =         [Long]$intDarkRed
    $rangeColor1.Borders.LineStyle =  -4119
    $rangeColor1.Borders.ColorIndex = 16
    $rangeColor2.Interior.Color =     [Long]$intGreen
    $rangeColor2.Font.Color =         [Long]$intDarkGreen
    $rangeColor2.Borders.LineStyle =  -4119
    $rangeColor2.Borders.ColorIndex = 16
}

exit

function test()
{
    #$playerID   = $wsPlayers.Cells.Item($row, $col + 1)
    $playerName = $wsPlayers.Cells.Item($row, $col + 2)
    $teamName   = $wsPlayers.Cells.Item($row, $col + 3)
    #$teamID     = $wsPlayers.Cells.Item($row, $col + 4)
    #$1SW        = $wsPlayers.Cells.Item($row, $col + 5)
    $1SL        = $wsPlayers.Cells.Item($row, $col + 6).Text
    $2SW        = $wsPlayers.Cells.Item($row, $col + 7).Text
    $2SL        = $wsPlayers.Cells.Item($row, $col + 8).Text
    $1DW        = $wsPlayers.Cells.Item($row, $col + 9).Text
    $1DL        = $wsPlayers.Cells.Item($row, $col + 10).Text
    $2DW        = $wsPlayers.Cells.Item($row, $col + 11).Text
    $2DL        = $wsPlayers.Cells.Item($row, $col + 12).Text
    $3DW        = $wsPlayers.Cells.Item($row, $col + 13).Text
    $3DL        = $wsPlayers.Cells.Item($row, $col + 14).Text
    $Wins       = $wsPlayers.Cells.Item($row, $col + 15).Text
    $Loses      = $wsPlayers.Cells.Item($row, $col + 16).Text

    $playerName.Text
    $teamName.Text
}
