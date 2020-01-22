<#
-ConvertFrom-ExcelSheet-
.SYNOPSIS
Converts a single Excel sheet into powershell custom objects

.DESCRIPTION
Takes the the field of a excel sheet an d puts each row into a powershell object

.PARAMETER FilePath 
FullPath to the file that needs to be converted

.PARAMETER SheetNumber 
the sheet number that needs to be converted incase multiple options are availalble. first sheet is indexed as 1. default is 1

.PARAMETER NoHeaders 
Switch incase the sheet data has no headers. If else the first row will be made into headers.

.NOTES
Do not use this function on large sheets as it searches every field in excel seperatly.
Creating a function that saves the data to csv and imports that will be a lot faster
#>
function ConvertFrom-ExcelSheet {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true)][String]$FilePath,
        [Parameter(
            Mandatory = $false)][Int]$SheetNumber,
        [Parameter(
            Mandatory = $false)][Switch]$NoHeaders
    )
    try {
        $File = Get-Item -Path $FilePath
        $run = $true
    }
    catch {
        throw $_
        $run = $false
    }
    if (-not $SheetNumber) {
        $SheetNumber = 1
    }
    if ($run) {
        $EX = New-Object -Com Excel.Application
        $EX.Visible = $false
        $EX.DisplayAlerts = $false
        $WB = $EX.Workbooks.Open($File)
        $WS = $WB.Worksheets($SheetNumber)
        [Int]$ColumnsCount = $WS.UsedRange.Columns.Count
        [Int]$RowsCount = $WS.UsedRange.Rows.Count
        if ($NoHeaders) {
            $r = 0
        }
        else {
            $r = 1
        }
        [System.Collections.ArrayList]$Output = @();
        for ($r = $r + 1; $r -le $RowsCount; $r++) {
            $EntryObj = New-Object PSCustomObject
            for ($c = 1; $c -le $ColumnsCount; $c++) {
                [String]$Header = $WS.Cells.Item(1, $c).text
                $Value = $WS.Cells.Item($r, $c).text
                if (-not $Header) {$Header = Column$c}
                if (-not $Value) {$Value = $null}
                $EntryObj | Add-Member -NotePropertyName $Header -NotePropertyValue $Value
            }
            $Output.Add($EntryObj) | Out-Null
        }
        $WB.Close()
        $EX.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($EX) | Out-Null
    }
    Return $Output
}