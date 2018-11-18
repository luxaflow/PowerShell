<#
<-ConvertTo-ExcelSheet->
.SYNOPSIS
Converts a array of powershell keypair values (objects) in a single excel sheet

.DESCRIPTION
Takes the the field of a excel sheet an d puts each row into a powershell object

.PARAMETER Path 
Path to file or the directory it should be stored in. If no file name is given it will be stored as book1.xlsx

.PARAMETER Objects 
Array of Object that should converted into a single excel sheet.

.PARAMETER SheetNumber 
The sheet the data should be written to. default is 1

.PARAMETER SheetName 
A sheet name can be used to give a sheet a different name

.PARAMETER NoHeaders 
Switch incase the sheet data requires no headers. If else the first row will be made into headers.
#>
function ConvertTo-ExcelSheet {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true)][String]$Path,
        [Parameter(
            Mandatory = $true)][Array]$Objects,
        [Parameter(
            Mandatory = $false)][Int]$SheetNumber,
        [Parameter(
            Mandatory = $false)][Int]$SheetName,
        [Parameter(
            Mandatory = $false)][Switch]$NoHeaders
    )
    begin {
        try {
            if ([System.IO.Path]::GetFileName($Path) -ne "") {
                if ([System.IO.Path]::GetExtension($Path) -ne ".xlsx") {
                    $File = "$Path.xlsx"
                }
                elseif ([System.IO.Path]::GetExtension($Path) -ne ".xls") {
                    $File = "$Path.xlsx"
                }
                else {
                    $File = $Path
                }
            }
            else {
                $File = "$Path\book1.xlsx"
            }
            if (Test-Path -Path $File) {
                $exists = $true
            }
        }
        catch {
            throw $_
        }
        if (!$SheetNumber) {
            $SheetNumber = 1
        }
    }
    process {
        $EX = New-Object -Com Excel.Application
        $EX.Visible = $false
        $EX.DisplayAlerts = $false
        if ($exists) {
            $WB = $EX.Workbooks.Open($File)
        }
        else {
            $WB = $EX.Workbooks.add()
        }
        $WS = $WB.Worksheets.item($SheetNumber)

        if ($SheetName) {
            $WS.name = $SheetName
        }
        if (!$NoHeaders) {
            $headers = ($objects[$objects.count / 2].psobject.properties | Select-Object name)[0..(($objects | Get-Member -MemberType "Property").count - 1)]
            for ($i = 1; $i -le $headers.count; $i++) {
                $WS.cells.item(1, $i) = [String]($headers[$i - 1].Name)
            }
            $r = 2
        }
        else {
            $r = 1
        }
        for ($i = 0; $i -lt $Objects.count; $i++) {
            for ($c = 0; $c -lt $headers.count; $c++) {
                $WS.Cells.Item($r, $c + 1) = $objects[$i][$c]
            }
            $r++
        }
        $WB.SaveAs($Path)
        $WB.Close()
        $EX.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($EX) | Out-Null
    }
    end {
        Return $Output
    }
}