<#
<-Find-WithWildCard->
.SYNOPSIS
Searches an array for a sting value that is corresponding to the given value with the required *  

.DESCRIPTION
Will return all object that are -like the given value with the given wildcard. The function will not edit the original array
Note: the search is caseinsensitive

.PARAMETER Value 
Value that needs to be found in the array

.PARAMETER SearchArray
Array object that needs to be found in

.PARAMETER SearchColumn
If a ther column that needs to be searched is know this can be given and the fuction will only search the given column/key

.EXAMPLE
Find-WithWildCard -Value "*e" -SearchArray @(one,Two,Three,Four,fives)
// @(one,Three)
Find-WithWildCard -Value "*e" -SearchArray @(one,Two,Three,Four,fives)
// @()
Find-WithWildCard -Value "*e*" -SearchArray @(one,Two,Three,Four,fives)
// @(one,Three,fives)
#>
function Find-WithWildCard {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true)][String]$Value,
        [Parameter(
            Mandatory = $true)][Array]$SearchArray,
        [Parameter(
            Mandatory = $false)][String]$SearchColumn
    )
    begin {
        if (($Value[0] -ne '*' -or $Value[-1] -ne '*') -and ($Value[0] -ne '*' -and $value[-1] -ne '*')) {
            Write-Error '* is a required value, and is only allowed at end, start or both start and end'
            exit
        }
    }
    process {
        [System.Collections.ArrayList]$records = @()
        if ($SearchColumn) {
            foreach ($Item in $SearchArray) {
                if ($Item.($SearchColumn) -ilike $Value) {
                    $records.Add($item) | Out-Null
                }
            }
        } 
        else {
            foreach ($Item in $SearchArray) {
                foreach ($PropertyValue in $Item.PSObject.Properties.value) {
                    if ($PropertyValue -like $Value) {
                        $records.Add($Item) | out-null
                        break
                    }
                }
            }
        }
    }
    end {
        return $records
    }
}