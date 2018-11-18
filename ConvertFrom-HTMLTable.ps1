<#
<-ConvertFrom-HTMLTable->
.SYNOPSIS
Converts html table to powershell objects from tag <table></table> 

.DESCRIPTION
Convert a html table to powershell objects. this automatically determines if there is a header available.

.PARAMETER HTMLContent 
Add html content in more that 1 table is added in 1 go then the script will only grab the first table.

.NOTES
The script assumes that the table has a vartical layout
#>
Function ConvertFrom-HTMLTable {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true)][String]$HTMLContent
    )
    #filter the first table out of the html in case there is more markup
    $table = -join $HTMLContent[$HTMLContent.indexof('<table>')..($HTMLContent.indexof('</table>') + ('</table>'.length - 1))]
    $rows = @()
    while ($table -match '<tr>' -and $table -match '</tr>') {
        $StartRow = $table.IndexOf('<tr>')
        $EndRow = $table.IndexOf('</tr>') + 4
        $rows += (-join $table[$StartRow..$EndRow])
        $table = $table.Remove($StartRow, $EndRow - $StartRow + 1)
    }
    $headers = @()
    [System.Collections.ArrayList]$itemValues = @()
    foreach ($row in $rows) {
        if ($row -match '<th>' -and $row -match '</th>') {
            while ($row -match '<th>' -and $row -match '</th>') {
                $startTag = $row.IndexOf('<th>')
                $endTag = $row.IndexOf('</th>') + 4
                $headers += (-join $row[$startTag..$endTag]).Replace('<th>', '').Replace('</th>', '')
                $row = $row.remove($startTag, $endTag - $startTag + 1)
            }
        }
        elseif ($row -match '<td>' -and $row -match '</td>') {
            $values = @()
            while ($row -match '<td>' -and $row -match '</td>') {
                $startTag = $row.IndexOf('<td>')
                $endTag = $row.IndexOf('</td>') + 4
                $values += (-join $row[$startTag..$endTag]).Replace('<td>', '').Replace('</td>', '')
                $row = $row.remove($startTag, $endTag - $startTag + 1)
            }
            $itemValues.add($values) | Out-Null
        }
    }
    #Check if any of the rows with values has more values than the amount of headers. if so add fields to correspond
    foreach ($obj in $itemValues) {
        if ($obj.count -gt $headers.count) {
            $diff = $obj.count - $headers.count
            for ($i = 1; $i -le $diff; $i++) {
                $headers += ('Field' + ($headers.count + $i)) 
            }
        }
    }
    [System.Collections.ArrayList]$objects = @()
    for ($i = 0; $i -lt $itemValues.count; $i++) {
        $object = [PSCustomObject]@{}
        for ($v = 0; $v -lt $itemValues[$i].count; $v++) {
            $object | Add-Member -MemberType NoteProperty -Name $headers[$v] -Value $itemValues[$i][$v]
        }
        $objects.Add($object) | Out-Null
    }
    return $objects
}