<#
<-ConvertTo-HTMLTable->
.SYNOPSIS
Converts array of powershell objects into a HTML table format

.DESCRIPTION
The headers for the table can be added to the top or left side of the table this can be distignuished by the <th> tag. 
Note: No styling is added using this function only the data inbetween the <table></table> tags is created

.PARAMETER InputArray 
The array of objects that needs to be converted

.PARAMETER ColumnsOrder 
the order the given columns should be adhered to

.PARAMETER NoHeaders
a switch param to remove all headers form the table

.PARAMETER Horizontal
this Switch will move the headers to the left side of the table

.PARAMETER CenterData
This switch will center all information between <td></td> tags
#>
function ConvertTo-HTMLTable {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true)][Array]$InputArray,
        [Parameter(
            Mandatory = $true)][Array]$ColumnsOrder,
        [Parameter(
            Mandatory = $false)][Array]$ErrorValues,
        [Parameter(
            Mandatory = $false)][Switch]$NoHeaders,
        [Parameter(
            Mandatory = $false)][Switch]$Horizontal,
        [Parameter(
            Mandatory = $false)][Switch]$CenterData
    )
    $Body = '<table>'
    if ($CenterData) {
        $inlineCenter = 'style="text-align:center;"'
    }
    if (!$NoHeaders) {
        if (!$Horizontal) {
            $body += '<tr>'
            foreach ($Column in $ColumnsOrder) {
                $Body += '<th>' + $Column + '</th>'
            }
            $body += '</tr>'
            foreach ($Object in $InputArray) {
                if ($ErrorValues) {
                    for ($i = 1; $i -le $ErrorValues.count; $i++) {
                        if ($Object -imatch $ErrorValues[$i - 1]) {
                            $rowStart = '<tr class="error">'
                            break
                        }
                        elseif ($i -eq $errorValues.count) {
                            $rowStart = '<tr>' 
                        }
                    }
                }
                else {
                    $rowStart = '<tr>'
                }
                $body += $rowStart
                foreach ($Column in $ColumnsOrder) {
                    $Body += "<td $inlineCenter>" + $Object.($Column) + '</td>'
                }
                $body += '</tr>'
            }
        }
        elseif ($Horizontal) {
            foreach ($Column in $ColumnsOrder) {
                $body += '<tr><th>' + $Column + '</th>'
                foreach ($Object in $InputArray) {
                    $body += "<td $inlineCenter>" + $Object.($Column) + '</td>'
                }
                $body += '</tr>'
            }
        }
    }
    elseif ($NoHeaders) {
        if (!$Horizontal) {
            foreach ($Object in $InputArray) {
                if ($ErrorValues) {
                    for ($i = 1; $i -le $ErrorValues.count; $i++) {
                        if ($Object -imatch $ErrorValues[$i - 1]) {
                            $rowStart = '<tr class="error">'
                            break
                        }
                        elseif ($i -eq $errorValues.count) {
                            $rowStart = '<tr>' 
                        }
                    }
                }
                else {
                    $rowStart = '<tr>'
                }
                $body += $rowStart
                foreach ($Column in $ColumnsOrder) {
                    $Body += "<td $inlineCenter>" + $Object.($Column) + '</td>'
                }
                $body += '</tr>'
            }
        }
        elseif ($Horizontal) {
            foreach ($Column in $ColumnsOrder) {
                $body += '<tr>'
                foreach ($Object in $InputArray) {
                    $body += "<td $inlineCenter>" + $Object.($Column) + '</td>'
                }
                $body += '</tr>'
            }
        }
    }
    return $body + '</table>'
}