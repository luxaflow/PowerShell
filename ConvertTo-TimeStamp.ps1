<#
<-ConvertTo-TimeStamp->
.SYNOPSIS
Converts DateTime object to Standardized TimeStamp for files paths

.DESCRIPTION
This is the timestamp convention that should be used across scripts that use this module to prevent incosistencies.
Converts to string value form datetime object in format yyyy-MM-ddTHH mm ss

.PARAMETER DateTimeObject 
A datetime Object will be formated

.EXAMPLE 
ConvertTo-TimeStamp -DateTimeObject (09-01-2018)
// 2018-09-01T00 00 00
#>

function ConvertTo-TimeStamp {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            Position = 1)][Datetime]$DateTimeObject
    )
    try {
        return ([datetime]$DateTimeObject).ToString('yyyy-MM-ddTHH mm ss')
    }
    catch {
        throw $_
    }
}