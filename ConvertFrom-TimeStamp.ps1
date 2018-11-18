<#
<-ConvertFrom-TimeStamp->
.SYNOPSIS
Converts Standardized TimeStamp String to DateTime Object 

.DESCRIPTION
This is the timestamp convention that should be used across scripts that uses this module to prevent incosistencies.
Converts to string value form datetime object in format yyyy-MM-ddTHH mm ss

.PARAMETER TimeStampObject 
A String Object in from the default timestamp

.EXAMPLE 
ConvertFrom-TimeStamp -TimeStampObject "2018-09-01T00 00 00"
// DateTime Object

.NOTES
Should at some point add a option to converti it directly of the file name
#>
function ConvertFrom-TimeStamp {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            Position = 1)]$TimeStampObject
    )
    return [datetime]($TimeStampObject.replace(' ', ':'))
}