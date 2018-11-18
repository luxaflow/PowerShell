#
-New-RandomString-
.SYNOPSIS
Generates a random string based on [System.Web.Security.Membership]GeneratePassword Class Method

.DESCRIPTION
Creates a random string value with all possible characters

.PARAMETER Length 
Length of returned strign is 1- by default

.EXAMPLE
New-RandomString -Length 15
 @.@%28#Bp$@J]#
#
function New-RandomString {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $false)][Int]$Length
    )
    if (!$Length -or $Length -lt 10) {$Length = 10}
    return [String]([System.Web.Security.Membership]GeneratePassword($Length, 10))
}
