<#
<-New-DummyArray->
.SYNOPSIS
Creates a dummy array for testing purposes

.DESCRIPTION
If for testing purposes a array is required this can be used to create a dummy array

.PARAMETER AmountOfObjects 
Defines mount of created objects. default is 10

.PARAMETER ItemsPerObject
Amount of keys that are need in the objects. default is 10

.PARAMETER CharPerItem
Characters per value.default is 10

.EXAMPLE
New-DummyArray -AmountOfObjects 15 -ItemsPerObject 20 -CharPerItem 20
// returns array with 15 object each with 20 keypair values and 20 characters per value

.NOTES
This function uses the New-Random String function, if you do not want to import tmhe fuction make sure you set the $Value variable in script
#>
function New-DummyArray {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true)][Int]$AmountOfObjects,
        [Parameter(
            Mandatory = $true)][Int]$ItemsPerObject,
        [Parameter(
            Mandatory = $false)][Int]$CharPerItem
    )
    if (!$AmountOfObjects) {
        $AmountOfObjects = 10
    }
    if (!$ItemsPerObject) {
        $ItemsPerObject = 10
    }
    [System.Collections.ArrayList]$Objects = @()
    for ($i = 0; $i -lt $AmountOfObjects; $i++) {
        $Object = [PSCustomObject]@{}
        for ($n = 0; $n -lt $ItemsPerObject; $n++) {
            if ($CharPerItem) {
                $Value = New-RandomString -Length $CharPerItem #set something else if you do not want to import this function
            } 
            else {
                $Value = New-RandomString #set something else if you do not want to import this function
            }
            
            $Object | Add-Member -MemberType NoteProperty -Name ('Field' + ($n + 1)) -Value $Value
        }
        $Objects.Add($Object) | Out-Null
        Remove-Variable 'Object'
    }
    return $Objects
}
