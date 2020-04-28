<#
    AUTHOR: MARCO LEWEKE
#>
[CmdletBinding()]
Param(
    [Parameter(Mandatory=$true)]
    [String]$File
)
# open up the workbook
$excel = new-object -comobject Excel.Application
$excel.visible = $false
$workbook = $excel.workbooks.open($File)

$excel.visible = $true
