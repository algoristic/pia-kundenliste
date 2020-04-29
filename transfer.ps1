<#
    AUTHOR: MARCO LEWEKE
#>
[CmdletBinding()]
Param(
    [Parameter(Mandatory=$true)]
    [String]$MergeFile,

    [Parameter(Mandatory=$false)]
    [String]$WorkingFile = "$PSScriptRoot\Kundenliste.xlsx",

    [Parameter(Mandatory=$false)]
    [String]$BackupDir = "$PSScriptRoot\archiv",

    [Parameter(Mandatory=$false)]
    [String]$ImportArchive = "$PSScriptRoot\import",

    [Parameter(Mandatory=$false)]
    [String]$DateFormat = "MM.dd.yyyy",

    [Parameter(Mandatory=$false)]
    [String]$DateFormatLocal = "TT.MM.JJJJ"
)
# init structures
$ImportFile = @{
    Data = @{
        Headline = 1
        Index = 1
        Layer = 1
        CustomerNo = 2
        Status = 3
        Surname = 4
        Forename = 5
        Company = 6
        PostalCode = 7
        Place = 8
        Phone = 9
        EMail = 10
        BelongsTo = 11
        FirstOrderDate = 12
        LastSale = 13

    }
}
$Base = @{
    Overview = @{
        Headline = 1
        Index = 1
        Layer = 1
        CustomerNo = 2
        Status = 3
        Surname = 4
        Forename = 5
        Phone = 6
        FirstContactDate = 7
        FirstOrderDate = 8
        LastSale = 9
        NextOrder = 10
    }
    Masterdata = @{
        Headline = 1
        Index = 2
        CustomerNo = 1
        Surname = 2
        Forename = 3
        DogCanSize = 4
        DogAmount = 5
        DogRation = 6
        DogReserve = 7
        CatCanSize = 8
        CatAmount = 9
        CatRation = 10
        CatReserve = 11
        Note = 12
    }
}

# make backup filename globally visible
# so we can delete the file if not needed
$WorkingFileBackup = ""

Function Parse-ImportData
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        $Worksheet,
        [Parameter(Mandatory=$true)]
        $Line
    )
    $ImportLine = @{
        Layer = $Worksheet.Cells.Item($Line, $ImportFile.Data.Layer).Text
        CustomerNo = $Worksheet.Cells.Item($Line, $ImportFile.Data.CustomerNo).Value2
        Status = $Worksheet.Cells.Item($Line, $ImportFile.Data.Status).Text
        Surname = $Worksheet.Cells.Item($Line, $ImportFile.Data.Surname).Text
        Forename = $Worksheet.Cells.Item($Line, $ImportFile.Data.Forename).Text
        Company = $Worksheet.Cells.Item($Line, $ImportFile.Data.Company).Text
        PostalCode = $Worksheet.Cells.Item($Line, $ImportFile.Data.PostalCode).Text
        Place = $Worksheet.Cells.Item($Line, $ImportFile.Data.Place).Text
        Phone = $Worksheet.Cells.Item($Line, $ImportFile.Data.Phone).Text
        EMail = $Worksheet.Cells.Item($Line, $ImportFile.Data.EMail).Text
        BelongsTo = $Worksheet.Cells.Item($Line, $ImportFile.Data.BelongsTo).Value2
        FirstOrderDate = $null
        LastSale = $null
    }
    $FirstOrderDateValue = $Worksheet.Cells.Item($Line, $ImportFile.Data.FirstOrderDate).Text
    If($FirstOrderDateValue -ne "")
    {
        $FirstOrderDate = [datetime]::ParseExact($FirstOrderDateValue, 'yyyy-MM-dd', $null)
        $ImportLine.FirstOrderDate = $FirstOrderDate
    }

    $LastSaleValue = $Worksheet.Cells.Item($Line, $ImportFile.Data.LastSale).Text
    If($LastSaleValue -ne "")
    {
        $LastSale = [datetime]::ParseExact($LastSaleValue, 'yyyy-MM-dd', $null)
        $ImportLine.LastSale = $LastSale
    }
    return $ImportLine
}

Function Find-Line
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        $CustomerNo,
        [Parameter(Mandatory=$true)]
        $Index,
        [Parameter(Mandatory=$true)]
        $Worksheet,
        [Parameter(Mandatory=$false)]
        [int]$StartLineIndex = 2
    )
    $Line = --$StartLineIndex
    Do {
        $Line++
        $CurrentCustomerNo = $Worksheet.Cells.Item($Line, $Index)
        If ($CurrentCustomerNo.Text -ne "")
        {
            If ($CurrentCustomerNo.Value2 -eq $CustomerNo)
            {
                return $Line
            }
        }
    } While ($CurrentCustomerNo.Text -ne "")
    # we iterated the whole dataset, so index is equals (last line + 1)
    return $Line
}

Function Set-DateData
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        $Cell,
        [Parameter(Mandatory=$true)]
        $Date
    )
    $Day = $Date.Day
    If ($Day -lt 10)
    {
        $Day = "0$Day"
    }
    $Month = $Date.Month
    If ($Month -lt 10)
    {
        $Month = "0$Month"
    }
    $Year = $Date.Year
    $Cell.Value2 = "$Day.$Month.$Year"
    $Cell.NumberFormat = $DateFormat
    $Cell.NumberFormatLocal = $DateFormatLocal
}

Function Get-NextOrderDate
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        $Line
    )
    return "=I2+(WENNNV(WENN(MIN(WENN(SVERWEIS(B2;Stammdaten!A:K;7)>0;SVERWEIS(B2;Stammdaten!A:K;7);9999);WENN(SVERWEIS(B2;Stammdaten!A:K;11)>0;SVERWEIS(B2;Stammdaten!A:K;11);9999))=9999;30;MIN(WENN(SVERWEIS(B2;Stammdaten!A:K;7)>0;SVERWEIS(B2;Stammdaten!A:K;7);9999);WENN(SVERWEIS(B2;Stammdaten!A:K;11)>0;SVERWEIS(B2;Stammdaten!A:K;11);9999)));30))"
}

Function Write-MasterdataLine
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        $Masterdata,
        [Parameter(Mandatory=$true)]
        $Line,
        [Parameter(Mandatory=$true)]
        $Data
    )
    $CustomerNo = $Masterdata.Cells.Item($Line, $Base.Masterdata.CustomerNo)
    If ($CustomerNo.Text -eq "")
    {
        $CustomerNo.Value2 = $Data.CustomerNo.ToString()
    }
    $Surname = $Masterdata.Cells.Item($Line, $Base.Masterdata.Surname)
    If ($Surname.Text -eq "")
    {
        $Surname.Value2 = $Data.Surname.ToString()
    }
    $Forename = $Masterdata.Cells.Item($Line, $Base.Masterdata.Forename)
    If ($Forename.Text -eq "")
    {
        $Forename.Value2 = $Data.Forename.ToString()
    }

    # available fields:
    # masterdata: CustomerNo, Surname, Forename
    # import:     Layer, CustomerNo, Status, Surname, Forename, Company, PostalCode, Place, Phone, EMail, BelongsTo, FirstOrderDate, LastSale
}

Function Write-OverviewLine
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        $Overview,
        [Parameter(Mandatory=$true)]
        $Line,
        [Parameter(Mandatory=$true)]
        $Data
    )
    $Layer = $Overview.Cells.Item($Line, $Base.Overview.Layer)
    $Layer.Value2 = $Data.Layer.ToString()

    $CustomerNo = $Overview.Cells.Item($Line, $Base.Overview.CustomerNo)
    If ($CustomerNo.Text -eq "")
    {
        $CustomerNo.Value2 = $Data.CustomerNo.ToString()
    }

    # customer can become a teampartner etc. -> so overwrite this
    $Status = $Overview.Cells.Item($Line, $Base.Overview.Status)
    $Status.Value2 = $Data.Status.ToString()

    $Surname = $Overview.Cells.Item($Line, $Base.Overview.Surname)
    If ($Surname.Text -eq "")
    {
        $Surname.Value2 = $Data.Surname.ToString()
    }
    $Forename = $Overview.Cells.Item($Line, $Base.Overview.Forename)
    If ($Forename.Text -eq "")
    {
        $Forename.Value2 = $Data.Forename.ToString()
    }

    $Phone = $Overview.Cells.Item($Line, $Base.Overview.Phone)
    $Phone.Value2 = $Data.Phone.ToString()

    If ($Data.LastSale)
    {
        $LastSale = $Overview.Cells.Item($Line, $Base.Overview.LastSale)
        Set-DateData -Cell $LastSale -Date $Data.LastSale
    }
    Else
    {
        If ($Data.FirstOrderDate)
        {
            $LastSale = $Overview.Cells.Item($Line, $Base.Overview.LastSale)
            Set-DateData -Cell $LastSale -Date $Data.FirstOrderDate
        }
    }

    If ($Data.FirstOrderDate)
    {
        $FirstOrderDate = $Overview.Cells.Item($Line, $Base.Overview.FirstOrderDate)
        Set-DateData -Cell $FirstOrderDate -Date $Data.FirstOrderDate
    }

    # $NextOrder = $Overview.Cells.Item($Line, $Base.Overview.NextOrder)
    # If ($NextOrder.Text -eq "")
    # {
    #     $NextOrder.formula = (Get-NextOrderDate -Line $Line)
    # }

    # available fields:
    # overview: Layer, CustomerNo, Status, Surname, Forename, Phone, FirstContactDate, FirstOrderDate, LastSale, NextOrder
    # import: Layer, CustomerNo, Status, Surname, Forename, Company, PostalCode, Place, Phone, EMail, BelongsTo, FirstOrderDate, LastSale
}

Function Transfer-CustomerData
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        $Import,
        [Parameter(Mandatory=$true)]
        $Overview,
        [Parameter(Mandatory=$true)]
        $Masterdata
    )
    $Layer = $null
    $Line = $ImportFile.Data.Headline + 1
    Do {
        $Layer = $Import.Cells.Item($Line, $ImportFile.Data.Layer)
        If ($Layer.Text -ne "")
        {
            $ImportLine = Parse-ImportData -Worksheet $Import -Line $Line
            $OverviewLineIndex = Find-Line -CustomerNo $ImportLine.CustomerNo -Index $Base.Overview.CustomerNo -Worksheet $Overview
            Write-OverviewLine -Overview $Overview -Line $OverviewLineIndex -Data $ImportLine
            $MasterdataLineIndex = Find-Line -CustomerNo $ImportLine.CustomerNo -Index $Base.Masterdata.CustomerNo -Worksheet $Masterdata
            Write-MasterdataLine -Masterdata $Masterdata -Line $MasterdataLineIndex -Data $ImportLine
        }
        $Line++
    } While ($Layer.Text -ne "")
}

# create a generic backup file
Function Create-Backup
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [String]$FileToBackup,
        [Parameter(Mandatory=$true)]
        [String]$FileBackupDirectory,
        [Parameter(Mandatory=$true)]
        [String]$BackupFileName,
        [Parameter(Mandatory=$false)]
        [String]$BackupFileFormat = "xlsx"
    )
    $Timestamp = (Get-Date -UFormat "%Y-%m-%d_%H-%M-%S").tostring()
    $BackupFileName = "$BackupFileName-$Timestamp"
    $WorkingFileBackup = "$FileBackupDirectory\$BackupFileName.$BackupFileFormat"
    Copy-Item $FileToBackup -Destination $WorkingFileBackup
}

# create a backup file for the customer list
Function Create-CustomerListBackup
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [String]$FileToBackup
    )
    Create-Backup -FileToBackup $FileToBackup -FileBackupDirectory $BackupDir -BackupFileName "Kundenliste"
}

# create a backup for the processed import and delete it afterwards
Function Create-ImportFileBackup
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [String]$FileToBackup
    )
    Create-Backup -FileToBackup $FileToBackup -FileBackupDirectory $ImportArchive -BackupFileName "Import" -BackupFileFormat "xls"
    # TODO: delete import-file after backup
    # Remove-Item $FileToBackup
}

# backup working file
Create-CustomerListBackup $WorkingFile

# open up the workbook
$Excel = new-object -comobject Excel.Application
$Excel.visible = $false

$MergeFileData = $Excel.Workbooks.Open($MergeFile)
$WorkingFileData = $Excel.Workbooks.Open($WorkingFile)

$Import = $MergeFileData.Worksheets.Item($ImportFile.Data.Index)
$Overview = $WorkingFileData.Worksheets.Item($Base.Overview.Index)
$Masterdata = $WorkingFileData.Worksheets.Item($Base.Masterdata.Index)

Transfer-CustomerData -Import $Import -Overview $Overview -Masterdata $Masterdata

$WorkingFileData.SaveAs($WorkingFile)
$Excel.Quit()
Create-ImportFileBackup $MergeFile

echo "Fertig! Sie koennen das Fenster jetzt schliessen."
