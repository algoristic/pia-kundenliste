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
    [String]$ImportArchive = "$PSScriptRoot\import"
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

# create a backup file
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

Function Create-CustomerListBackup
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [String]$FileToBackup
    )
    Create-Backup -FileToBackup $FileToBackup -FileBackupDirectory $BackupDir -BackupFileName "Kundenliste"
}

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
    
}

# backup working file
Create-CustomerListBackup $WorkingFile

Try
{
    # open up the workbook
    $Excel = new-object -comobject Excel.Application
    $Excel.visible = $false

    $MergeFileData = $Excel.Workbooks.Open($MergeFile)
    $WorkingFileData = $Excel.Workbooks.Open($WorkingFile)

    $Import = $MergeFileData.Worksheets.Item($ImportFile.Data.Index)
    $Overview = $WorkingFileData.Worksheets.Item($Base.Overview.Index)
    $Masterdata = $WorkingFileData.Worksheets.Item($Base.Masterdata.Index)

    Transfer-CustomerData -Import $Import -Overview $Overview -Masterdata $Masterdata

    # $WorkingFileData.SaveAs($WorkingFile)
    $Excel.Quit()
    Create-ImportFileBackup $MergeFile
}
Catch
{
    # delete backup file if no changes take place due to error while processing
    Remove-Item $WorkingFileBackup
}
echo "Fertig! Sie koennen das Fenster jetzt schliessen."
