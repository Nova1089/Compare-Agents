<#
Objective:
Compare device exports from Azure, FreshService, Absolute and Sophos. Find out where devices are not showing up, and view all the data in one spreadsheet.

Procedure:
1. Export device reports from any of Azure, FreshService, Absolute, and Sophos. (Choose 2 or more.)
    Note: Be sure to export in CSV format.
    Note: When exporting from FreshService, it's useful (but not necessary) to filter by Laptop and select the following columns:
        Display Name
        Serial Number
        Used By
        Last Login By
        Acquisition Date
        Warranty Expiry Date
        Asset State
        Last Audit Date
2. Place each export into their own folder together.
3. Run the script and follow prompts.
#>

function Prompt-FolderPath
{    
    return Read-Host "Enter folder path. (i.e. C:\Users\Username\Desktop\My Folder)"
}

function Prompt-BaseFile
{
    do
    {
        $prompt = Read-Host "Which file would you like to compare against? 
        For Azure type: az
        For FreshService type: fs
        For Absolute type: ab
        For Sophos type: so `n"

        if ($prompt -inotmatch "\baz\b|\bfs\b|\bso\b|\bab\b")
        {
            Write-Warning "Invalid response. Please try again."
        }
    }
    while ($prompt -inotmatch "\baz\b|\bfs\b|\bso\b|\bab\b")

    switch -Regex ($prompt)
    {
        "\b[aA][zZ]\b" { return "Azure" }
        "\b[fF][sS]\b" { return "FreshService" }
        "\b[aA][bB]\b" { return "Absolute" }
        "\b[sS][oO]\b" { return "Sophos" }
    }
}

function Validate-Inputs($files, $maxSize, $maxFileCount, $minFileCount)
{    
    $isSizeNormal = Test-SizeNormal -files $files -maxSize $maxSize
    if (!($isSizeNormal)) { return $False }
    $isFileCountValid = Test-FileCount -files $files -maxFileCount $maxFileCount -minFileCount $minFileCount
    if (!($isFileCountValid)) { return $false }
    return $true
}

function Test-SizeNormal($files, $maxSize)
{  
    $sizeTotal = Get-SizeMB($files)

    if ($sizeTotal -gt $maxSize)
    {
        Write-Warning "File size total is abormally large (greater than 100MB)."
        Write-Warning "Size is: $sizeTotal"
        Read-Host "Press Enter to continue or CTRL+C to cancel."
    }
    if ($sizeTotal -eq 0)
    {
        Write-Warning "Path invalid or folder is empty. Please try again."
        return $False
    }
    return $True
}

function Test-FileCount([object[]]$files, $maxFileCount, $minFileCount)
{
    if ($files.Length -gt $maxFileCount)
    {
        Write-Warning "There are more than $maxFileCount CSV files in this folder. Please try again."
        return $False
    }
    if ($files.Length -lt $minFileCount)
    {
        Write-Warning "There are less than $minFileCount CSV files in this folder. Please try again."
        return $False
    }
    return $True
}

function Get-SizeMB($files)
{
    $sizeKB = $files | Measure-Object -Sum Length | Select-Object -ExpandProperty Sum
    $sizeMB = $sizeKB / 1MB
    return $sizeMB
}

function Import-CSVList($csvFiles)
{
    Write-Host "Importing CSV files..."

    $importedCSVList = New-Object -TypeName object[] $csvFiles.Length

    for ($i = 0; $i -lt $importedCSVList.Length; $i++)
    {
        $importedCSVList[$i] = $csvFiles[$i].PSPath | Import-Csv
    }

    return $importedCSVList
}

function Get-AzureFileIndex($importedCSVList)
{
    $index = Get-FileIndex -importedCSVList $importedCSVList -uniqueHeaderValue "joinType (trustType)"
    if ($null -eq $index)
    {
        Write-Warning "Didn't find Azure file."
        return $null
    }
    Write-Host "Found Azure file."
    return $index
}

function Get-FreshServiceFileIndex($importedCSVList)
{
    $index = Get-FileIndex -importedCSVList $importedCSVList -uniqueHeaderValue "Used By"
    if ($null -eq $index)
    {
        Write-Warning "Didn't find FreshService file."
        return $null
    }
    Write-Host "Found FreshService file."
    return $index
}

function Get-SophosFileIndex($importedCSVList)
{
    $index = Get-FileIndex -importedCSVList $importedCSVList -uniqueHeaderValue "Health Status"
    if ($null -eq $index)
    {
        Write-Warning "Didn't find Sophos file."
        return $null
    }
    Write-Host "Found Sophos file."
    return $index
}

function Get-AbsoluteFileIndex($importedCSVList)
{
    $index = Get-FileIndex -importedCSVList $importedCSVList -uniqueHeaderValue "Encryption status"
    if ($null -eq $index)
    {
        Write-Warning "Didn't find Absolute file."
        return $null
    }
    Write-Host "Found Absolute file."
    return $index
}

function Get-FileIndex($importedCSVList, $uniqueHeaderValue)
{
    for ($i = 0; $i -lt $importedCSVList.Length; $i++)
    { 
        $row1 = $importedCSVList[$i] | Select-Object -First 1   
        if ([bool]$row1.PSObject.Properties[$uniqueHeaderValue]) # if statement checks if object has specified property 
        {
            return $i
        }
    }
    return $null
}

function Export-Report($report)
{
    if ($null -eq $report) { return }
    Write-Host "Exporting report..."
    $path = New-Path
    $report | Export-CSV $path -NoTypeInformation
    Write-Host "Finished exporting to $path."
}

function New-Path
{
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $timeStamp = New-TimeStamp
    return "$desktopPath\Agent Comparison $timeStamp.csv"
}

function New-TimeStamp
{
    return (Get-Date -Format yyyy-MM-dd-hh-mm).ToString()
}

class ReportBuilder
{
    # properties
    [object[]]$importedCSVList
    $azureIndex
    $freshServiceIndex
    $absoluteIndex    
    $sophosIndex
    
    # constructors
    ReportBuilder([object[]]$importedCSVList, $azureIndex, $absoluteIndex, $freshServiceIndex, $sophosIndex)
    {
        $this.importedCSVList = $importedCSVList
        $this.azureIndex = $azureIndex
        $this.freshServiceIndex = $freshServiceIndex
        $this.absoluteIndex = $absoluteIndex        
        $this.sophosIndex = $sophosIndex
    }

    # methods
    NewReport()
    {
        throw "Must override method."
    }

    AddDataFromAzure($deviceInfoRow, $baseRowDeviceName)
    {
        if ($null -eq $this.azureIndex) { return }

        $inAzure = $False
        foreach ($azureRow in $this.importedCSVList[$this.azureIndex])
        {
            if ($azureRow.displayName -eq $baseRowDeviceName)
            {
                $deviceInfoRow.AddAzureData($azureRow)
                $inAzure = $True
                break
            }
        }
        if ($inAzure -eq $False)
        {
            $deviceInfoRow.inAzure = $False
        }
    }

    AddDataFromFreshService($deviceInfoRow, $baseRowDeviceName)
    {
        if ($null -eq $this.freshServiceIndex) { return }

        $inFreshService = $False
        foreach ($fsRow in $this.importedCSVList[$this.freshServiceIndex])
        {
            if ($fsRow.'Display Name' -eq $baseRowDeviceName)
            {
                $deviceInfoRow.AddFreshServiceData($fsRow)
                $inFreshService = $True
                break
            }
        }
        if ($inFreshService -eq $False)
        {
            $deviceInfoRow.inFreshService = $False
        }
    }

    AddDataFromAbsolute($deviceInfoRow, $baseRowDeviceName)
    {
        if ($null -eq $this.absoluteIndex) { return }

        $inAbsolute = $False
        foreach ($absRow in $this.importedCSVList[$this.absoluteIndex])
        {
            if ($absRow.'Device name' -eq $baseRowDeviceName)
            {
                $deviceInfoRow.AddAbsoluteData($absRow)
                $inAbsolute = $True
                break
            }        
        }
        if ($inAbsolute -eq $False)
        {
            $deviceInfoRow.inAbsolute = $False
        }
    }

    AddDataFromSophos($deviceInfoRow, $baseRowDeviceName)
    {
        if ($null -eq $this.sophosIndex) { return }

        $inSophos = $False
        foreach ($sophosRow in $this.importedCSVList[$this.sophosIndex])
        {
            if ($sophosRow.Name -eq $baseRowDeviceName)
            {
                $deviceInfoRow.AddSophosData($sophosRow)
                $inSophos = $True
                break
            }
        }
        if ($inSophos -eq $False)
        {
            $deviceInfoRow.inSophos = $False
        }
    }
}

class ReportBuilderAzureBase : ReportBuilder
{
    # constructors
    ReportBuilderAzureBase([object[]]$importedCSVList, $azureIndex, $absoluteIndex, $freshServiceIndex, $sophosIndex) :
    base([object[]]$importedCSVList, $azureIndex, $absoluteIndex, $freshServiceIndex, $sophosIndex)
    {
        # calls the constructor of the base class and then the following
        if ($null -eq $azureIndex) { Throw "Azure CSV was not found." }
    }
    
    # methods
    [object[]] NewReport()
    {
        Write-Host "Building report..."
        $report = New-Object -TypeName object[] $this.importedCSVList[$this.azureIndex].Length
        $i = 0
        foreach ($azureRow in $this.importedCSVList[$this.azureIndex])
        {
            $deviceInfoRow = New-Object -TypeName DeviceInfo
            $baseRowDeviceName = $azureRow.displayName

            # add data from each CSV
            $deviceInfoRow.AddAzureData($azureRow)
            # calling methods of base class
            [ReportBuilder]$this.AddDataFromAbsolute($deviceInfoRow, $baseRowDeviceName)
            [ReportBuilder]$this.AddDataFromFreshService($deviceInfoRow, $baseRowDeviceName)
            [ReportBuilder]$this.AddDataFromSophos($deviceInfoRow, $baseRowDeviceName)

            $report[$i] = $deviceInfoRow
            $i++ 
        }
        return $report
    }
}

class ReportBuilderFreshServiceBase : ReportBuilder
{
    # constructors
    ReportBuilderFreshServiceBase([object[]]$importedCSVList, $azureIndex, $absoluteIndex, $freshServiceIndex, $sophosIndex) :
    base([object[]]$importedCSVList, $azureIndex, $absoluteIndex, $freshServiceIndex, $sophosIndex)
    {
        # calls the constructor of the base class and then the following
        if ($null -eq $freshServiceIndex) { Throw "FreshService CSV was not found." }
    }
    
    # methods
    [object[]] NewReport()
    {
        Write-Host "Building report..."
        $report = New-Object -TypeName object[] $this.importedCSVList[$this.freshServiceIndex].Length
        $i = 0
        foreach ($freshServiceRow in $this.importedCSVList[$this.freshServiceIndex])
        {
            $deviceInfoRow = New-Object -TypeName DeviceInfo
            $baseRowDeviceName = $freshServiceRow.'Display Name'

            # add data from each CSV
            $deviceInfoRow.AddFreshServiceData($freshServiceRow)
            # calling methods of base class           
            [ReportBuilder]$this.AddDataFromAbsolute($deviceInfoRow, $baseRowDeviceName)
            [ReportBuilder]$this.AddDataFromAzure($deviceInfoRow, $baseRowDeviceName)            
            [ReportBuilder]$this.AddDataFromSophos($deviceInfoRow, $baseRowDeviceName)

            $report[$i] = $deviceInfoRow
            $i++ 
        }
        return $report
    }
}

class ReportBuilderAbsoluteBase : ReportBuilder
{
    # constructors
    ReportBuilderAbsoluteBase([object[]]$importedCSVList, $azureIndex, $absoluteIndex, $freshServiceIndex, $sophosIndex) :
    base([object[]]$importedCSVList, $azureIndex, $absoluteIndex, $freshServiceIndex, $sophosIndex)
    {
        # calls the constructor of the base class and then the following
        if ($null -eq $absoluteIndex) { Throw "Absolute CSV was not found." }
    }
    
    # methods
    [object[]] NewReport()
    {
        Write-Host "Building report..."
        $report = New-Object -TypeName object[] $this.importedCSVList[$this.absoluteIndex].Length
        $i = 0
        foreach ($absRow in $this.importedCSVList[$this.absoluteIndex])
        {
            $deviceInfoRow = New-Object -TypeName DeviceInfo
            $baseRowDeviceName = $absRow.'Device name'

            # add data from each CSV
            $deviceInfoRow.AddAbsoluteData($absRow)
            # calling methods of base class
            [ReportBuilder]$this.AddDataFromAzure($deviceInfoRow, $baseRowDeviceName)
            [ReportBuilder]$this.AddDataFromFreshService($deviceInfoRow, $baseRowDeviceName)
            [ReportBuilder]$this.AddDataFromSophos($deviceInfoRow, $baseRowDeviceName)

            $report[$i] = $deviceInfoRow
            $i++ 
        }
        return $report
    }
}

class ReportBuilderSophosBase : ReportBuilder
{
    # constructors
    ReportBuilderSophosBase([object[]]$importedCSVList, $azureIndex, $absoluteIndex, $freshServiceIndex, $sophosIndex) :
    base([object[]]$importedCSVList, $azureIndex, $absoluteIndex, $freshServiceIndex, $sophosIndex)
    {
        # calls the constructor of the base class and then the following
        if ($null -eq $sophosIndex) { Throw "Sophos CSV was not found." }
    }
    
    # methods
    [object[]] NewReport()
    {
        Write-Host "Building report..."
        $report = New-Object -TypeName object[] $this.importedCSVList[$this.sophosIndex].Length
        $i = 0
        foreach ($sophosRow in $this.importedCSVList[$this.sophosIndex])
        {
            $deviceInfoRow = New-Object -TypeName DeviceInfo
            $baseRowDeviceName = $sophosRow.Name

            # add data from each CSV
            $deviceInfoRow.AddSophosData($sophosRow)
            # calling methods of base class
            [ReportBuilder]$this.AddDataFromAbsolute($deviceInfoRow, $baseRowDeviceName)
            [ReportBuilder]$this.AddDataFromFreshService($deviceInfoRow, $baseRowDeviceName)
            [ReportBuilder]$this.AddDataFromAzure($deviceInfoRow, $baseRowDeviceName)

            $report[$i] = $deviceInfoRow
            $i++ 
        }
        return $report
    }
}

class DeviceInfo
{
    # properties
    [string]$inAzure
    [string]$inFreshService
    [string]$inAbsolute
    [string]$inSophos

    # properties from Azure    
    [string]$Azure_Display_Name
    [string]$Azure_OS
    [string]$Azure_Join_Type
    [string]$Azure_Username
    [string]$Azure_Registration_Time
    [string]$Azure_Last_Sign_In_Time

    # properties from FreshService    
    [string]$FreshService_Display_Name
    [string]$FreshService_SN
    [string]$FreshService_Used_By
    [string]$FreshService_Last_Login_By
    [string]$FreshService_Acquisition_Date
    [string]$FreshService_Warranty_Expiry_Date
    [string]$FreshService_Asset_State
    [string]$FreshService_Last_Audit_Date

    # properties from Absolute    
    [string]$Absolute_Device_Name
    [string]$Absolute_SN
    [string]$Absolute_Last_Connected
    [string]$Absolute_Username
    [string]$Absolute_Make
    [string]$Absolute_Model
    [string]$Absolute_Private_IP
    [string]$Absolute_Public_IP
    [string]$Absolute_Encryption_Status

    # properties from Sophos    
    [string]$Sophos_Name
    [string]$Sophos_Health_Status
    [string]$Sophos_IP
    [string]$Sophos_OS
    [string]$Sophos_Protection
    [string]$Sophos_Last_User
    [string]$Sophos_Last_Active

    #methods
    [void] AddAzureData($info)
    {
        $this.inAzure = $True
        $this.Azure_Display_Name = $info.displayName
        $this.Azure_OS = $info.operatingSystem
        $this.Azure_Join_Type = $info.'joinType (trustType)'
        $this.Azure_Username = $info.userNames
        $this.Azure_Registration_Time = $info.registrationTime
        $this.Azure_Last_Sign_In_Time = $info.approximateLastSignInDateTime
    }

    [void] AddFreshServiceData($info)
    {
        $this.inFreshService = $True
        $this.FreshService_Display_Name = $info.'Display Name'
        $this.FreshService_SN = $info.'Serial Number'
        $this.FreshService_Used_By = $info.'Used By'
        $this.FreshService_Last_Login_By = $info.'Last login by'
        $this.FreshService_Acquisition_Date = $info.'Acquisition Date'
        $this.FreshService_Warranty_Expiry_Date = $info.'Warranty Expiry Date'
        $this.FreshService_Asset_State = $info.'Asset State'
        $this.FreshService_Last_Audit_Date = $info.'Last Audit Date'
    }

    [void] AddAbsoluteData($info)
    {
        $this.inAbsolute = $True
        $this.Absolute_Device_Name = $info.'Device name'
        $this.Absolute_SN = $info.'Serial number'
        $this.Absolute_Last_Connected = $info.'Last connected'
        $this.Absolute_Username = $info.Username
        $this.Absolute_Make = $info.Make
        $this.Absolute_Model = $info.Model
        $this.Absolute_Private_IP = $info.'Local IP address'
        $this.Absolute_Public_IP = $info.'Public IP address'
        $this.Absolute_Encryption_Status = $info.'Encryption status'
    }

    [void] AddSophosData($info)
    {
        $this.inSophos = $True
        $this.Sophos_Name = $info.Name
        $this.Sophos_Health_Status = $info.'Health Status'
        $this.Sophos_IP = $info.IP
        $this.Sophos_OS = $info.OS
        $this.Sophos_Protection = $info.Protection
        $this.Sophos_Last_User = $info.'Last User'
        $this.Sophos_Last_Active = $info.'Last Active'
    }
}

# fields
$sizeWarningThresholdMB = 100
$maxFileCount = 4
$minFileCount = 2

# main
Write-Host "To begin, ensure all device reports are exported as CSV and placed into their own folder."
do
{
    $folderPath = Prompt-FolderPath
    $csvFiles = Get-ChildItem $folderPath -Filter "*.csv" -ErrorAction SilentlyContinue
    $areInputsValid = Validate-Inputs -files $csvFiles -maxSize $sizeWarningThresholdMB -maxFileCount $maxFileCount -minFileCount $minFileCount
}
while (!($areInputsValid))
$baseFile = Prompt-BaseFile
$importedCSVList = Import-CSVList $csvFiles
$azureIndex = Get-AzureFileIndex $importedCSVList
$freshServiceIndex = Get-FreshServiceFileIndex $importedCSVList
$absoluteIndex = Get-AbsoluteFileIndex $importedCSVList
$sophosIndex = Get-SophosFileIndex $importedCSVList
switch ($baseFile)
{
    "Azure" { $reportBuilder = New-Object -TypeName ReportBuilderAzureBase -ArgumentList ($importedCSVList, $azureIndex, $absoluteIndex, $freshServiceIndex, $sophosIndex) }
    "FreshService" { $reportBuilder = New-Object -TypeName ReportBuilderFreshServiceBase -ArgumentList ($importedCSVList, $azureIndex, $absoluteIndex, $freshServiceIndex, $sophosIndex) }
    "Absolute" { $reportBuilder = New-Object -TypeName ReportBuilderAbsoluteBase -ArgumentList ($importedCSVList, $azureIndex, $absoluteIndex, $freshServiceIndex, $sophosIndex) }
    "Sophos" { $reportBuilder = New-Object -TypeName ReportBuilderSophosBase -ArgumentList ($importedCSVList, $azureIndex, $absoluteIndex, $freshServiceIndex, $sophosIndex) }
}
if ($null -ne $reportBuilder )
{
    $report = $reportBuilder.NewReport()
}
Export-Report($report)
Read-Host "Press Enter to end script."