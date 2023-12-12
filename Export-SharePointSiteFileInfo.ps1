# Version 1.0

# functions
function Initialize-ColorScheme
{
    Set-Variable -Name "successColor" -Value "Green" -Scope "Script" -Option "Constant"
    Set-Variable -Name "infoColor" -Value "DarkCyan" -Scope "Script" -Option "Constant"
    Set-Variable -Name "warningColor" -Value "Yellow" -Scope "Script" -Option "Constant"
    Set-Variable -Name "failColor" -Value "Red" -Scope "Script" -Option "Constant"
}

function Show-Introduction
{
    Write-Host "This script does some stuff..." -ForegroundColor $infoColor
    Read-Host "Press Enter to continue"
}

function Use-Module($moduleName)
{    
    $keepGoing = -not(Test-ModuleInstalled $moduleName)
    while ($keepGoing)
    {
        Prompt-InstallModule $moduleName
        Test-SessionPrivileges
        Install-Module $moduleName

        if ((Test-ModuleInstalled $moduleName) -eq $true)
        {
            Write-Host "Importing module..." -ForegroundColor $infoColor
            Import-Module $moduleName
            $keepGoing = $false
        }
    }
}

function Test-ModuleInstalled($moduleName)
{    
    $module = Get-Module -Name $moduleName -ListAvailable
    return ($null -ne $module)
}

function Prompt-InstallModule($moduleName)
{
    do 
    {
        Write-Host "$moduleName module is required." -ForegroundColor $infoColor
        $confirmInstall = Read-Host -Prompt "Would you like to install the module? (y/n)"
    }
    while ($confirmInstall -inotmatch "^\s*y\s*$") # regex matches a y but allows spaces
}

function Test-SessionPrivileges
{
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentSessionIsAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if ($currentSessionIsAdmin -ne $true)
    {
        Write-Host ("Please run script with admin privileges.`n" +
            "1. Open Powershell as admin.`n" +
            "2. CD into script directory.`n" +
            "3. Run .\scriptname`n") -ForegroundColor $failColor
        Read-Host "Press Enter to exit"
        exit
    }
}

function TryConnect-MgGraph($scopes)
{
    $connected = Test-ConnectedToMgGraph
    while (-not($connected))
    {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor $infoColor

        if ($null -ne $scopes)
        {
            Connect-MgGraph -Scopes $scopes -ErrorAction SilentlyContinue | Out-Null
        }
        else
        {
            Connect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }
        
        $connected = Test-ConnectedToMgGraph
        if (-not($connected))
        {
            Read-Host "Failed to connect to Microsoft Graph. Press Enter to try again"
        }
        else
        {
            Write-Host "Successfully connected!" -ForegroundColor $successColor
        }
    }    
}

function Test-ConnectedToMgGraph
{
    return $null -ne (Get-MgContext)
}

function Prompt-YesOrNo($question)
{
    Write-Host "$question`n[Y] Yes  [N] No"

    do
    {
        $response = Read-Host
        $validResponse = $response -imatch '^\s*[yn]\s*$' # regex matches y or n but allows spaces
        if (-not($validResponse)) 
        {
            Write-Warning "Please enter y or n."
        }
    }
    while (-not($validResponse))

    if ($response -imatch '^\s*y\s*$') # regex matches a y but allows spaces
    {
        return $true
    }
    return $false
}

function PromptFor-Site
{
    do
    {
        $url = Read-Host "Enter site URL"
        $formattedUrl = Format-URL $url # Get the format needed for the API call.
        try
        {
            # Uri is https://graph.microsoft.com/v1.0/sites/domain.sharepoint.com:/sites/siteName/?select=Id,DisplayName
            $site = Invoke-MgGraphRequest -Method "Get" -Uri "$baseUri/sites/$formattedUrl/?`$select=Id,DisplayName"            
        }
        catch
        {
            Write-Warning "There was an issue getting site. Please try again."
            $keepGoing = $true
            continue
        }

        if ($null -eq $site)
        {
            Write-Warning "Site not found or access denied. Please try again."
            $keepGoing = $true
            continue
        }

        $keepGoing = $false
    }
    while ($keepGoing)

    Write-Host "Site found: $($site.DisplayName)" -ForegroundColor $successColor
    return $site
}

function Format-URL($url)
{
    $url = $url.Trim() # remove leading and trailing spaces
    $url = $url.Replace('https://', '') # remove https://
    $url = $url.Replace('.com', '.com:') # add colon
    return $url # result is domain.sharepoint.com:/sites/siteName
}

function New-MetaReport
{
    return [PSCustomObject]@{        
        TotalStorageConsumed            = 0
        StorageConsumedCurrentVersions  = 0
        CountItems                      = 0
        CountFolders                    = 0
        CountFiles                      = 0        
        CountDrives                     = 0
        Drives                          = (New-Object System.Collections.Generic.List[object])
        CountSubsites                   = 0
        Subsites                        = (New-Object System.Collections.Generic.List[object])
        CountLists                      = 0
        Lists                           = (New-Object System.Collections.Generic.List[object])
        CountNotebooks                  = 0
        Notebooks                       = (New-Object System.Collections.Generic.List[object])
    }
}

function Get-Drives($site)
{
    try
    {
        $drives = Invoke-MgGraphRequest -Method "Get" -Uri "$baseUri/sites/$($site.Id)/drives?`$select=Id,Name,DriveType,Quota,WebUrl"
        $drives = $drives.Value
    }
    catch
    {
        Write-Host "There was an issue getting site drives. Exiting script." -ForegroundColor $failColor
        exit
    }

    if ($null -eq $drives)
    {
        Write-Host "Was unable to find any drives in this site. Exiting script." -ForegroundColor $failColor
        exit
    }

    return $drives
}

function Get-DriveLookup($drives)
{
    $driveLookup = @{}
    foreach ($drive in $drives)
    {
        $driveLookup.Add($drive.Id, $drive.Name)
    }
    return $driveLookup
}

function Export-ItemsInAllDrives($drives, $exportPath)
{
    foreach ($drive in $drives)
    {
        Export-ItemsInDrive -Drive $drive -ExportPath $exportPath
    }
    Write-Host "Finished exporting to $exportPath" -ForegroundColor $successColor
}

function Export-ItemsInDrive($drive, $exportPath)
{
    Export-ItemsRecursively -Uri "$baseUri/drives/$($drive.Id)/items/root/children" -ExportPath $exportPath
}

function Export-ItemsRecursively($uri, $exportPath)
{
    if ($script:itemCounter -ge 100) { return } # For debugging
    
    # Uri is $baseUri/drives/$($drive.Id)/items/root/children or $baseUri/drives/{drive-id}/items/{item-id}/children
    $itemPage = Invoke-GraphRequest -Method "Get" -Uri $uri
    if ($itemPage.Value.Count -eq 0) { return }

    # For debugging
    if ($itemPage.Value.Count -ge 200)
    {
        # This page has over 200 items: 
        # https://blueravensolar.sharepoint.com/:f:/s/BusinessIntelligenceTeam/EoofhFpfjDdPh0Cl6OOp-1UB0N856UdcGFGelm2FbeN4gQ?e=JLGS69
        Write-Host "Over 200 items! Count is: $($itemPage.Value.Count)" -ForegroundColor "DarkMagenta"
        Write-Host "Starting with: $($itemPage.Value[0].Name)" -ForegroundColor "DarkMagenta"
    }

    foreach ($item in $itemPage.Value)
    {
        if ($script:itemCounter -ge 100) { return } # For debugging
        
        # For debugging
        Write-Host "Exporting $($item.Name)" -ForegroundColor $infoColor

        if ($getVersionInfo)
        {
            $itemUri = ($uri -Replace 'items\/.+\/children', "items/$($item.Id)")
        }
        $itemRecord = New-ItemRecord -Item $item -ItemUri $itemUri
        Update-MetaReportWithItem $itemRecord
        $itemRecord | Export-CSV -Path $exportPath -Append -NoTypeInformation

        $isFolder = Test-ItemIsFolder $item
        if ($isFolder)
        {
            # String replace using regex.
            $uri = ($uri -Replace 'items\/.+\/children', "items/$($item.Id)/children")

            # Uri is $baseUri/drives/$($drive.Id)/items/$($item.Id)/children
            Export-ItemsRecursively -Uri $uri -ExportPath $exportPath
        }
    }

    $nextLink = $itemPage."@odata.nextLink"
    if ($nextLink)
    {
        Export-ItemsRecursively -Uri $nextLink -ExportPath $exportPath
    }
}

function Test-ItemIsFolder($item)
{
    return $item.ContainsKey("folder")
}

function Test-ItemIsFile($item)
{
    return $item.ContainsKey("file")
}

function New-ItemRecord($item, $itemUri)
{
    $script:itemCounter++ # For debugging
    
    $isFolder = Test-ItemIsFolder $item
    if ($isFolder)
    {
        $type = [ItemType]::Folder
        $childCount = $item.Folder.ChildCount
    }

    $isFile = Test-ItemIsFile $item
    if ($isFile)
    {
        $type = [ItemType]::File
        if ($getVersionInfo)
        {
            # Uri is $baseUri/drives/$($drive.Id)/items/$($item.Id)/versions
            $versions = Invoke-MgGraphRequest -Method "Get" -Uri "$itemUri/versions?`$select=size"
            $versions = $versions.Value
            $versionCount = $versions.Count
            $versionsTotalSizeInBytes = Get-VersionsTotalSize $versions
            $versionsTotalSizeFormatted = Format-FileSize $versionsTotalSizeInBytes            
        }
    }

    return [PSCustomObject]@{
        ParentPath               = (Get-ReadablePath $item.ParentReference.Path)
        Name                     = $item.Name        
        Type                     = $type
        ChildCount               = $childCount
        Size                     = (Format-FileSize $item.Size)
        SizeInBytes              = $item.Size
        VersionCount             = $versionCount
        VersionsTotalSize        = $versionsTotalSizeFormatted
        VersionsTotalSizeInBytes = $versionsTotalSizeInBytes        
        CreatedBy                = $item.CreatedBy.User.DisplayName
        CreatedDateTime          = $item.CreatedDateTime
        LastModifiedBy           = $item.LastModifiedBy.User.DisplayName
        LastModifiedDateTime     = $item.LastModifiedDateTime
        Url                      = $item.WebUrl
    }
}

function Get-VersionsTotalSize($versions)
{
    $totalSize = 0
    foreach ($version in $versions)
    {
        $totalSize += $version.Size
    }
    return $totalSize
}

function Get-ReadablePath($path)
{
    $driveId = Get-SubstringWithRegex -String $path -Regex '(?<=drives\/).+?(?=\/)'
    $driveName = $driveLookup[$driveId]
    $path = ($path -Replace '(?<=drives\/).+?(?=\/)', $driveName) # replace driveId with driveName
    $path = $path.Replace('/root:', '')
    return $path
}

function Get-SubstringWithRegex($string, $regex)
{
    if ($string -match $regex)
    {
        # $matches is an automatic variable that is populated when using the -match operator.
        return $matches[0]
    }
    else
    {
        Write-Warning "Could not find substring in string: $string with regex: $regex"
    }
}

function Format-FileSize($sizeInBytes)
{
    if ($sizeInBytes -lt 1KB)
    {
        $formattedSize = $sizeInBytes.ToString() + " B"
    }
    elseif ($sizeInBytes -lt 1MB)
    {
        $formattedSize = $sizeInBytes / 1KB
        $formattedSize = ("{0:n2}" -f $formattedSize) + " KB"
    }
    elseif ($sizeInBytes -lt 1GB)
    {
        $formattedSize = $sizeInBytes / 1MB
        $formattedSize = ("{0:n2}" -f $formattedSize) + " MB"
    }
    elseif ($sizeInBytes -lt 1TB)
    {
        $formattedSize = $sizeInBytes / 1GB
        $formattedSize = ("{0:n2}" -f $formattedSize) + " GB"
    }
    elseif ($sizeInBytes -ge 1TB)
    {
        $formattedSize = $sizeInBytes / 1TB
        $formattedSize = ("{0:n2}" -f $formattedSize) + " TB"
    }
    return $formattedSize
}

function Update-MetaReportWithItem($itemRecord)
{
    $script:metaReport.CountItems++
    $script:metaReport.StorageConsumedCurrentVersions += $itemRecord.SizeInBytes
    if ($itemRecord.Type -eq [ItemType]::File)
    {
        $script:metaReport.CountFiles++
        $script:metaReport.TotalStorageConsumed += $itemRecord.VersionsTotalSizeInBytes
    }
    elseif ($itemRecord.Type -eq [ItemType]::Folder)
    {
        $script:metaReport.CountFolders++
        $script:metaReport.TotalStorageConsumed += $itemRecord.SizeInBytes
    }    
}

function Get-Subsites($site)
{
    $subsites = Invoke-MgGraphRequest -Method "Get" -Uri "$baseUri/sites/$($site.Id)/sites?`$select=Name,DisplayName,WebUrl"
    return $subsites.Value
}

function Update-MetaReportWithSubsites($subsites)
{
    $script:metaReport.CountSubsites = $subsites.Count
    foreach ($site in $subsites)
    {
        $script:metaReport.Subsites.Add($site)
    }    
}

function Get-Lists($site)
{
    $lists = Invoke-MgGraphRequest -Method "Get" -Uri "$baseUri/sites/$($site.Id)/lists?`$select=Name,DisplayName,Description,WebUrl,List"
    return $lists.Value
}

function Update-MetaReportWithLists($lists)
{
    $script:metaReport.CountLists = $lists.Count
    foreach ($list in $lists)
    {
        $script:metaReport.Lists.Add($list)
    }
}

function Get-Notebooks($site)
{
    $notebooks = Invoke-MgGraphRequest -Method "Get" -Uri "$baseUri/sites/$($site.Id)/onenote/notebooks?`$select=DisplayName,Links"
    return $notebooks.Value
}

function Update-MetaReportWithNotebooks($notebooks)
{
    $script:metaReport.CountNotebooks = $notebooks.Count
    foreach ($notebook in $notebooks)
    {
        $script:metaReport.Notebooks.Add($notebook)
    }
}

function Update-MetaReportWithDrives($drives)
{
    $script:metaReport.CountDrives = $drives.Count
    foreach ($drive in $drives)
    {
        $script:metaReport.Drives.Add($drive)
    }       
}

function Show-MetaReport
{
    Show-Separator -Title "Meta-report"

    $topSection = [PSCustomObject]@{
        TotalStorageConsumed                     = (Format-FileSize $script:metaReport.TotalStorageConsumed)
        "TotalStorageConsumed (Bytes)"           = $script:metaReport.TotalStorageConsumed
        StorageConsumedCurrentVersions           = (Format-FileSize $script:metaReport.StorageConsumedCurrentVersions)
        "StorageConsumedCurrentVersions (Bytes)" = $script:metaReport.StorageConsumedCurrentVersions        
        PercentConsumedByCurrentVersions         = (Get-Percent -Divisor $script:metaReport.StorageConsumedCurrentVersions -Dividend $script:metaReport.TotalStorageConsumed)
        CountDrives                              = $script:metaReport.CountDrives
        CountItems                               = $script:metaReport.CountItems
        CountFolders                             = $script:metaReport.CountFolders
        CountFiles                               = $script:metaReport.CountFiles
        CountSubsites                            = $script:metaReport.CountSubsites
        CountLists                               = $script:metaReport.CountLists
        CountNotebooks                           = $script:metaReport.CountNotebooks
    }
    $topSection | Out-Host

    Show-Separator -Title "Drives"
    $script:metaReport.Drives | Out-Host

    Show-Separator -Title "Lists"
    $script:metaReport.Lists | Out-Host

    Show-Separator -Title "Notebooks"
    $script:metaReport.Notebooks | Out-Host

    Show-Separator -Title "Subsites"
    $script:metaReport.Subsites | Out-Host
}

function Show-Separator($title, [ConsoleColor]$color = "DarkCyan", [switch]$noLineBreaks)
{
    if ($title)
    {
        $separator = " $title "
    }
    else
    {
        $separator = ""
    }

    # Truncate if it's too long.
    If (($separator.length - 6) -gt ((Get-host).UI.RawUI.BufferSize.Width))
    {
        $separator = $separator.Remove((Get-host).UI.RawUI.BufferSize.Width - 5)
    }

    # Pad with dashes.
    $separator = "--$($separator.PadRight(((Get-host).UI.RawUI.BufferSize.Width)-3,"-"))"

    if (-not($noLineBreaks))
    {        
        # Add line breaks.
        $separator = "`n$separator`n"
    }

    Write-Host $separator -ForegroundColor $color
}

function Get-Percent($divisor, $dividend)
{
    $percent = $divisor / $dividend * 100
    $roundedToInt = [Math]::Round($percent)
    return "$roundedToInt%"
}

function New-TimeStamp
{
    return (Get-Date -Format yyyy-MM-dd-hh-mmtt).ToString()
}

# main
Initialize-ColorScheme
Show-Introduction
Use-Module "Microsoft.Graph.Authentication"
TryConnect-MgGraph -Scopes "Sites.Read.All", "Notes.Read.All"
Set-Variable -Name "getVersionInfo" -Value (Prompt-YesOrNo "Would you like to get file version info? (Takes much longer as it must enumerate each version.)") -Scope "Script" -Option "Constant"
Set-Variable -Name "baseUri" -Value "https://graph.microsoft.com/v1.0" -Scope "Script" -Option "Constant"
$site = PromptFor-Site

$script:itemCounter = 0 # For debugging

enum ItemType
{
    File
    Folder
}

$script:metaReport = New-MetaReport

$drives = Get-Drives $site
Set-Variable -Name "driveLookup" -Value (Get-DriveLookup $drives) -Scope "Script" -Option "Constant"
Export-ItemsInAllDrives -Drives $drives -ExportPath "$PSScriptRoot/SharePoint $($site.DisplayName) File Info $(New-TimeStamp).csv"

Update-MetaReportWithDrives $drives
Update-MetaReportWithSubsites (Get-Subsites $site)
Update-MetaReportWithLists (Get-Lists $site)
Update-MetaReportWithNotebooks (Get-Notebooks $site)
Show-MetaReport

Read-Host "Press Enter to exit"