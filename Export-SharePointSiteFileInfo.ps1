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
    Write-Host "This script exports info about the content on a SharePoint site and helps determine where space is being occupied." -ForegroundColor $infoColor
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

        if ($formattedUrl -inotmatch '.+\.sharepoint\.com:\/sites\/.+')
        {
            Write-Warning "Invalid site URL. Provide a URL in the format: domain.sharepoint.com/sites/siteName"
            $keepGoing = $true
            continue
        }

        try
        {
            # URI: https://graph.microsoft.com/v1.0/sites/domain.sharepoint.com:/sites/siteName
            # Docs: https://learn.microsoft.com/en-us/graph/api/site-get
            $uri = "$baseUri/sites/$formattedUrl/?`$select=Id,DisplayName"  
            $site = Invoke-MgGraphRequest -Method "Get" -Uri $uri           
        }
        catch
        {
            $errorRecord = $_
            if ($errorRecord.Exception.Response.StatusCode -eq "Forbidden")
            {
                Write-Warning "Response: 403 Forbidden. You are not authorized to this site."
            }
            else
            {
                Write-Warning "There was an issue getting site. Please try again."
                Write-Host "Call to URI: $uri" -ForegroundColor $warningColor
                Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
            }            
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
    $url = $url.Replace('.com/', '.com:/') # add colon
    return $url # result is domain.sharepoint.com:/sites/siteName
}

function Get-Drives($site)
{
    try
    {
        $uri = "$baseUri/sites/$($site.Id)/drives?`$select=Id,Name,DriveType,Quota,WebUrl"
        $drives = Invoke-MgGraphRequest -Method "Get" -Uri $uri        
    }
    catch
    {
        $errorRecord = $_
        Write-Host "There was an issue getting site drives. Exiting script." -ForegroundColor $failColor
        Write-Host "Call to URI: $uri" -ForegroundColor $failColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $failColor
        exit
    }

    $drives = $drives.Value
    if (($null -eq $drives) -or ($drives.Count -eq 0))
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
    
    # URI: $baseUri/drives/{drive-id}/items/root/children or $baseUri/drives/{drive-id}/items/{item-id}/children
    # Docs: https://learn.microsoft.com/en-us/graph/api/driveitem-list-children
    $itemPage = Invoke-GraphRequest -Method "Get" -Uri $uri
    if ($itemPage.Value.Count -eq 0) { return }

    # For debugging
    if ($itemPage.Value.Count -ge 200)
    {
        Write-Host "Over 200 items! Count is: $($itemPage.Value.Count)" -ForegroundColor "DarkMagenta"
        Write-Host "Starting with: $($itemPage.Value[0].Name)" -ForegroundColor "DarkMagenta"
    }

    foreach ($item in $itemPage.Value)
    {
        if ($script:itemCounter -ge 100) { return } # For debugging
        
        if ($script:getVersionInfo)
        {
            $itemUri = ($uri -Replace 'items\/.+\/children', "items/$($item.Id)")
        }
        $itemRecord = New-ItemRecord -Item $item -ItemUri $itemUri        
        $script:metaReport.AddItem($itemRecord)
        Write-Progress -Activity "Exporting items..." -Status "$($script:metaReport.CountItems): $($itemRecord.ParentPath)/$($itemRecord.Name)"
        $itemRecord | Export-CSV -Path $exportPath -Append -NoTypeInformation

        $isFolder = Test-ItemIsFolder $item
        if ($isFolder)
        {
            # String replace using regex.
            $uri = ($uri -Replace 'items\/.+\/children', "items/$($item.Id)/children")

            # URI: $baseUri/drives/{drive-id}/items/{item-id}/children
            # Docs: https://learn.microsoft.com/en-us/graph/api/driveitem-list-children            
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
        if ($script:getVersionInfo)
        {
            try
            {
                # URI: $baseUri/drives/{drive-id}/items/{item-id}/versions
                # Docs: https://learn.microsoft.com/en-us/graph/api/driveitem-list-versions
                $uri = "$itemUri/versions?`$select=size"
                $versions = Invoke-MgGraphRequest -Method "Get" -Uri $uri
            }
            catch
            {
                $errorRecord = $_
                Write-Warning "There was an issue getting versions for item: $($item.Name)."
                Write-Host "Call to URI: $uri" -ForegroundColor $warningColor
                Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
            }
            
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

function Get-Lists($site)
{
    try
    {
        $uri = "$baseUri/sites/$($site.Id)/lists?`$select=Name,DisplayName,Description,WebUrl,List"
        $lists = Invoke-MgGraphRequest -Method "Get" -Uri $uri
    }
    catch
    {
        $errorRecord = $_
        Write-Warning "There was an issue getting lists."
        Write-Host "Call to URI: $uri" -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor        
    }
    
    return $lists.Value
}

function Get-Notebooks($site)
{
    try
    {
        $uri = "$baseUri/sites/$($site.Id)/onenote/notebooks?`$select=DisplayName,Links"
        $notebooks = Invoke-MgGraphRequest -Method "Get" -Uri $uri
    }
    catch
    {
        $errorRecord = $_
        Write-Warning "There was an issue getting notebooks."
        Write-Host "Call to URI: $uri" -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
    }
    
    return $notebooks.Value
}

function Get-Subsites($site)
{
    try
    {
        $uri = "$baseUri/sites/$($site.Id)/sites?`$select=Name,DisplayName,WebUrl"
        $subsites = Invoke-MgGraphRequest -Method "Get" -Uri $uri
    }
    catch
    {
        $errorRecord = $_
        Write-Warning "There was an issue getting subsites."
        Write-Host "Call to URI: $uri" -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
    }
    
    return $subsites.Value
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

class MetaReport
{
    [Int64]$TotalStorageConsumed
    [Int64]$StorageConsumedCurrentVersions
    [Int]$CountItems
    [Int]$CountFolders
    [Int]$CountFiles
    [Int]$CountDrives
    [System.Collections.Generic.List[object]]$Drives
    [Int]$CountLists
    [System.Collections.Generic.List[object]]$Lists
    [Int]$CountNotebooks
    [System.Collections.Generic.List[object]]$Notebooks
    [Int]$CountSubsites
    [System.Collections.Generic.List[object]]$Subsites

    MetaReport()
    {
        $this.Drives = New-Object System.Collections.Generic.List[object]        
        $this.Lists = New-Object System.Collections.Generic.List[object]
        $this.Notebooks = New-Object System.Collections.Generic.List[object]
        $this.Subsites = New-Object System.Collections.Generic.List[object]
    }
    
    AddItem($itemRecord)
    {
        $this.CountItems++
        $this.StorageConsumedCurrentVersions += $itemRecord.SizeInBytes
        if ($itemRecord.Type -eq [ItemType]::File)
        {
            $this.CountFiles++
            $this.TotalStorageConsumed += $itemRecord.VersionsTotalSizeInBytes
        }
        elseif ($itemRecord.Type -eq [ItemType]::Folder)
        {
            $this.CountFolders++
            $this.TotalStorageConsumed += $itemRecord.SizeInBytes
        }
    }

    AddDrives($drives)
    {
        if ($drives -is [HashTable])
        {
            $this.CountDrives++
        }
        else
        {
            $this.CountDrives += $drives.Count
        }

        foreach ($drive in $drives)
        {
            $drive = [PSCustomObject]@{
                Name         = $drive.Name
                DriveType    = $drive.DriveType
                "Size"       = (Format-FileSize $drive.Quota.Used)
                "QuotaTotal" = (Format-FileSize $drive.Quota.Total)     
                URL          = $drive.WebUrl         
            }
            $this.Drives.Add($drive)
        }  
    }

    AddLists($lists)
    {
        if ($lists -is [HashTable])
        {
            $this.CountLists++
        }
        else
        {
            $this.CountLists += $lists.Count
        }

        foreach ($list in $lists)
        {
            $list = [PSCustomObject]@{
                Name        = $list.Name
                DisplayName = $list.DisplayName
                Description = $list.Description
                Hidden      = $list.List.Hidden
                URL         = $list.WebUrl
            }
            $this.Lists.Add($list)
        }
    }

    AddNotebooks($notebooks)
    {
        if ($notebooks -is [HashTable])
        {
            $this.CountNotebooks++
        }
        else
        {
            $this.CountNotebooks += $notebooks.Count
        }
        
        foreach ($notebook in $notebooks)
        {
            $notebook = [PSCustomObject]@{
                DisplayName = $notebook.DisplayName
                URL         = $notebook.Links.OneNoteWebUrl.Href
            }
            $this.Notebooks.Add($notebook)
        }
    }

    AddSubSites($subsites)
    {
        if ($subsites -is [HashTable])
        {
            $this.CountSubsites++
        }
        else
        {
            $this.CountSubsites += $subsites.Count
        }
        
        foreach ($site in $subsites)
        {
            $site = [PSCustomObject]@{
                Name        = $site.Name
                DisplayName = $site.DisplayName
                URL         = $site.WebUrl
            }
            $this.Subsites.Add($site)
        }  
    }

    Show()
    {
        Show-Separator -Title "Meta-report"

        if ($script:getVersionInfo)
        {
            $totalStorageOutput = Format-FileSize $this.TotalStorageConsumed
            $percentOutput = Get-Percent -Divisor $this.StorageConsumedCurrentVersions -Dividend $this.TotalStorageConsumed
        }
        else
        {
            $totalStorageOutput = "Get version info when running script for accurate number."
            $percentOutput = "Get version info when running script for accurate number."
        }        

        $topSection = [PSCustomObject]@{
            TotalStorageConsumed                     = $totalStorageOutput
            StorageConsumedCurrentVersions           = (Format-FileSize $this.StorageConsumedCurrentVersions)     
            PercentConsumedByCurrentVersions         = $percentOutput
            CountDrives                              = $this.CountDrives
            CountItems                               = $this.CountItems
            CountFolders                             = $this.CountFolders
            CountFiles                               = $this.CountFiles            
            CountLists                               = $this.CountLists
            CountNotebooks                           = $this.CountNotebooks
            CountSubsites                            = $this.CountSubsites
        }
        $topSection | Out-Host

        Show-Separator -Title "Drives"
        $this.Drives | Out-Host

        Show-Separator -Title "Lists"
        $this.Lists | Out-Host

        Show-Separator -Title "Notebooks"
        $this.Notebooks | Out-Host

        Show-Separator -Title "Subsites"
        $this.Subsites | Out-Host
    }
}

enum ItemType
{
    File
    Folder
}

# main
Initialize-ColorScheme
Show-Introduction
Use-Module "Microsoft.Graph.Authentication"
TryConnect-MgGraph -Scopes "Sites.Read.All", "Notes.Read.All"

Set-Variable -Name "getVersionInfo" -Value (Prompt-YesOrNo "Would you like to get file version info? (Takes much longer as it must enumerate each version.)") -Scope "Script" -Option "Constant"
Set-Variable -Name "baseUri" -Value "https://graph.microsoft.com/v1.0" -Scope "Script" -Option "Constant"
$script:itemCounter = 0 # For debugging
$script:metaReport = New-Object MetaReport

$site = PromptFor-Site
$drives = Get-Drives $site
Set-Variable -Name "driveLookup" -Value (Get-DriveLookup $drives) -Scope "Script" -Option "Constant"
Export-ItemsInAllDrives -Drives $drives -ExportPath "$PSScriptRoot/SharePoint $($site.DisplayName) File Info $(New-TimeStamp).csv"

$script:metaReport.AddDrives($drives)
$script:metaReport.AddSubSites((Get-Subsites $site))
$script:metaReport.AddLists((Get-Lists $site))
$script:metaReport.AddNotebooks((Get-Notebooks $site))
$script:metaReport.Show()

Read-Host "Press Enter to exit"