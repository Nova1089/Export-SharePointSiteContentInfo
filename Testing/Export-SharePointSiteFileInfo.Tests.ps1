BeforeAll {
    # Optional
    # BeforeAll runs once at the beginning of the file.

    # Dot sourcing the script to test. (Must comment out main functionality.)
    . $PSScriptRoot\..\Export-SharePointSiteFileInfo.ps1

    Initialize-ColorScheme
    TryConnect-MgGraph -Scopes "Sites.Read.All", "Notes.Read.All"
    Set-Variable -Name "baseUri" -Value "https://graph.microsoft.com/v1.0" -Scope "Script" -Option "Constant"
    $script:metaReport = New-Object MetaReport
    Mock Read-Host { return Get-Content "$PSScriptRoot/site.txt" }
    $site = PromptFor-Site
    $drives = Get-Drives $site
    Set-Variable -Name "driveLookup" -Value (Get-DriveLookup $drives) -Scope "Script" -Option "Constant"    
}

Describe "Get-Notebooks" {
    BeforeEach {
        # Optional
        # Runs once before each test (It block) within the current Describe or Context block.
    }

    Context "When passed valid site" {
        It "Should return a hashtable" {
            $notebooks = Get-Notebooks $site
            $notebooks | Should -BeOfType [HashTable]
        }
    }

    AfterEach {
        # Optional
        # Runs once after each test (It block) within the current Describe or Context block.
    }
}

AfterAll {
    # Optional
    # Runs once at the end of the file.
}