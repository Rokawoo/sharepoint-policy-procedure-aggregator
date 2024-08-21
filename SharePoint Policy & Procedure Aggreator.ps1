<#
.SYNOPSIS
    This script connects to a SharePoint Online site's SharePoint list and aggregates all PDF documentscontaining 'Policy' or 'Procedure' from all other sites within the SharePoint domain.

.DESCRIPTION
    The script performs the following tasks:
    1. Connects to a SharePoint Online site using PnP PowerShell.
    2. Clears the specified SharePoint list by removing all items.
    3. Searches for documents from sites in the SharePoint domain using a specified query.
    4. Updates existing items or adds new items to the list based on the search results.
    5. Disconnects from SharePoint Online.

.PARAMETER SiteUrl
    The URL of the SharePoint Online site where the list is located. Default is "https://rhoadsindustries.sharepoint.com/sites/Policy".

.PARAMETER ListName
    The name of the SharePoint list to be managed. Default is "Policies and Procedures by Department".

.NOTES
    - The script uses the PnP PowerShell module to interact with SharePoint Online.
    - Make sure you have the necessary permissions to access and modify the SharePoint list.
    - Ensure that the PnP PowerShell module is installed and updated.

.EXAMPLE
    .\ManageSharePointList.ps1 -SiteUrl "https://example.sharepoint.com/sites/Policy" -ListName "Policies List"

    This example connects to the specified SharePoint Online site, clears the "Department Policies" list, searches for relevant documents, updates or adds them to the list, and removes items older than 60 days.

.Link
    https://github.com/Rokawoo/rhoads-sharepoint-policy-procedure-aggregator

.ONETIME-SETUP
    1. Run SharePoint Online Management Shell Installer: https://www.microsoft.com/en-US/download/details.aspx?id=35588&msockid=1873099af97a68ec13ce1d1ff8186956
    2. Run in Terminal: Install-Module -Name SharePointPnPPowerShellOnline -Force -AllowClobber

.AUTHOR
    Augustus Sroka
#>

param (
    [string]$SiteUrl = "https://rhoadsindustries.sharepoint.com/sites/Policy",
    [string]$ListName = "Policies & Procedures by Department"
)

function Write-Yellow {
    param (
        [string]$Text
    )
    Write-Host $Text -ForegroundColor Yellow
}

function Connect-ToSharePoint {
    <#
    .SYNOPSIS
        Connects to the specified SharePoint Online site.
    #>
    try {
        Write-Yellow "Connecting to SharePoint Online..."
        Connect-PnPOnline -Url $SiteUrl -UseWebLogin
        Write-Yellow "Connected to SharePoint Online."
    } catch {
        Write-Error "Failed to connect to SharePoint: $_"
        throw
    }
}

function Clear-List {
    <#
    .SYNOPSIS
        Clears all items from the specified SharePoint list in batches.
    #>
    [CmdletBinding()]
    param (
        [int]$BatchSize = 200
    )

    try {
        Write-Yellow "Clearing the SharePoint list..."
        do {
            $items = Get-PnPListItem -List $ListName -PageSize $BatchSize
            if ($items.Count -eq 0) { break }
            $items | ForEach-Object {
                Remove-PnPListItem -List $ListName -Identity $_.Id -Force
            }
        } while ($items.Count -eq $BatchSize)
        Write-Yellow "List cleared successfully."
    } catch {
        Write-Error "Failed to clear the list: $_"
        throw
    }
}

function Search-Documents {
    <#
    .SYNOPSIS
        Searches for documents in the SharePoint Online site using the provided query.
    #>
    param (
        [string]$Query = "Title:Policy OR Title:Procedure AND FileExtension:pdf"
    )

    try {
        Write-Yellow "Searching for documents..."
        $results = Submit-PnPSearchQuery -Query $Query
        Write-Yellow "Search completed."
        return $results.ResultRows
    } catch {
        Write-Error "Search query failed: $_"
        throw
    }
}

function Update-Or-AddItem {
    <#
    .SYNOPSIS
        Updates an existing item or adds a new item to the SharePoint list.
    #>
    param (
        [string]$Title,
        [string]$DocumentLink,
        [string]$DocumentType,
        [string]$Department,
        [string]$LastModified,
        [string]$Author
    )

    try {
        $existingItem = Get-PnPListItem -List $ListName -Query @"
            <View>
                <Query>
                    <Where>
                        <Eq>
                            <FieldRef Name='Title'/>
                            <Value Type='Text'>$Title</Value>
                        </Eq>
                    </Where>
                </Query>
            </View>
"@

        if ($existingItem) {
            Write-Yellow "Updating existing item: $Title"
            Set-PnPListItem -List $ListName -Identity $existingItem.Id -Values @{
                DocumentLink = $DocumentLink
                DocumentType = $DocumentType
                Department = $Department
                LastModified = $LastModified
                Author = $Author
            }
        } else {
            Write-Yellow "Adding new item: $Title"
            Add-PnPListItem -List $ListName -Values @{
                Title = $Title
                DocumentLink = $DocumentLink
                DocumentType = $DocumentType
                Department = $Department
                LastModified = $LastModified
                Author = $Author
            }
        }
    } catch {
        Write-Error "Failed to update or add item: $_"
        throw
    }
}

function Get-DocumentType {
    <#
    .SYNOPSIS
        Extracts the document type from a given document title.
    #>
    param (
        [string]$docTitle
    )

    if ($docTitle -match "policy" -and $docTitle -match "procedure") {
        return "Policy & Procedure"
    }
    elseif ($docTitle -match "policy") {
        return "Policy"
    }
    else {
        return "Procedure"
    }
}

function Get-DepartmentFromUrl {
    <#
    .SYNOPSIS
        Extracts the department name from a given SharePoint document URL.
    #>
    param (
        [string]$Url
    )

    try {
        Write-Host "Processing URL: $Url"

        if ($Url -match "/sites/") {
            $departmentPart = ($Url -split '/sites/')[1] -split '/' | Select-Object -First 1
            Write-Host "Extracted Department: $departmentPart"
            return $departmentPart
        }

        Write-Warning "URL does not contain '/sites/'. Returning 'Unknown'."
        return "Unknown"
    } catch {
        Write-Error "Failed to extract department from URL: $_"
        return "Unknown"
    }
}

# Main Execution
try {
    Connect-ToSharePoint

    $ListLastModifiedDate = (Get-PnPList -Identity $ListName).LastItemUserModifiedDate
    Write-Yellow "Policies & Procedures List Last Aggregated: $($ListLastModifiedDate)"
    
    Clear-List

    $results = Search-Documents

    foreach ($result in $results) {
        $docTitle = $result.Title
        $docUrl = $result.Path
        $lastFileModifiedDate = $result.LastModifiedTime
        $docType = Get-DocumentType -docTitle $docTitle

        if ($docUrl -match "\.pdf$") {
            $department = Get-DepartmentFromUrl -Url $docUrl
            if ($department -ne "Unknown") {
                Update-Or-AddItem -Title $docTitle -DocumentLink $docUrl -DocumentType $docType -Department $department -LastModified $lastFileModifiedDate
            } else {
                Write-Warning "Skipping document with unknown department: $docTitle"
            }
        }
    }
}
finally {
    Disconnect-PnPOnline
    Write-Yellow "Disconnected from SharePoint Online."
}
