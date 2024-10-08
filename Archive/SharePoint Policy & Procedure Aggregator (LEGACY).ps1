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
    The URL of the SharePoint Online site where the list is located. Default is "https://example.sharepoint.com/sites/Policy".

.PARAMETER ListName
    The name of the SharePoint list to be managed. Default is "Policies and Procedures by Department".

.NOTES
    - The script uses the PnP PowerShell module to interact with SharePoint Online.
    - Make sure you have the necessary permissions to access and modify the SharePoint list.
    - Ensure that the PnP PowerShell module is installed and updated.

.EXAMPLE
    .\SharePoint Policy & Procedure Aggregator (LEGACY).ps1 -SiteUrl "https://example.sharepoint.com/sites/Policy" -ListName "Policies List"

    This example connects to the specified SharePoint Online site, clears the "Department Policies" list, searches for relevant documents, updates or adds them to the list, and removes items older than 60 days.

.Link
    https://github.com/Rokawoo/sharepoint-policy-procedure-aggregator

.ONETIME-SETUP
    1. Run SharePoint Online Management Shell Installer: https://www.microsoft.com/en-US/download/details.aspx?id=35588&msockid=1873099af97a68ec13ce1d1ff8186956
    2. Run in PowerShell: Install-Module -Name SharePointPnPPowerShellOnline -Force -AllowClobber; $env:PNPLEGACYMESSAGE='false'

.AUTHOR
    Roka Awoo
    8/21/2024
#>

param (
    [string]$SiteUrl = "https://example.sharepoint.com/sites/Policy",
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
        Searches for all policy and procedure documents in SharePoint Online sites based on title or category.

    #>
    param (
        [string]$Query = '(contentclass:STS_ListItem_DocumentLibrary) AND (Title:Policy OR Title:Procedure) AND (FileExtension:pdf)'
    )

    try {
        Write-Host "Initiating search across all document libraries in the domain..." -ForegroundColor Yellow

        $results = Submit-PnPSearchQuery -Query $Query -TrimDuplicates $true -All 

        if ($results.ResultRows.Count -eq 0) {
            Write-Yellow "No documents found matching the criteria."
        } else {
            Write-Host "$($results.ResultRows.Count) documents found." -ForegroundColor Green
        }

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
        [string]$DocumentCategory,
        [string]$Department,
        [string]$LastModified,
        [string]$DocumentAuthor
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
                Category = $DocumentCategory
                Department = $Department
                LastModified = $LastModified
                DocumentAuthor = $DocumentAuthor
            }
        } else {
            Write-Yellow "Adding new item: $Title"
            Add-PnPListItem -List $ListName -Values @{
                Title = $Title
                DocumentLink = $DocumentLink
                Category = $DocumentCategory
                Department = $Department
                LastModified = $LastModified
                DocumentAuthor = $DocumentAuthor
            }
        }
    } catch {
        Write-Error "Failed to update or add item: $_"
        throw
    }
}

function Get-DocumentCategory {
    <#
    .SYNOPSIS
        Extracts the document category from a given document string attribute.
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

function Check-UrlConditions {
    <#
    .SYNOPSIS
        Checks if the provided URL meets specified conditions.
    #>
    param (
        [string]$Url
    )

    if ($Url -match "/sites/" -and $Url -match "/Shared Documents/" -and $Url -notmatch "(?i)archive") {
        $slashCount = ($Url -split '/').Count - 1

        if ($slashCount -le 7) {
            return $true
        } else {
            return $false
        }
    } else {
        return $false
    }
}

function AddSpaceBetweenCase {
    <#
    .SYNOPSIS
        Adds a space between (lowercase and uppercase) chars and between (uppercase and uppercase) if the following char is (lowercase).
    #>
    param (
        [string]$inputString
    )

    if (-not $inputString) {
        return $inputString
    }

    $stringBuilder = [System.Text.StringBuilder]::new()
    $length = $inputString.Length

    for ($i = 0; $i -lt $length; $i++) {
        $currentChar = $inputString[$i]
        $stringBuilder.Append($currentChar) | Out-Null

        if ($i -lt $length - 1) {
            $nextChar = $inputString[$i + 1]

            if ([char]::IsLower($currentChar) -and [char]::IsUpper($nextChar)) {
                $stringBuilder.Append(" ") | Out-Null
            }
            elseif ($i -lt $length - 2) {
                $nextNextChar = $inputString[$i + 2]
                if ([char]::IsUpper($currentChar) -and [char]::IsUpper($nextChar) -and [char]::IsLower($nextNextChar)) {
                    $stringBuilder.Append(" ") | Out-Null
                }
            }
        }
    }

    return $stringBuilder.ToString()
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

        if (Check-UrlConditions -Url $Url) {
            $departmentPart = ($Url -split '/sites/')[1] -split '/' | Select-Object -First 1

            $formattedDepartment = AddSpaceBetweenCase -inputString $departmentPart


            Write-Host "Formatted Department: $formattedDepartment"
            return $formattedDepartment
        }

        Write-Warning "URL does not meet conditions. Returning 'Unknown'."
        return "Unknown"
    } catch {
        Write-Error "Failed to extract department from URL: $_"
        return "Unknown"
    }
}

function Format-Authors {
    <#
    .SYNOPSIS
        Formats the Author string by removing emails and spacing Authors properly.
    #>
    param (
        [string]$AuthorString
    )
    $emailPattern = '^[\w\.-]+@[\w\.-]+\.\w+$'

    $formattedAuthors = ($AuthorString -split ';' |
        Where-Object { $_ -notmatch $emailPattern } |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -ne '' }) -join "; "

    return $formattedAuthors
}

# Main Execution
try {
    Connect-ToSharePoint

    $listLastModifiedDate = (Get-PnPList -Identity $ListName).LastItemUserModifiedDate
    Write-Yellow "Policies & Procedures List Last Aggregated: $($listLastModifiedDate)"
    
    Clear-List

    $results = Search-Documents

    foreach ($result in $results) {
        $docTitle = $result.Title
        $docCategory = Get-DocumentCategory -docTitle $docTitle
        $docUrl = $result.Path
        $docLastModified = $result.LastModifiedTime
        $docAuthor = Format-Authors -AuthorString $result.Author

        if ($docUrl -match "\.(doc|docx|pdf)$") {
            $department = Get-DepartmentFromUrl -Url $docUrl
            if ($department -ne "Unknown") {
                Update-Or-AddItem -Title $docTitle -DocumentLink $docUrl -DocumentCategory $docCategory -Department $department -LastModified $docLastModified -DocumentAuthor $docAuthor
            } else {
                Write-Warning "Skipping document with unknown department: $docTitle"
            }

            Write-Host "---"
        }
    }
}
finally {
    $listItemCount = (Get-PnPList -Identity $ListName).ItemCount
    Write-Yellow "Total Documents in List: $listItemCount"
    Disconnect-PnPOnline
    Write-Yellow "Disconnected from SharePoint Online."
}
