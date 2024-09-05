<#
.SYNOPSIS
    This script connects to a SharePoint Online site's SharePoint list and aggregates all PDF documents containing 'Policy' or 'Procedure' from all other sites within the SharePoint domain.

.DESCRIPTION
    The script performs the following tasks:
    1. Connects to a SharePoint Online site using PnP PowerShell.
    2. Updates or deletes items in the specified SharePoint list based on their presence in search results.
    3. Searches for documents from sites in the SharePoint domain using a specified query.
    4. Adds new items to the list based on the search results.
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
    .\SharePoint Policy & Procedure Aggreator.ps1 -SiteUrl "https://example.sharepoint.com/sites/Policy" -ListName "Policies List"

    This example connects to the specified SharePoint Online site, updates or deletes items in the list based on the search results, searches for relevant documents, adds new items to the list, and displays the total document count.

.Link
    https://github.com/Rokawoo/sharepoint-policy-procedure-aggregator

.ONETIME-SETUP
    1. Run SharePoint Online Management Shell Installer: https://www.microsoft.com/en-US/download/details.aspx?id=35588&msockid=1873099af97a68ec13ce1d1ff8186956
    2. Run in PowerShell: Install-Module -Name SharePointPnPPowerShellOnline -Force -AllowClobber; $env:PNPLEGACYMESSAGE='false'
    3. Add the department site names exactly as they appear in the URL to 'Valid Departments.txt'—these will be used to populate the SharePoint lists.

.AUTHOR
    Roka Awoo
    8/26/2024
#>

param (
    [string]$SiteUrl = "https://example.sharepoint.com/sites/Policy",
    [string]$ListName = "Policies List"
)

function Initialize-DepartmentsHashSet {
    <#
    .SYNOPSIS
    Creates a HashSet from the lines of a .txt file named 'Valid Departments.txt' located in the same directory as the script.
    Each line is enclosed in '/' on both sides before being added to the HashSet.
    #>
    try {
        if (-not $PSScriptRoot) {
            throw "Could not determine the script directory."
        }

        $filePath = Join-Path -Path $PSScriptRoot -ChildPath "Valid Departments.txt"

        if (-not $filePath) {
            throw "Could not construct the file path."
        }

        $DepartmentsHashSet = [System.Collections.Generic.HashSet[string]]::new()

        if (Test-Path -Path $filePath) {
            Get-Content -Path $filePath | ForEach-Object {
                $enclosedValue = "/" + $_.Trim() + "/"
                $DepartmentsHashSet.Add($enclosedValue) | Out-Null
            }
        } else {
            throw "The file $filePath does not exist."
        }

        return $DepartmentsHashSet
    } catch {
        Write-Error "An error occurred: $_"
        exit 1
    }
}
$DepartmentsHashSet = Initialize-DepartmentsHashSet

function Write-Yellow {
    <#
    .SYNOPSIS
        Writes text to the console in yellow color.
    #>
    param (
        [string]$Text
    )
    Write-Host $Text -ForegroundColor Yellow
}

function Connect-ToSharePoint {
    <#
    .SYNOPSIS
        Connects to the specified SharePoint Online site using PnP PowerShell.
    #>
    try {
        Write-Yellow "Connecting to SharePoint Online..."
        Connect-PnPOnline -Url $SiteUrl -UseWebLogin
        Write-Yellow "Connected to SharePoint Online."
    }
    catch {
        Write-Error "Failed to connect to SharePoint: $_"
        throw
    }
}

function Search-Documents {
    <#
    .SYNOPSIS
        Searches for documents in SharePoint Online that match the criteria for policy and procedure documents.
    #>
    param (
        [string]$Query = '(contentclass:STS_ListItem_DocumentLibrary) AND
        (Title:Policy OR Title:Procedure) AND (FileExtension:pdf)'
    )

    try {
        Write-Host "Initiating search across all document libraries in the domain..." -ForegroundColor Yellow
        $results = Submit-PnPSearchQuery -Query $Query -TrimDuplicates $true -SelectProperties Title, Name, Category -All

        if ($results.ResultRows.Count -eq 0) {
            Write-Yellow "No documents found matching the criteria."
        }
        else {
            Write-Host "$($results.ResultRows.Count) documents found." -ForegroundColor Green
        }

        return $results.ResultRows
    }
    catch {
        Write-Error "Search query failed: $_"
        throw
    }
}

$dateFormat = "MM/dd/yyyy HH:mm:ss"
function Update-Or-AddItem {
    <#
    .SYNOPSIS
        Updates an existing item or adds a new item to the SharePoint list based on document details.
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
            $newModifiedDate = [DateTime]::ParseExact($LastModified, $dateFormat, $null)
            $existingModifiedDate = [DateTime]::ParseExact($existingItem.FieldValues["LastModified"], $dateFormat, $null)

            if ($newModifiedDate -gt $existingModifiedDate) {
                Write-Yellow "Updating existing item: $Title"
                Set-PnPListItem -List $ListName -Identity $existingItem.Id -Values @{
                    DocumentLink   = $DocumentLink
                    Category       = $DocumentCategory
                    Department     = $Department
                    LastModified   = $LastModified
                    DocumentAuthor = $DocumentAuthor
                }
            } else {
                Write-Yellow "Existing item is already up-to-date: $Title"
            }
        }
        else {
            Write-Yellow "Adding new item: $Title"
            Add-PnPListItem -List $ListName -Values @{
                Title          = $Title
                DocumentLink   = $DocumentLink
                Category       = $DocumentCategory
                Department     = $Department
                LastModified   = $LastModified
                DocumentAuthor = $DocumentAuthor
            }
        }
    }
    catch {
        Write-Error "Failed to update or add item: $_"
        throw
    }
}

function Remove-ObsoleteItems {
    <#
    .SYNOPSIS
        Removes items from a SharePoint list that are no longer present in the current list of items.
    #>
    param (
        [System.Collections.Generic.HashSet[string]]$CurrentItems
    )

    try {
        Write-Yellow "Removing obsolete items from the list..."
        $allItems = Get-PnPListItem -List $ListName

        $itemsToDelete = $allItems | Where-Object { -not $CurrentItems.Contains($_.FieldValues.Title) }

        if ($itemsToDelete.Count -eq 0) {
            Write-Yellow "No obsolete items found."
        } else {
            foreach ($item in $itemsToDelete) {
                Write-Yellow "Deleting obsolete item: $($item.FieldValues.Title)"
                Remove-PnPListItem -List $ListName -Identity $item.Id -Force
            }
            Write-Yellow "Obsolete items removed."
        }
    }
    catch {
        Write-Error "Failed to remove obsolete items: $_"
        throw
    }
}

function Get-DocumentCategory {
    <#
    .SYNOPSIS
        Determines the category of a document based on its title.
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
        [string]$Url,
        [System.Collections.Generic.HashSet[string]]$DepartmentsHashSet
    )

    if ($Url -match "/sites/" -and $Url -match "/Shared Documents/" -and $Url -notmatch "(?i)archive") {
        $containsDepartment = $DepartmentsHashSet | Where-Object { $Url -match $_ }

        if ($containsDepartment) {
            $slashCount = ($Url -split '/').Count - 1
            return $slashCount -le 6
        }
        else {
            return $false
        }
    }
    else {
        return $false
    }
}

$formattedResultCache = @{}
function AddSpaceBetweenCase {
    <#
    .SYNOPSIS
        Adds spaces between uppercase and lowercase letters in a string for better readability.
    #>
    param (
        [string]$inputString
    )

    if ($formattedResultCache.ContainsKey($inputString)) {
        return $formattedResultCache[$inputString]
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

    $result = $stringBuilder.ToString()
    $formattedResultCache[$inputString] = $result

    return $result
}

function Get-DepartmentFromUrl {
    <#
    .SYNOPSIS
        Extracts and formats the department name from a SharePoint document URL.
    #>
    param (
        [string]$Url
    )

    try {
        Write-Host "Processing URL: $Url"

        if (Check-UrlConditions -Url $Url -DepartmentsHashSet $DepartmentsHashSet) {
            $departmentPart = ($Url -split '/sites/')[1] -split '/' | Select-Object -First 1
            $formattedDepartment = AddSpaceBetweenCase -inputString $departmentPart

            Write-Host "Formatted Department: $formattedDepartment"
            return $formattedDepartment
        }

        Write-Warning "URL does not meet conditions. Returning 'Invalid'."
        return "Invalid"
    }
    catch {
        Write-Error "Failed to extract department from URL: $_"
        return "Invalid"
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

    Write-Yellow "Policies & Procedures List Last Aggregated: $((Get-PnPList -Identity $ListName).LastItemUserModifiedDate)"
    $searchResults = Search-Documents

    $successfulDocs = [System.Collections.Generic.HashSet[string]]::new()
    
    foreach ($result in $searchResults) {
        $docTitle = $result.Title
        $docCategory = Get-DocumentCategory -docTitle $docTitle
        $docUrl = $result.Path
        $docLastModified = $result.LastModifiedTime
        $docAuthor = Format-Authors -AuthorString $result.Author

        if ($docUrl -match "\.(doc|docx|pdf)$") {
            $department = Get-DepartmentFromUrl -Url $docUrl
            if ($department -ne "Invalid") {
                Update-Or-AddItem -Title $docTitle -DocumentLink $docUrl -DocumentCategory $docCategory -Department $department -LastModified $docLastModified -DocumentAuthor $docAuthor
                $successfulDocs.Add($docTitle) | Out-Null   
            }
            else {
                Write-Warning "Skipping document with invalid department: $docTitle"
            }

            Write-Host "---"
        }
    }

    Remove-ObsoleteItems -CurrentItems $successfulDocs
}
finally {
    Write-Yellow "Total Documents in List: $((Get-PnPList -Identity $ListName).ItemCount)"
    Disconnect-PnPOnline
    Write-Yellow "Disconnected from SharePoint Online."
}
