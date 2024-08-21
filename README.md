## Setup

### 1. **Run and Complete SharePoint Online Management Shell Installer**

- **Link**: [SharePoint Online Management Shell Installer](https://www.microsoft.com/en-US/download/details.aspx?id=35588&msockid=1873099af97a68ec13ce1d1ff8186956)
  
### 2. **Install PnP PowerShell Module**

- **Command**:

  ```powershell
  Install-Module -Name SharePointPnPPowerShellOnline -Force -AllowClobber; $env:PNPLEGACYMESSAGE='false'
  ```

### 3. **Set Up a SharePoint List**

- **Instructions**:
  1. **Go to Your SharePoint Site**: Navigate to your SharePoint Online site where you want to create the list.
  2. **Create a New List**: 
     - Click on **Site Contents**.
     - Select **New** and then **List**.
     - Give your list a name and create it.

- **Fields to Add**:
  - **`Document Type`** (Single Line Text): A text field to specify the type of document (e.g., Policy, Procedure).
    - **Field Type**: Single Line Text

  - **`Title`** (Single Line Text): The title or name of the document.
    - **Field Type**: Single Line Text

  - **`Department`** (Single Line Text): A text field to indicate the department associated with the document.
    - **Field Type**: Single Line Text

  - **`Last Modified`** (Date Time): A date and time field to record when the document was last modified.
    - **Field Type**: Date Time

  - **`Document Author`** (Single Line Text): A text field for the name of the document's author.
    - **Field Type**: Single Line Text

## Example Usage
```ps
.\SharePoint Policy & Procedure Aggreator.ps1 -SiteUrl "https://example.sharepoint.com/sites/Policy" -ListName "Policies List"
```
