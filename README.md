## Setup

### 1. **Run SharePoint Online Management Shell Installer**

- **Link**: [SharePoint Online Management Shell Installer](https://www.microsoft.com/en-US/download/details.aspx?id=35588&msockid=1873099af97a68ec13ce1d1ff8186956)
  
  **Purpose**: This installer sets up the SharePoint Online Management Shell, which is a PowerShell module specifically designed for managing SharePoint Online.

  **Action**:
  - Download the installer from the link provided.
  - Run the installer to set up the SharePoint Online Management Shell on your computer.
  - This provides the necessary cmdlets and tools for managing SharePoint Online via PowerShell.

### 2. **Install PnP PowerShell Module**

- **Command**:

  ```powershell
  Install-Module -Name SharePointPnPPowerShellOnline -Force -AllowClobber; $env:PNPLEGACYMESSAGE='false'
  ```

  **Parameters**:

  - **`Install-Module`**: This cmdlet installs a module from the PowerShell Gallery.
    - **`-Name SharePointPnPPowerShellOnline`**: Specifies the module name to be installed. `SharePointPnPPowerShellOnline` is a module provided by the SharePoint Patterns and Practices (PnP) team to simplify managing SharePoint Online.
    - **`-Force`**: Forces the installation of the module, even if it’s already installed. This parameter can be used to override existing installations or updates.
    - **`-AllowClobber`**: Allows the cmdlets in the module to overwrite existing cmdlets with the same names if there are any conflicts.

  - **`$env:PNPLEGACYMESSAGE='false'`**: This command sets an environment variable that disables legacy messages from the PnP PowerShell module. Legacy messages are warnings related to deprecated features or practices. Setting this variable to `'false'` suppresses these messages, providing a cleaner output.

### 3. **Set Up a SharePoint List**

- **Instructions**:
  1. **Go to Your SharePoint Site**: Navigate to your SharePoint Online site where you want to create the list.
  2. **Create a New List**: 
     - Click on **Site Contents**.
     - Select **New** and then **List**.
     - Give your list a name and create it.

  **Fields to Add**:
  
  - **`Document Type`** (Single Line Text): A text field to specify the type of document (e.g., Policy, Procedure).
    - **Field Type**: Single Line Text
    - **Purpose**: To store the document type as a short text entry.

  - **`Title`** (Single Line Text): The title or name of the document.
    - **Field Type**: Single Line Text
    - **Purpose**: To provide a descriptive name or title for the document.

  - **`Department`** (Single Line Text): A text field to indicate the department associated with the document.
    - **Field Type**: Single Line Text
    - **Purpose**: To categorize the document by department.

  - **`Last Modified`** (Date Time): A date and time field to record when the document was last modified.
    - **Field Type**: Date Time
    - **Purpose**: To keep track of the last modification date and time of the document.

  - **`Document Author`** (Single Line Text): A text field for the name of the document's author.
    - **Field Type**: Single Line Text
    - **Purpose**: To specify who created or authored the document.

## Example Usage
```ps
.\SharePoint Policy & Procedure Aggreator.ps1 -SiteUrl "https://example.sharepoint.com/sites/Policy" -ListName "Policies List"
```
