<div align="center">
  <img src="https://github.com/user-attachments/assets/63b26005-7a51-4b5f-a142-00d56397cfa4" alt="Aggregator" align="center" width="235px"/>
  <h1>SharePoint Policy & Procedure Aggregator</h1>
  <p>By Rokawoo</p>
</div>

> [!CAUTION]
> ⭐ This script is superduper cool!! :3

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

  - **`Category`** (Single Line Text): A text field to specify the category of document (e.g., Policy, Procedure).

    - **Field Type**: Single Line Text

  - **`Title`** (Single Line Text): The title or name of the document.

    - **Field Type**: Single Line Text

    - **Format Json**: 
    ```json
    {
      "$schema": "https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json",
      "elmType": "a",
      "attributes": {
        "href": "[$DocumentLink]",
        "target": "_blank"
      },
      "style": {
        "color": "black",
        "text-decoration": "underline",
        "font-size": "14px",
        "font-weight": "600"
      },
      "txtContent": "@currentField"
    }
    ```

  - **`Department`** (Single Line Text): A text field to indicate the department associated with the document.

    - **Field Type**: Single Line Text

  - **`Last Modified`** (Date Time): A date and time field to record when the document was last modified.

    - **Field Type**: Date Time

  - **`Document Author`** (Single Line Text): A text field for the name of the document's author.

    - **Field Type**: Single Line Text

  - **`Document Link`** (Single Line Text): A text field for the name of the document's author.

    - **Field Type**: Multiple Line Text

    - **Format Json**: 
    ```json
    {
      "$schema": "https://developer.microsoft.com/json-schemas/sp/column-formatting.schema.json",
      "elmType": "a",
      "attributes": {
        "href": "=if(@currentField, @currentField, '#')",
        "target": "_blank"
      },
      "style": {
        "text-decoration": "none",
        "color": "=if(@currentField, '#0078d4', 'red')",
        "font-weight": "=if(@currentField, 'normal', 'bold')"
      },
      "txtContent": "=if(@currentField, 'Link', 'URL Error')"
    }
    ```

- **Formatting List View**:
  - Group the list by **`Department`**
  - Sort the list alphabetically by **`Title`**
  - **Format Json**:
  ```json
  {
    "$schema": "https://developer.microsoft.com/json-schemas/sp/v2/row-formatting.schema.json",
    "groupProps": {
      "headerFormatter": {
        "elmType": "div",
        "style": {
          "display": "flex",
          "align-items": "center",
          "width": "48vw",
          "height": "3vh",
          "padding": "8px 10px",
          "background-color": "#DEDEDE",
          "color": "#000000",
          "border": "1px solid #ABABAB",
          "border-radius": "5px"
        },
        "children": [
          {
            "elmType": "span",
            "txtContent": "=@group.fieldData",
            "style": {
              "font-weight": "550",
              "font-size": "16px"
            }
          },
          {
            "elmType": "span",
            "txtContent": "='(' + @group.count + ')'",
            "style": {
              "margin-left": "5px",
              "font-weight": "550",
              "font-size": "16px"
            }
          }
        ]
      }
    }
  }

### 4. **Set Up the Valid Departments List**
> [!Note]
> The sites from which the `.\SharePoint Policy & Procedure Aggregator.ps1` script will pull documents must be configured.

1. **Locate the File**:
   - Find the file named `Valid Departments.txt`. This file should be located in the same directory as your PowerShell script, `.\SharePoint Policy & Procedure Aggregator.ps1`.

2. **Enter Department Site Names**:
   - In the `Valid Departments.txt` file, list each department's site name exactly as it appears in the SharePoint URLs. Each department name should be on its own line. Ensure there are no extra spaces or incorrect characters.
   
   For example, if your SharePoint URLs are like:
   - `https://example.sharepoint.com/sites/HumanResources`
   - `https://example.sharepoint.com/sites/Finance`

   Then, you should enter:
   ```
   HumanResources
   Finance
   ```

3. **Save the File**:
- After entering all the department names, save the `Valid Departments.txt` file.
---
## Example Usage

**Note:** The script defaults to the Rhoads SharePoint Domain and the "Polices & Procedures by Department" List.
```ps
.\SharePoint Policy & Procedure Aggreator.ps1
```

**With Custom Parameters:**
```ps
.\SharePoint Policy & Procedure Aggreator.ps1 -SiteUrl "https://example.sharepoint.com/sites/Policy" -ListName "Policies List"
```
---

