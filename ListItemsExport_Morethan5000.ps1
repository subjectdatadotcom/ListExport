$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Load SharePoint Online Client Assembly
Add-Type -Path "$myDir\CSOMLibs\Microsoft.SharePoint.Client.dll"
Add-Type -Path "$myDir\CSOMLibs\Microsoft.SharePoint.Client.Runtime.dll"

# Function to authenticate and get context
Function Get-SharePointContext {
    param (
        [string]$SiteUrl,
        [string]$Username,
        [string]$Password
    )

    $SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username, $SecurePassword)
    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
    $Context.Credentials = $Credentials
    return $Context
}

# Function to export list items to CSV
Function Export-ListItemsToCSV {
    param (
        [Microsoft.SharePoint.Client.ClientContext]$Context,
        [string]$ListName,
        [string]$OutputCsvPath
    )

    # Get the list
    $List = $Context.Web.Lists.GetByTitle($ListName)
    $Context.Load($List.Fields)
    $Context.ExecuteQuery()

    # Initialize CAML query for pagination
    $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
    $Query.ViewXml = "<View><RowLimit>5000</RowLimit></View>"
    $Query.ListItemCollectionPosition = $null

    # Create an array to hold the exported data
    $ExportData = @()

    # Loop through items in batches
    do {
        # Fetch list items in the current batch
        $ListItems = $List.GetItems($Query)
        $Context.Load($ListItems)
        $Context.ExecuteQuery()

        foreach ($Item in $ListItems) {
            $ItemData = @{}

            # Loop through all fields and add them to the item data
            foreach ($Field in $List.Fields) {
                try {
                    $FieldValue = $Item[$Field.InternalName]
                    $ItemData[$Field.Title] = $FieldValue -join ";" # Handle multi-value fields
                } catch {
                    # Ignore fields that cannot be read
                    $ItemData[$Field.Title] = ""
                }
            }

            # Add the item data to the export data
            $ExportData += New-Object PSObject -Property $ItemData
        }

        # Update the ListItemCollectionPosition to the next batch
        $Query.ListItemCollectionPosition = $ListItems.ListItemCollectionPosition

    } while ($Query.ListItemCollectionPosition -ne $null)

    # Export the data to CSV
    $ExportData | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Exported all list items to $OutputCsvPath" -ForegroundColor Green
}

# Variables
$SiteUrl = "https://smartpro2"  # Replace with your site URL
$Username = ""               # Replace with your admin username
$Password = ""                                   # Replace with your admin password
$ListName = "RTRY - Eastern Search" #"Registry - QC - Variant" # Replace with your list name
$OutputCsvPath = "$myDir\$($ListName).csv"                  # Replace with your desired CSV output file path

# Get SharePoint context
$Context = Get-SharePointContext -SiteUrl $SiteUrl -Username $Username -Password $Password

# Export list items to CSV
Export-ListItemsToCSV -Context $Context -ListName $ListName -OutputCsvPath $OutputCsvPath
