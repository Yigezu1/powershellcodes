<#
.NOTES
    Please update line number 6 to 8 before executing the script
#>

$newList      = 'Test List New'
$sourceSite   = 'https://tenant.sharepoint.com/sites/testsite'
$sourceFolder = "G:\Projs\yigezu_01"

# Define source file paths
$itemTypesFile       = "ItemTypes.csv"
$itemValuesFile      = "ItemValues.csv"
$lookupItemTypesFile = "LookupItemTypes.csv"
$sourceListIdFile    = 'SourceListId.txt'

# Define functions
Function Skip-ItemIds
{
    [CmdletBinding()]
    param (
        [Parameter()]$Counter
    )
    Write-Host "Waiting for ID slot" -ForegroundColor Cyan
    1..$Counter | ForEach-Object {
        $dummy = Add-PnPListItem -List $newList -Values @{Title = "Dummy"}
        Remove-PnPListItem -Identity $dummy -Force -List $newList
    }
}

Connect-PnPOnline -Url $sourceSite

# Read all files
$itemValues = Import-Csv -Path (Join-Path -Path $sourceFolder -ChildPath $itemValuesFile)
$newListDescription = Get-Content -Path (Join-Path -Path $sourceFolder -ChildPath $sourceListIdFile)
$itemTypes = @{}
Import-Csv -Path (Join-Path -Path $sourceFolder -ChildPath $itemTypesFile) | Foreach-Object {$itemTypes[$_.Key] = $_.Value}
$lookupItemTypes = Import-Csv -Path (Join-Path -Path $sourceFolder -ChildPath $lookupItemTypesFile)

# Create the destination list, if not already created
$newListObject = Get-PnPList -Identity $newList -ErrorAction Ignore

# Get all the lists and Create their ID and description hashtable
$allLists = @{}
Get-PnPList | ForEach-Object {$allLists[$_.Description] = $_.Id}

if ($newListObject)
{
    Write-Host "Destination list already exist. Removing it now"
    Get-PnPListItem -List $newList | Remove-PnPListItem -Force
    Remove-PnPList -Identity $newListObject -Force
}

Write-host "Creating the destination list" -ForegroundColor Cyan
$newPnpList = New-PnPList -Title $newList -Template GenericList

# Set the new list description
Set-PnPList -Identity $newList -Description $newListDescription | Out-Null

# Generate destination columns while skipping ID
$fieldsTobeCreated = $itemTypes.Keys.Where({$_ -ne "ID" -AND $_ -ne "Title"})
foreach ($field in $fieldsTobeCreated)
{
    Add-PnPField -List $newList -DisplayName $field -InternalName $field -Type $itemTypes[$field] -AddToDefaultView
}

# Set lookup fields
foreach ($lookupField in $lookupItemTypes)
{
    $lookupList = $allLists[$lookupField.LookupList]

    # Check if the list exist
    if ($lookupList)
    {
        Set-PnPField -List $newList -Identity $lookupField.FieldName -Values @{LookupList=$LookupList.ToString(); LookupField=$LookupField.LookupField}
    }else
    {
        Write-Warning "Lookup list for Lookup field $($lookupField.FieldName) does not exist"
    }
}

# Create Items
$fieldsToSet = $itemTypes.keys.Where({$_ -ne "ID"})
$nextId = 1
foreach ($item in $itemValues)
{
    $currentId = $item.ID
    if ($currentId -ne $nextId)
    {
        $Counter = $currentId - $nextId

        # Jump the IDs
        Skip-ItemIds -Counter $Counter
    }
    $listItem = @{}
    $fieldsToSet | Foreach-Object {
        if ($item.$_.Split(':')[0].ToLower() -eq "lookup".ToLower())
        {
            $listItem[$_] = $item.$_.Split(':')[1]
        }else
        {
            $listItem[$_] = $item.$_
        }
    }
    $listItemId = (Add-PnPListItem -List $newList -Values $listItem).Id
    $nextId = $listItemId + 1
}


