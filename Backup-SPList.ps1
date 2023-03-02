<#
.NOTES
    Please update line number 6 to 9 before executing the script
#>

$srcList = 'TestList1'
$sourceSite = 'https://sharepoint.yoururl/sites/contosoteam'
$outputFolder = "C:\Users\yigez\Documents\ListArchive"
$siteCols = @("Title","ID","DC","Owner","Users","WebUrl")

# Connect to Sharepoint OnPrem
Connect-PnPOnline -Url $sourceSite
$itemValues = @()
$itemTypes = @{}
$lookupItemTypes = @()

# Get all fields
$listFields = Get-PnPField -List $srcList

# Construct Item types
foreach ($col in $siteCols)
{
    $itemType = $listFields.Where({$_.InternalName -eq $col})
    if ($itemType.TypeAsString.Tolower() -eq "Lookup".ToLower())
    {
        $lookupObject = "" | Select-Object -property FieldName,LookupList,LookupField,DependentLookupName
        $lookupObject.FieldName = $col
        $lookupObject.LookupList = $itemType.LookupList
        $lookupObject.LookupField = $itemType.LookupField
        $lookupItemTypes += $lookupObject
    }else
    {
        $itemTypes[$col] = ($itemType | Select-Object -first 1).TypeAsString
    }
}

# Get all list items
$listItems = Get-PnPListItem -List $SrcList -Fields $siteCols

# Process each item
Foreach ($item in $listItems)
{
    $tempo = "" | Select-Object -Property $siteCols
    $siteCols | ForEach-Object {
        if ($item.FieldValues[$_].LookupId)
        {
            $tempo.$_ = "{0}:{1}" -f "lookup", $item.FieldValues[$_].LookupId
        }else
        {
            $tempo.$_ = $item.FieldValues[$_]
        }
    }
    $itemValues += $tempo
}

# Backup list id
$srcListId = (Get-PnPList -Identity $srcList).Id.Guid
$srcListId | Out-File -FilePath (Join-Path -Path $outputFolder -ChildPath 'SourceListId.txt')

# Backup Item Types
$itemTypes.GetEnumerator() | Export-Csv -Path (Join-Path -Path $outputFolder -ChildPath "ItemTypes.csv") -NoTypeInformation -Force

# Backup Item Values
$itemValues | Export-Csv -Path (Join-Path -Path $outputFolder -ChildPath "ItemValues.csv") -NoTypeInformation -Force

# Backup Lookup fields
$lookupItemTypes | Export-Csv -Path (Join-Path -Path $outputFolder -ChildPath "LookupItemTypes.csv") -NoTypeInformation -Force
