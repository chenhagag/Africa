Connect-PnPOnline -Url "https://africaisrael.sharepoint.com/ContractsNEW" -UseWebLogin
$items = Get-PnPListItem -List "ContractsDocs" -Fields "FileLeafRef","SiteName","ContractNumber" -PageSize 500

$results = @()
foreach ($item in $items) {
    $site = "$($item.FieldValues['SiteName'])".Trim()
    $fn = "$($item.FieldValues['FileLeafRef'])"
    $cn = "$($item.FieldValues['ContractNumber'])"
    $results += [PSCustomObject]@{
        SiteName = $site
        FileName = $fn
        ContractNumber = $cn
    }
}

$results | Export-Csv -Path "$PSScriptRoot\new_library_items.csv" -NoTypeInformation -Encoding UTF8
Write-Host "Exported $($results.Count) items to new_library_items.csv"

Disconnect-PnPOnline
