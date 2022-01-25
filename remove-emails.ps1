# Folder structure:
# Inbox
#   - Alerte
#       - Veeam
#           - Veeam1
#           - Veeam2
#           - Veeam3
#           - Veeam4


$olApp = new-object -comobject outlook.application
$namespace = $olApp.GetNamespace("MAPI")
$inbox = $namespace.GetDefaultFolder(6)

$alerte = $inbox.Folders | where { $_.name -eq "Alerte" }
$veeam = $alerte.Folders | where { $_.name -eq "Veeam"}
$date = (Get-Date).AddDays(-1)

# Delete from main folder
$veeamDelete = $veeam.Items | where -Property SentOn -LT $date
$numberToDelete = $veeamDelete.Count
if($numberToDelete){
    Write-Host "Deleting from folder Veeam $numberToDelete messages"
    $veeamDelete.Delete()
}

# Delete from subfolders
foreach ($veeamFolder in $veeam.Folders){
    $itemsToDelete = $veeamFolder.items | where -Property SentOn -LT $date
    $itemsToDelete.Count
    foreach ($item in $itemsToDelete){
        $item.Delete()
    }

}