Function deleteAllLabelPolicies 
{
    $labelPolicies=get-labelPolicy

    $labelPolicies.ForEach( {
       Write-Host $_.Name 
       remove-labelPolicy -identity $_.Name -Confirm:$false
    })
    return $flag
}

Function deleteAllChildLabels
{
    $labels=get-label | where parentid -ne $null

    $labels.ForEach( {
       Write-Host $_.Name 
       remove-label -identity $_.Name -Confirm:$false
    })
    return $flag
}

Function deleteAllParentLabels 
{
    $labels=get-label | where parentid -eq $null

    $labels.ForEach( {
       Write-Host $_.Name 
       remove-label -identity $_.Name -Confirm:$false
    })
    return $flag
}

$upn=$Env:DIVE_INTO_M365_UPN

connect-IPPSSession -UserPrincipalName $upn
deleteAllLabelPolicies 

Start-Sleep -Seconds 30
deleteAllChildLabels

Start-Sleep -Seconds 30
deleteAllParentLabels

get-labelPolicy | select name,mode
get-label | select name,mode

Disconnect-ExchangeOnline -Confirm:$false

