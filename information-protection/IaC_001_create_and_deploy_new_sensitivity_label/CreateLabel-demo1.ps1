Function prepareLabelSettings ($settings)
{
     $s=@{}
     foreach($item in $settings){
         $s[$item.key]=$item.value
     }
     Write-Host $s
     return $s
}
    
$path = "C:\m365\repository\purview\information-protection\IaC_001_create_and_deploy_new_sensitivity_label\article001_IaC_create_and_deploy_new_label\"

try {
    $file_label = "Label-demo-1.json"
	$ConfigLabel = Get-Content "$path$file_label" | ConvertFrom-Json
	
    $file_labelPolicy = "LabelPolicy-demo-1.json"
    $ConfigLabelPolicy = Get-Content "$path$file_labelPolicy" | ConvertFrom-Json

    #Possible values: SilentlyContinue, Continue
    $DebugPreference = "Continue"

    Write-Debug "Label"
	Write-Debug "-----"
	Write-Debug "Name: $($ConfigLabel.label.name)"
    Write-Debug "DisplayName: $($ConfigLabel.label.displayName)"
    Write-Debug "Tooltip: $($ConfigLabel.label.tooltip)"
    Write-Debug "Comment: $($ConfigLabel.label.comment)"
    Write-Debug "ContentType: $($ConfigLabel.label.contentType)"
    Write-Debug "Settings Key-1 : $($ConfigLabel.label.advancedsettings[0].key)"
    Write-Debug "Settings Value-1: $($ConfigLabel.label.advancedsettings[0].value)"
    Write-Debug ""

	Write-Debug "Label Policy"
	Write-Debug "------------"
    Write-Debug "Name: $($ConfigLabelPolicy.labelPolicy.name)"
    Write-Debug "Comment: $($ConfigLabelPolicy.labelPolicy.comment)"
    Write-Debug "LabelS: $($ConfigLabelPolicy.labelPolicy.labels[0].name)"
        
    connect-IPPSSession -UserPrincipalName <upn>

    Write-Host "Create a new label: " $ConfigLabel.label.name
    New-Label -Name $ConfigLabel.label.name -DisplayName $ConfigLabel.label.displayName -comment $ConfigLabel.label.comment -ContentType $ConfigLabel.label.contentType -Tooltip $ConfigLabel.label.tooltip 

    $s=prepareLabelSettings $ConfigLabel.label.advancedsettings

    Write-Host "Update a new label: " $ConfigLabel.label.name
    Set-Label -Identity $ConfigLabel.label.name -AdvancedSettings $s

    New-LabelPolicy -name $ConfigLabelPolicy.labelPolicy.name -comment $ConfigLabelPolicy.labelPolicy.comment -Label $ConfigLabelPolicy.labelPolicy.labels[0].name -ExchangeLocation $ConfigLabelPolicy.labelPolicy.ExchangeLocation

    Disconnect-ExchangeOnline -Confirm:$false
} catch {
	Write-Host "Error in the script!"  $_
}


