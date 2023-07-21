Function prepareLabelSettings ($settings)
{
     $s=@{}
     foreach($item in $settings){
         $s[$item.key]=$item.value
     }
     Write-Host $s
     return $s
}
    
Function processLabel ($l)
{
    #Create a new label. 
    New-Label -Name $l.label.name -DisplayName $l.label.displayName -comment $l.label.comment -ContentType $l.label.contentType -Tooltip $l.label.tooltip

    #Test if advanced settings are defined.
    if($prop = $l.label.PSObject.Properties['advancedsettings']) {
        $s=prepareLabelSettings $l.label.advancedsettings
        Set-Label -Identity $l.label.name -AdvancedSettings $s
    }
}


Function deployLabels ($lp)
{
     New-LabelPolicy -name $lp.labelPolicy.name -comment $lp.labelPolicy.comment -Label $lp.labelPolicy.labels[0].name -ExchangeLocation $lp.labelPolicy.ExchangeLocation
}

$path = "C:\m365\repository\purview\information-protection\001_create_and_deploy_public_label\"

Function loadJSONFile($name)
{
    Write-host "$path$name"
    return Get-Content "$path$name" | ConvertFrom-Json
}

try {
    #Possible values: SilentlyContinue, Continue
    $DebugPreference = "Continue"
    $upn="<upn>"

    #load config files.
	$publicLabel = loadJSONFile("Label-public.json")
    $publicLabelPolicy = loadJSONFile("LabelPolicy-public.json")
    
    connect-IPPSSession -UserPrincipalName $upn

    processLabel($publicLabel)
    deployLabels ($publicLabelPolicy)

    Disconnect-ExchangeOnline -Confirm:$false
} catch {
	Write-Host "Error in the script!"  $_
}


