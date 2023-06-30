Function prepareLabelSettings ($settings)
{
     $s=@{}
     foreach($item in $settings){
         $s[$item.key]=$item.value
     }
     Write-Host $s
     return $s
}
    
$path = "C:\m365\repository\purview\information-protection\IaC_002_create_label_with_content_marking\"

try {
    $file_label = "Label-demo-2.json"
    Write-host "$path$file_label"
	$ConfigLabel = Get-Content "$path$file_label" | ConvertFrom-Json
	
    $file_labelPolicy = "LabelPolicy-demo-2.json"
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
    Write-Debug "watermark.ApplyWaterMarkingText: $($ConfigLabel.label.contentmarking.watermark.ApplyWaterMarkingText)"
    Write-Debug "header.ApplyContentMarkingHeaderText: $($ConfigLabel.label.contentmarking.header.ApplyContentMarkingHeaderText)"
    Write-Debug "footer.ApplyContentMarkingFooterText: $($ConfigLabel.label.contentmarking.footer.ApplyContentMarkingFooterText)"
    Write-Debug ""

	Write-Debug "Label Policy"
	Write-Debug "------------"
    Write-Debug "Name: $($ConfigLabelPolicy.labelPolicy.name)"
    Write-Debug "Comment: $($ConfigLabelPolicy.labelPolicy.comment)"
    Write-Debug "LabelS: $($ConfigLabelPolicy.labelPolicy.labels[0].name)"
       
    connect-IPPSSession -UserPrincipalName <upn>

    Write-Host "Create a new label: " $ConfigLabel.label.name
    New-Label -Name $ConfigLabel.label.name -DisplayName $ConfigLabel.label.displayName -comment $ConfigLabel.label.comment -ContentType $ConfigLabel.label.contentType -Tooltip $ConfigLabel.label.tooltip `
        -ApplyWaterMarkingEnabled $($ConfigLabel.label.contentmarking.watermark.ApplyWaterMarkingEnabled) `
        -ApplyWaterMarkingFontColor $($ConfigLabel.label.contentmarking.watermark.ApplyWaterMarkingFontColor) `
        -ApplyWaterMarkingFontName $($ConfigLabel.label.contentmarking.watermark.ApplyWaterMarkingFontName) `
        -ApplyWaterMarkingFontSize $($ConfigLabel.label.contentmarking.watermark.ApplyWaterMarkingFontSize) `
        -ApplyWaterMarkingLayout $($ConfigLabel.label.contentmarking.watermark.ApplyWaterMarkingLayout) `
        -ApplyWaterMarkingText $($ConfigLabel.label.contentmarking.watermark.ApplyWaterMarkingText) `
        -ApplyContentMarkingFooterAlignment $($ConfigLabel.label.contentmarking.footer.ApplyContentMarkingFooterAlignment) `
        -ApplyContentMarkingFooterEnabled $($ConfigLabel.label.contentmarking.footer.ApplyContentMarkingFooterEnabled) `
        -ApplyContentMarkingFooterFontColor $($ConfigLabel.label.contentmarking.footer.ApplyContentMarkingFooterFontColor) `
        -ApplyContentMarkingFooterFontName $($ConfigLabel.label.contentmarking.footer.ApplyContentMarkingFooterFontName) `
        -ApplyContentMarkingFooterFontSize $($ConfigLabel.label.contentmarking.footer.ApplyContentMarkingFooterFontSize) `
        -ApplyContentMarkingFooterMargin $($ConfigLabel.label.contentmarking.footer.ApplyContentMarkingFooterMargin) `
        -ApplyContentMarkingFooterText $($ConfigLabel.label.contentmarking.footer.ApplyContentMarkingFooterText) `
        -ApplyContentMarkingHeaderAlignment $($ConfigLabel.label.contentmarking.header.ApplyContentMarkingHeaderAlignment) `
        -ApplyContentMarkingHeaderEnabled $($ConfigLabel.label.contentmarking.header.ApplyContentMarkingHeaderEnabled) `
        -ApplyContentMarkingHeaderFontColor $($ConfigLabel.label.contentmarking.header.ApplyContentMarkingHeaderFontColor) `
        -ApplyContentMarkingHeaderFontName $($ConfigLabel.label.contentmarking.header.ApplyContentMarkingHeaderFontName) `
        -ApplyContentMarkingHeaderFontSize $($ConfigLabel.label.contentmarking.header.ApplyContentMarkingHeaderFontSize) `
        -ApplyContentMarkingHeaderMargin $($ConfigLabel.label.contentmarking.header.ApplyContentMarkingHeaderMargin) `
        -ApplyContentMarkingHeaderText $($ConfigLabel.label.contentmarking.header.ApplyContentMarkingHeaderText) `
            
    Write-Host "Update a new label: " $ConfigLabel.label.name
    $s=prepareLabelSettings $ConfigLabel.label.advancedsettings
    Set-Label -Identity $ConfigLabel.label.name -AdvancedSettings $s
            
    New-LabelPolicy -name $ConfigLabelPolicy.labelPolicy.name -comment $ConfigLabelPolicy.labelPolicy.comment -Label $ConfigLabelPolicy.labelPolicy.labels[0].name -ExchangeLocation $ConfigLabelPolicy.labelPolicy.ExchangeLocation

    Disconnect-ExchangeOnline -Confirm:$false
} catch {
	Write-Host "Error in the script!"  $_
}


