Function prepareLabelSettings ($settings)
{
     $s=@{}
     foreach($item in $settings){
         $s[$item.key]=$item.value
     }
     Write-Host $s
     return $s
}
    
# 
Function processLabel ($l)
{
    New-Label -Name $l.label.name -DisplayName $l.label.displayName -comment $l.label.comment -ContentType $l.label.contentType -Tooltip $l.label.tooltip

    #Test if advanced settings are defined.
    if($prop = $l.label.PSObject.Properties['advancedsettings']) {
        $s=prepareLabelSettings $l.label.advancedsettings
        Set-Label -Identity $l.label.name -AdvancedSettings $s
    }

    #Test if content marking settings are defined.
    if($prop = $l.label.PSObject.Properties['contentmarking']) {
        $cm=$l.label.contentmarking

        #Test if watermark is defined.
        if($prop = $cm.PSObject.Properties['watermark']) {
            $w=$cm.watermark
            set-Label -Identity $l.label.name `
                -ApplyWaterMarkingEnabled $($w.ApplyWaterMarkingEnabled) `
                -ApplyWaterMarkingFontColor $($w.ApplyWaterMarkingFontColor) `
                -ApplyWaterMarkingFontName $($w.ApplyWaterMarkingFontName) `
                -ApplyWaterMarkingFontSize $($w.ApplyWaterMarkingFontSize) `
                -ApplyWaterMarkingLayout $($w.ApplyWaterMarkingLayout) `
                -ApplyWaterMarkingText $($w.ApplyWaterMarkingText) `
        }

        #Test if footer is defined.
        if($prop = $cm.PSObject.Properties['footer']) {
            $f=$cm.footer
            set-Label -Identity $l.label.name `
                -ApplyContentMarkingFooterAlignment $($f.ApplyContentMarkingFooterAlignment) `
                -ApplyContentMarkingFooterEnabled $($f.ApplyContentMarkingFooterEnabled) `
                -ApplyContentMarkingFooterFontColor $($f.ApplyContentMarkingFooterFontColor) `
                -ApplyContentMarkingFooterFontName $($f.ApplyContentMarkingFooterFontName) `
                -ApplyContentMarkingFooterFontSize $($f.ApplyContentMarkingFooterFontSize) `
                -ApplyContentMarkingFooterMargin $($f.ApplyContentMarkingFooterMargin) `
                -ApplyContentMarkingFooterText $($f.ApplyContentMarkingFooterText) `
        }

        #Test if header is defined.
        if($prop = $cm.PSObject.Properties['header']) {
            $h=$cm.header
            set-Label -Identity $l.label.name `
                -ApplyContentMarkingHeaderAlignment $($h.ApplyContentMarkingHeaderAlignment) `
                -ApplyContentMarkingHeaderEnabled $($h.ApplyContentMarkingHeaderEnabled) `
                -ApplyContentMarkingHeaderFontColor $($h.ApplyContentMarkingHeaderFontColor) `
                -ApplyContentMarkingHeaderFontName $($h.ApplyContentMarkingHeaderFontName) `
                -ApplyContentMarkingHeaderFontSize $($h.ApplyContentMarkingHeaderFontSize) `
                -ApplyContentMarkingHeaderMargin $($h.ApplyContentMarkingHeaderMargin) `
                -ApplyContentMarkingHeaderText $($h.ApplyContentMarkingHeaderText) `
        }
    }
}

Function deployLabels ($lp)
{
     New-LabelPolicy -name $lp.labelPolicy.name -comment $lp.labelPolicy.comment -Label $lp.labelPolicy.labels[0].name -ExchangeLocation $lp.labelPolicy.ExchangeLocation
}

$path = "C:\m365\repository\purview\information-protection\002_create_internal_label_with_content_marking\"

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
	$internalLabel = loadJSONFile("Label-internal.json")
    $internalLabelPolicy = loadJSONFile("LabelPolicy-internal.json")
    
    connect-IPPSSession -UserPrincipalName $upn

    processLabel($internalLabel)
    deployLabels ($internalLabelPolicy)

    Disconnect-ExchangeOnline -Confirm:$false

} catch {
	Write-Host "Error in the script!"  $_
}


