#Check if the label already exists
Function checkIfLabelExist ($labelName)
{
    $labels=get-label
    $flag=$false

    $Labels.ForEach( {
       Write-Host $_.Name 
       if( $labelName.equals($_.Name)){
            Write-Host $_.Name "=" $labelName
            $flag=$true
        }
    })
    return $flag
}

#
Function prepareLabelSettings ($settings)
{
     $s=@{}
     foreach($item in $settings){
         $s[$item.key]=$item.value
     }
     return $s
}

# 
Function processLabel ($l)
{
    
    $existingLabelFlag=checkIfLabelExist ($l.label.name)

    #Create a new label. 
    if($existingLabelFlag -eq $false){
        New-Label -Name $l.label.name -DisplayName $l.label.displayName -comment $l.label.comment -ContentType $l.label.contentType -Tooltip $l.label.tooltip
    }

    #Test if encryption settings are defined.
    if($prop = $l.label.PSObject.Properties['encryption']) {
        $e=$l.label.encryption

        #Test if Template encryption settings are defined.
        if ($e.EncryptionProtectionType -eq "Template"){
            try {
                $adp=$e.adminDefinedPermissions
                set-Label -Identity $l.label.name `
                    -EncryptionEnabled $($e.EncryptionEnabled)`
                    -EncryptionProtectionType $($e.EncryptionProtectionType)`
                    -EncryptionContentExpiredOnDateInDaysOrNever $($adp.EncryptionContentExpiredOnDateInDaysOrNever)`
                    -EncryptionOfflineAccessDays $($adp.EncryptionOfflineAccessDays)`
                    -EncryptionRightsDefinitions $($adp.EncryptionRightsDefinitions)`
            } catch {
	            Write-Host "Template Error..."  $_
            }
        }

        #Test if User Defined encryption settings are defined.
        if ($e.EncryptionProtectionType -eq "UserDefined"){
            try {
                $udp=$e.userDefinedPermissions
                set-Label -Identity $l.label.name `
                    -EncryptionEnabled $($e.EncryptionEnabled)`
                    -EncryptionProtectionType $($e.EncryptionProtectionType)`
                    -EncryptionPromptUser $($udp.EncryptionPromptUser)`
                    -EncryptionDoNotForward $($udp.EncryptionDoNotForward)`
            } catch {
	            Write-Host "UserDefined Error..."  $_
            }
        }
    }

    #Test if advanced settings are defined.
    if($prop = $l.label.PSObject.Properties['advancedsettings']) {
        $s=prepareLabelSettings $l.label.advancedsettings
        Set-Label -Identity $l.label.name -AdvancedSettings $s
    }

    #Test if content marking settings are defined.
    if($prop = $l.label.PSObject.Properties['contentmarking']) {
        try {
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
        } catch {
            Write-Host "UserDefined Error..."  $_
        }

    }
    if($prop = $l.label.PSObject.Properties['parentLabel']) {
        set-Label -Identity $l.label.name -ParentId $($l.label.parentLabel)
    }

}

Function deployLabels ($lp)
{
     $s = ""
     $labels = foreach ($item in $lp.labelPolicy.labels) {
        $item.name
     }
     $s = $labels -join ","

     $multiValuedProperty = New-Object System.Collections.ArrayList

     foreach ($item in $lp.labelPolicy.labels) {
        $multiValuedProperty.Add($item.name) | Out-Null
     }

     New-LabelPolicy -name $lp.labelPolicy.name -comment $lp.labelPolicy.comment -Label $multiValuedProperty -ExchangeLocation $ConfigLabelPolicy.labelPolicy.ExchangeLocation
}

$path = "C:\m365\repository\purview\information-protection\003_create_confidenital_labels_with_encryption\"

Function loadJSONFile($name)
{
    Write-host "$path$name"
    return Get-Content "$path$name" | ConvertFrom-Json
}

try {
    #Possible values: SilentlyContinue, Continue
    $DebugPreference = "Continue"
    $upn=$Env:DIVE_INTO_M365_UPN

    #load config files.
	$confidentialLabel = loadJSONFile("Label-confidential.json")
	$confidentialAllEmployees = loadJSONFile("Label-confidential-allEmployee.json")
	$confidentialTrustedPeople = loadJSONFile("Label-confidential-trustedPeople.json")

    $ConfigLabelPolicy = loadJSONFile("LabelPolicy-confidential.json")

    connect-IPPSSession -UserPrincipalName $upn

    #process config files.
    processLabel($confidentialLabel)
    processLabel($confidentialAllEmployees)
    processLabel($confidentialTrustedPeople)
    deployLabels ($ConfigLabelPolicy)

    Disconnect-ExchangeOnline -Confirm:$false
} catch {
	Write-Host "Error in the script!"  $_
}


