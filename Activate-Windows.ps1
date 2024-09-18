#Replace this value with your desired MAK key.
$Key_MAK="AAAAA-BBBBB-CCCCC-DDDDD-EEEEE"
#Leave this value blank, it will be set later.
$Activation_Key = ""

#Check if we're already activated and terminate if we don't need to activate.
$License = (Get-WmiObject -ClassName SoftwareLicensingProduct -filter 'PartialProductKey is not null' | Where-Object {$_.Description -like "*Windows*"} | Select-Object -First 1)
if ($License.LicenseStatus -eq 1) {
    $LicenseName = $License.Name
    $LicenseType = $License.ProductKeyChannel
    Write-Output "This host is already licensed with: $LicenseName ($LicenseType)"
    Exit 0
}

#Get the OA3Key Info from the system firmware
$OA3Info=(Get-WmiObject SoftwareLicensingService | Select-Object OA3x*)

Write-Output "Windows Activation Keys in Firmware:"
$OA3Info | Format-List

Write-Output " "
#Check if we can use the firmware key for activation
if (($OA3Info.OA3xOriginalProductKeyDescription.ToUpper() -like '*PROFESS*') -and ($OA3Info.OA3xOriginalProductKeyDescription.ToUpper()-like '*OEM*')) {
    Write-Host "Professional key found in firmware. Activating with Professional OEM key"
    $Activation_Key = $OA3Info.OA3xOriginalProductKey
} elseif (($OA3Info.OA3xOriginalProductKeyDescription.ToUpper() -like '*ENTERPR*') -and ($OA3Info.OA3xOriginalProductKeyDescription.ToUpper() -like '*OEM*')) {
    Write-Host "Enterprise key found in firmware. Activating with Enterprise OEM key"
    $Activation_Key = $OA3Info.OA3xOriginalProductKey
} elseif (($OA3Info.OA3xOriginalProductKeyDescription.ToUpper() -like '*EDU*') -and ($OA3Info.OA3xOriginalProductKeyDescription.ToUpper() -like '*OEM*')) {
    Write-Host "Education edition key found in firmware. Activating with Education OEM key"
    $Activation_Key = $OA3Info.OA3xOriginalProductKey
} else {
    Write-Host "No eligible Professional or higher SKU keys found in firmware. Activating with MAK key"
    $Activation_Key = $Key_MAK
}

#Remove existing product key
cscript //B //Nologo "$ENV:SYSTEMROOT\SYSTEM32\slmgr.vbs" /upk
#Install the desired key
cscript //B //Nologo "$ENV:SYSTEMROOT\SYSTEM32\slmgr.vbs" /ipk "$Activation_Key"
#Activate Windows
cscript //B //Nologo "$ENV:SYSTEMROOT\SYSTEM32\slmgr.vbs" /ato
