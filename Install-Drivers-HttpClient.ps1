#Large portions of this script have been copied from BuildFFUVM.ps1 from https://aka.ms/ffu

function WriteLog($LogText) { 
    Write-Output "$((Get-Date).ToString()) $LogText" -Force -ErrorAction SilentlyContinue
    Write-Verbose $LogText
}

function Test-Url {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Url
    )
    try {
        # Create a web request and check the response
        $request = [System.Net.WebRequest]::Create($Url)
        $request.Method = 'HEAD'
        $response = $request.GetResponse()
        return $true
    }
    catch {
        return $false
    }
}

function Download-File
{
  
}

# Function to download a file using BITS with retry and error handling
function Start-BitsTransferWithRetry {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Source,
        [Parameter(Mandatory = $true)]
        [string]$Destination,
        [int]$Retries = 3
    )

    $attempt = 0
    [System.Net.WebClient]$webClient = New-Object -TypeName "System.Net.WebClient"
    while ($attempt -lt $Retries) {
        try {
            $OriginalVerbosePreference = $VerbosePreference
            $VerbosePreference = 'SilentlyContinue'
            $ProgressPreference = 'SilentlyContinue'
            #Start-BitsTransfer -Source $Source -Destination $Destination -ErrorAction Stop
            $webClient.DownloadFile($Source,$Destination)
            $ProgressPreference = 'Continue'
            $VerbosePreference = $OriginalVerbosePreference
            return
        }
        catch {
            $attempt++
            WriteLog "Attempt $attempt of $Retries failed to download $Source. Retrying..."
            Start-Sleep -Seconds 5
        }
    }
    WriteLog "Failed to download $Source after $Retries attempts."
    return $false
}

function Invoke-Process {
    [CmdletBinding(SupportsShouldProcess)]
    param
    (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$FilePath,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$ArgumentList
    )

    $ErrorActionPreference = 'Stop'

    try {
        $stdOutTempFile = "$env:TEMP\$((New-Guid).Guid)"
        $stdErrTempFile = "$env:TEMP\$((New-Guid).Guid)"

        $startProcessParams = @{
            FilePath               = $FilePath
            ArgumentList           = $ArgumentList
            RedirectStandardError  = $stdErrTempFile
            RedirectStandardOutput = $stdOutTempFile
            Wait                   = $true;
            PassThru               = $true;
            NoNewWindow            = $true;
        }
        if ($PSCmdlet.ShouldProcess("Process [$($FilePath)]", "Run with args: [$($ArgumentList)]")) {
            $cmd = Start-Process @startProcessParams
            $cmdOutput = Get-Content -Path $stdOutTempFile -Raw
            $cmdError = Get-Content -Path $stdErrTempFile -Raw
            if ($cmd.ExitCode -ne 0) {
                if ($cmdError) {
                    throw $cmdError.Trim()
                }
                if ($cmdOutput) {
                    throw $cmdOutput.Trim()
                }
            }
            else {
                if ([string]::IsNullOrEmpty($cmdOutput) -eq $false) {
                    WriteLog $cmdOutput
                }
            }
        }
    }
    catch {
        #$PSCmdlet.ThrowTerminatingError($_)
        WriteLog $_
        Write-Host "Script failed - $Logfile for more info"
        throw $_

    }
    finally {
        Remove-Item -Path $stdOutTempFile, $stdErrTempFile -Force -ErrorAction Ignore
		
    }
	
}


#### MICROSOFT DRIVERS ########################################################
function Get-MicrosoftDrivers {
    param (
        [string]$Make,
        [string]$Model,
        [int]$WindowsRelease
    )

    $url = "https://support.microsoft.com/en-us/surface/download-drivers-and-firmware-for-surface-09bb2e09-2a4b-cb69-0951-078a7739e120"
    
    # Download the webpage content
    WriteLog "Getting Surface driver information from $url"
    $OriginalVerbosePreference = $VerbosePreference
    $VerbosePreference = 'SilentlyContinue'
    $webContent = Invoke-WebRequest -Uri $url -UseBasicParsing -Headers $Headers -UserAgent $UserAgent
    $VerbosePreference = $OriginalVerbosePreference
    WriteLog "Complete"

    # Parse the content of the relevant nested divs
    WriteLog "Parsing web content for models and download links"
    $html = $webContent.Content
    $nestedDivPattern = '<div id="ID0EBHFBH[1-8]" class="ocContentControlledByDropdown.*?">(.*?)</div>'
    $nestedDivMatches = [regex]::Matches($html, $nestedDivPattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)

    $models = @()
    $modelPattern = '<p>(.*?)</p>\s*</td>\s*<td>\s*<p>\s*<a href="(.*?)"'

    foreach ($nestedDiv in $nestedDivMatches) {
        $nestedDivContent = $nestedDiv.Groups[1].Value
        $modelMatches = [regex]::Matches($nestedDivContent, $modelPattern)

        foreach ($match in $modelMatches) {
            $modelName = $match.Groups[1].Value
            $modelLink = $match.Groups[2].Value
            $models += [PSCustomObject]@{ Model = $modelName; Link = $modelLink }
        }
    }
    WriteLog "Parsing complete"

    # Validate the model
    $selectedModel = $models | Where-Object { $_.Model -eq $Model }

    if ($null -eq $selectedModel) {
        if ($VerbosePreference -ne 'Continue') {
            Write-Host "The model '$Model' was not found in the list of available models."
            Write-Host "Please run the script with the -Verbose switch to see the list of available models."
        }
        WriteLog "The model '$Model' was not found in the list of available models."
        WriteLog "Please select a model from the list below by number:"
        
        for ($i = 1; $i -lt $models.Count; $i++) {
            if ($VerbosePreference -ne 'Continue') {
                Write-Host "$i. $($models[$i].Model)"
            }
            WriteLog "$i. $($models[$i].Model)"
        }

        do {
            $selection = Read-Host "Enter the number of the model you want to select"
            WriteLog "User selected model number: $selection"
            
            if ($selection -match '^\d+$' -and [int]$selection -ge 0 -and [int]$selection -lt $models.Count) {
                $selectedModel = $models[$selection]
            } else {
                if ($VerbosePreference -ne 'Continue') {
                    Write-Host "Invalid selection. Please try again."
                }
                WriteLog "Invalid selection. Please try again."
            }
        } while ($null -eq $selectedModel)
    }

    $Model = $selectedModel.Model
    WriteLog "Model: $Model"
    WriteLog "Download Page: $($selectedModel.Link)"

    # Follow the link to the download page and parse the script tag
    WriteLog "Getting download page content"
    $OriginalVerbosePreference = $VerbosePreference
    $VerbosePreference = 'SilentlyContinue'
    $downloadPageContent = Invoke-WebRequest -Uri $selectedModel.Link -UseBasicParsing -Headers $Headers -UserAgent $UserAgent
    $VerbosePreference = $OriginalVerbosePreference
    WriteLog "Complete"
    WriteLog "Parsing download page for file"
    $scriptPattern = '<script>window.__DLCDetails__={(.*?)}</script>'
    $scriptMatch = [regex]::Match($downloadPageContent.Content, $scriptPattern)

    if ($scriptMatch.Success) {
        $scriptContent = $scriptMatch.Groups[1].Value

        # Extract the download file information from the script tag
        $downloadFilePattern = '"name":"(.*?)",.*?"url":"(.*?)"'
        $downloadFileMatches = [regex]::Matches($scriptContent, $downloadFilePattern)

        $downloadLink = $null
        foreach ($downloadFile in $downloadFileMatches) {
            $fileName = $downloadFile.Groups[1].Value
            $fileUrl = $downloadFile.Groups[2].Value

            if ($fileName -match "Win$WindowsRelease") {
                $downloadLink = $fileUrl
                break
            }
        }

        if ($downloadLink) {
            WriteLog "Download Link for Windows ${WindowsRelease}: $downloadLink"
        
            # Create directory structure
            if (-not (Test-Path -Path $DriversFolder)) {
                WriteLog "Creating Drivers folder: $DriversFolder"
                New-Item -Path $DriversFolder -ItemType Directory -Force | Out-Null
                WriteLog "Drivers folder created"
            }
            $surfaceDriversPath = Join-Path -Path $DriversFolder -ChildPath $Make
            $modelPath = Join-Path -Path $surfaceDriversPath -ChildPath $Model
            if (-Not (Test-Path -Path $modelPath)) {
                WriteLog "Creating model folder: $modelPath"
                New-Item -Path $modelPath -ItemType Directory | Out-Null
                WriteLog "Complete"
            }
        
            # Download the file
            $filePath = Join-Path -Path $surfaceDriversPath -ChildPath ($fileName)
            WriteLog "Downloading $Model driver file to $filePath"
            Start-BitsTransferWithRetry -Source $downloadLink -Destination $filePath
            WriteLog "Download complete"
        
            # Determine file extension
            $fileExtension = [System.IO.Path]::GetExtension($filePath).ToLower()
        
            if ($fileExtension -eq ".msi") {
                # Extract the MSI file using an administrative install
                WriteLog "Extracting MSI file to $modelPath"
                $arguments = "/a `"$($filePath)`" /qn TARGETDIR=`"$($modelPath)`""
                Invoke-Process -FilePath "msiexec.exe" -ArgumentList $arguments
                WriteLog "Extraction complete"
            } elseif ($fileExtension -eq ".zip") {
                # Extract the ZIP file
                WriteLog "Extracting ZIP file to $modelPath"
                $ProgressPreference = 'SilentlyContinue'
                Expand-Archive -Path $filePath -DestinationPath $modelPath -Force
                $ProgressPreference = 'Continue'
                WriteLog "Extraction complete"
            } else {
                WriteLog "Unsupported file type: $fileExtension"
            }
            # Remove the downloaded file
            WriteLog "Removing $filePath"
            Remove-Item -Path $filePath -Force
            WriteLog "Complete"
        } else {
            WriteLog "No download link found for Windows $WindowsRelease."
        }
    } else {
        WriteLog "Failed to parse the download page for the MSI file."
    }
}

#### HP DRIVERS ###############################################################
function Get-HPDrivers {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$Make,
        [Parameter()]
        [string]$Model,
        [Parameter()]
        [ValidateSet("x64", "x86", "ARM64")]
        [string]$WindowsArch,
        [Parameter()]
        [ValidateSet(10, 11)]
        [int]$WindowsRelease,
        [Parameter()]
        [string]$WindowsVersion
    )

    # Download and extract the PlatformList.cab
    $PlatformListUrl = 'https://hpia.hpcloud.hp.com/ref/platformList.cab'
    $DriversFolder = "$DriversFolder\$Make"
    $PlatformListCab = "$DriversFolder\platformList.cab"
    $PlatformListXml = "$DriversFolder\PlatformList.xml"

    if (-not (Test-Path -Path $DriversFolder)) {
        WriteLog "Creating Drivers folder: $DriversFolder"
        New-Item -Path $DriversFolder -ItemType Directory -Force | Out-Null
        WriteLog "Drivers folder created"
    }
    WriteLog "Downloading $PlatformListUrl to $PlatformListCab"
    Start-BitsTransferWithRetry -Source $PlatformListUrl -Destination $PlatformListCab
    WriteLog "Download complete"
    WriteLog "Expanding $PlatformListCab to $PlatformListXml"
    Invoke-Process -FilePath expand.exe -ArgumentList "$PlatformListCab $PlatformListXml"
    WriteLog "Expansion complete"

    # Parse the PlatformList.xml to find the SystemID based on the ProductName
    [xml]$PlatformListContent = Get-Content -Path $PlatformListXml
    $ProductNodes = $PlatformListContent.ImagePal.Platform | Where-Object { $_.ProductName.'#text' -match $Model }

    # Create a list of unique ProductName entries
    $ProductNames = @()
    foreach ($node in $ProductNodes) {
        foreach ($productName in $node.ProductName) {
            if ($productName.'#text' -match $Model) {
                $ProductNames += [PSCustomObject]@{
                    ProductName = $productName.'#text'
                    SystemID    = $node.SystemID
                    OSReleaseID = $node.OS.OSReleaseIdFileName -replace 'H', 'h'
                    IsWindows11 = $node.OS.IsWindows11 -contains 'true'
                }
            }
        }
    }

    if ($ProductNames.Count -gt 1) {
        Write-Output "More than one model found matching '$Model':"
        WriteLog "More than one model found matching '$Model':"
        $ProductNames | ForEach-Object -Begin { $i = 1 } -Process {
            if ($VerbosePreference -ne 'Continue') {
                Write-Output "$i. $($_.ProductName)"
            }
            WriteLog "$i. $($_.ProductName)"
            $i++
        }
        $selection = Read-Host "Please select the number corresponding to the correct model"
        WriteLog "User selected model number: $selection"
        if ($selection -match '^\d+$' -and [int]$selection -le $ProductNames.Count) {
            $SelectedProduct = $ProductNames[[int]$selection - 1]
            $ProductName = $SelectedProduct.ProductName
            WriteLog "Selected model: $ProductName"
            $SystemID = $SelectedProduct.SystemID
            WriteLog "SystemID: $SystemID"
            $ValidOSReleaseIDs = $SelectedProduct.OSReleaseID
            WriteLog "Valid OSReleaseIDs: $ValidOSReleaseIDs"
            $IsWindows11 = $SelectedProduct.IsWindows11
            WriteLog "IsWindows11 supported: $IsWindows11"
        }
        else {
            WriteLog "Invalid selection. Exiting."
            if ($VerbosePreference -ne 'Continue') {
                Write-Host "Invalid selection. Exiting."
            }
            exit
        }
    }
    elseif ($ProductNames.Count -eq 1) {
        $SelectedProduct = $ProductNames[0]
        $ProductName = $SelectedProduct.ProductName
        WriteLog "Selected model: $ProductName"
        $SystemID = $SelectedProduct.SystemID
        WriteLog "SystemID: $SystemID"
        $ValidOSReleaseIDs = $SelectedProduct.OSReleaseID
        WriteLog "OSReleaseID: $ValidOSReleaseIDs"
        $IsWindows11 = $SelectedProduct.IsWindows11
        WriteLog "IsWindows11: $IsWindows11"
    }
    else {
        WriteLog "No models found matching '$Model'. Exiting."
        if ($VerbosePreference -ne 'Continue') {
            Write-Host "No models found matching '$Model'. Exiting."
        }
        exit
    }

    if (-not $SystemID) {
        WriteLog "SystemID not found for model: $Model Exiting."
        if ($VerbosePreference -ne 'Continue') {
            Write-Host "SystemID not found for model: $Model Exiting."
        }
        exit
    }

    # Validate if WindowsRelease is 11 and there is no IsWindows11 element set to true
    if ($WindowsRelease -eq 11 -and -not $IsWindows11) {
        WriteLog "WindowsRelease is set to 11, but no drivers are available for this Windows release. Retrying to download Windows 10 Drivers."
        Write-Output "WindowsRelease is set to 11, but no drivers are available for this Windows release. Retrying to download Windows 10 Drivers."
		Get-HPDrivers -Model $Model -Make $Make -WindowsArch $WindowsArch -WindowsRelease $WindowsRelease -WindowsVersion 10
        exit
    }

    # Validate WindowsVersion against OSReleaseID
    $OSReleaseIDs = $ValidOSReleaseIDs -split ' '
    $MatchingReleaseID = $OSReleaseIDs | Where-Object { $_ -eq "$WindowsVersion" }

    if (-not $MatchingReleaseID) {
        Write-Output "The specified WindowsVersion value '$WindowsVersion' is not valid for the selected model. Please select a valid OSReleaseID:"
        $OSReleaseIDs | ForEach-Object -Begin { $i = 1 } -Process {
            Write-Output "$i. $_"
            $i++
        }
        $selection = Read-Host "Please select the number corresponding to the correct OSReleaseID"
        WriteLog "User selected OSReleaseID number: $selection"
        if ($selection -match '^\d+$' -and [int]$selection -le $OSReleaseIDs.Count) {
            $WindowsVersion = $OSReleaseIDs[[int]$selection - 1]
            WriteLog "Selected OSReleaseID: $WindowsVersion"
        }
        else {
            WriteLog "Invalid selection. Exiting."
            exit
        }
    }

    # Modify WindowsArch for URL
    $Arch = $WindowsArch -replace "^x", ""

    # Construct the URL to download the driver XML cab for the model
    $ModelRelease = $SystemID + "_$Arch" + "_$WindowsRelease" + ".0.$WindowsVersion"
    $DriverCabUrl = "https://hpia.hpcloud.hp.com/ref/$SystemID/$ModelRelease.cab"
    $DriverCabFile = "$DriversFolder\$ModelRelease.cab"
    $DriverXmlFile = "$DriversFolder\$ModelRelease.xml"

    if (-not (Test-Url -Url $DriverCabUrl)) {
        WriteLog "HP Driver cab URL is not accessible: $DriverCabUrl Exiting"
        if ($VerbosePreference -ne 'Continue') {
            Write-Host "HP Driver cab URL is not accessible: $DriverCabUrl Exiting"
        }
        exit
    }

    # Download and extract the driver XML cab
    Writelog "Downloading HP Driver cab from $DriverCabUrl to $DriverCabFile"
    Start-BitsTransferWithRetry -Source $DriverCabUrl -Destination $DriverCabFile
    WriteLog "Expanding HP Driver cab to $DriverXmlFile"
    Invoke-Process -FilePath expand.exe -ArgumentList "$DriverCabFile $DriverXmlFile"

    # Parse the extracted XML file to download individual drivers
    [xml]$DriverXmlContent = Get-Content -Path $DriverXmlFile
    $baseUrl = "https://ftp.hp.com/pub/softpaq/sp"

    WriteLog "Downloading drivers for $ProductName"
    foreach ($update in $DriverXmlContent.ImagePal.Solutions.UpdateInfo) {
        if ($update.Category -notmatch '^Driver') {
            continue
        }
    
        $Name = $update.Name
        # Fix the name for drivers that contain illegal characters for folder name purposes
        $Name = $Name -replace '[\\\/\:\*\?\"\<\>\|]', '_'
        WriteLog "Downloading driver: $Name"
        $Category = $update.Category
        $Category = $Category -replace '[\\\/\:\*\?\"\<\>\|]', '_'
        $Version = $update.Version
        $Version = $Version -replace '[\\\/\:\*\?\"\<\>\|]', '_'
        $DriverUrl = "https://$($update.URL)"
        WriteLog "Driver URL: $DriverUrl"
        $DriverFileName = [System.IO.Path]::GetFileName($DriverUrl)
        $downloadFolder = "$DriversFolder\$ProductName\$Category"
        $DriverFilePath = Join-Path -Path $downloadFolder -ChildPath $DriverFileName

        if (Test-Path -Path $DriverFilePath) {
            WriteLog "Driver already downloaded: $DriverFilePath, skipping"
            continue
        }

        if (-not (Test-Path -Path $downloadFolder)) {
            WriteLog "Creating download folder: $downloadFolder"
            New-Item -Path $downloadFolder -ItemType Directory -Force | Out-Null
            WriteLog "Download folder created"
        }

        # Download the driver with retry
        WriteLog "Downloading driver to: $DriverFilePath"
        Start-BitsTransferWithRetry -Source $DriverUrl -Destination $DriverFilePath
        WriteLog 'Driver downloaded'

        # Make folder for extraction
        $extractFolder = "$downloadFolder\$Name\$Version\" + $DriverFileName.TrimEnd('.exe')
        Writelog "Creating extraction folder: $extractFolder"
        New-Item -Path $extractFolder -ItemType Directory -Force | Out-Null
        WriteLog 'Extraction folder created'
    
        # Extract the driver
        $arguments = "/s /e /f `"$extractFolder`""
        WriteLog "Extracting driver"
        Invoke-Process -FilePath $DriverFilePath -ArgumentList $arguments
        WriteLog "Driver extracted to: $extractFolder"

        # Delete the .exe driver file after extraction
        Remove-Item -Path $DriverFilePath -Force
        WriteLog "Driver installation file deleted: $DriverFilePath"
    }
    # Clean up the downloaded cab and xml files
    Remove-Item -Path $DriverCabFile, $DriverXmlFile, $PlatformListCab, $PlatformListXml -Force
    WriteLog "Driver cab and xml files deleted"
}

#### LENOVO DRIVERS ###########################################################
function Get-LenovoDrivers {
    param (
        [Parameter()]
        [string]$Model,
        [Parameter()]
        [ValidateSet("x64", "x86", "ARM64")]
        [string]$WindowsArch,
        [Parameter()]
        [ValidateSet(10, 11)]
        [int]$WindowsRelease
    )

    function Get-LenovoPSREF {
        param (
            [string]$ModelName
        )

        $url = "https://psref.lenovo.com/api/search/DefinitionFilterAndSearch/Suggest?kw=$ModelName"
        WriteLog "Querying Lenovo PSREF API for model: $ModelName"
        $OriginalVerbosePreference = $VerbosePreference
        $VerbosePreference = 'SilentlyContinue'
        $response = Invoke-WebRequest -Uri $url -UseBasicParsing -Headers $Headers -UserAgent $UserAgent
        $VerbosePreference = $OriginalVerbosePreference
        WriteLog "Complete"

        $jsonResponse = $response.Content | ConvertFrom-Json

        $products = @()
        foreach ($item in $jsonResponse.data) {
            $productName = $item.ProductName
            $machineTypes = $item.MachineType -split " / "

            foreach ($machineType in $machineTypes) {
                if ($machineType -eq $ModelName) {
                    WriteLog "Model name entered is a matching machine type"
                    $products = @()
                    $products += [pscustomobject]@{
                        ProductName = $productName
                        MachineType = $machineType
                    }
                    WriteLog "Product Name: $productName Machine Type: $machineType"
                    return $products
                }
                $products += [pscustomobject]@{
                    ProductName = $productName
                    MachineType = $machineType
                }
            }
        }

        return ,$products
    }
    
    # Parse the Lenovo PSREF page for the model
    $machineTypes = Get-LenovoPSREF -ModelName $Model
    if ($machineTypes.ProductName.Count -eq 0) {
        WriteLog "No machine types found for model: $Model"
        WriteLog "Enter a valid model or machine type in the -model parameter"
        exit
    } elseif ($machineTypes.ProductName.Count -eq 1) {
        $machineType = $machineTypes[0].MachineType
        $model = $machineTypes[0].ProductName
    } else {
        if ($VerbosePreference -ne 'Continue'){
            Write-Output "Multiple machine types found for model: $Model"
        }
        WriteLog "Multiple machine types found for model: $Model"
        for ($i = 0; $i -lt $machineTypes.ProductName.Count; $i++) {
            if ($VerbosePreference -ne 'Continue'){
                Write-Output "$($i + 1). $($machineTypes[$i].ProductName) ($($machineTypes[$i].MachineType))"
            }
            WriteLog "$($i + 1). $($machineTypes[$i].ProductName) ($($machineTypes[$i].MachineType))"
        }
        $selection = Read-Host "Enter the number of the model you want to select"
        $machineType = $machineTypes[$selection - 1].MachineType
        WriteLog "Selected machine type: $machineType"
        $model = $machineTypes[$selection - 1].ProductName
        WriteLog "Selected model: $model"
    }
    

    # Construct the catalog URL based on Windows release and machine type
    $ModelRelease = $machineType + "_Win" + $WindowsRelease
    $CatalogUrl = "https://download.lenovo.com/catalog/$ModelRelease.xml"
    WriteLog "Lenovo Driver catalog URL: $CatalogUrl"

    if (-not (Test-Url -Url $catalogUrl)) {
        Write-Error "Lenovo Driver catalog URL is not accessible: $catalogUrl"
        WriteLog "Lenovo Driver catalog URL is not accessible: $catalogUrl"
        exit
    }

    # Create the folder structure for the Lenovo drivers
    $driversFolder = "$DriversFolder\$Make"
    if (-not (Test-Path -Path $DriversFolder)) {
        WriteLog "Creating Drivers folder: $DriversFolder"
        New-Item -Path $DriversFolder -ItemType Directory -Force | Out-Null
        WriteLog "Drivers folder created"
    }

    # Download and parse the Lenovo catalog XML
    $LenovoCatalogXML = "$DriversFolder\$ModelRelease.xml"
    WriteLog "Downloading $catalogUrl to $LenovoCatalogXML"
    Start-BitsTransferWithRetry -Source $catalogUrl -Destination $LenovoCatalogXML
    WriteLog "Download Complete"
    $xmlContent = [xml](Get-Content -Path $LenovoCatalogXML)

    WriteLog "Parsing Lenovo catalog XML"
    # Process each package in the catalog
    foreach ($package in $xmlContent.packages.package) {
        $packageUrl = $package.location
        $category = $package.category

        #If category starts with BIOS, skip the package
        if ($category -like 'BIOS*') {
            continue
        }

        #If category name is 'Motherboard Devices Backplanes core chipset onboard video PCIe switches', truncate to 'Motherboard Devices' to shorten path
        if ($category -eq 'Motherboard Devices Backplanes core chipset onboard video PCIe switches') {
            $category = 'Motherboard Devices'
        }

        $packageName = [System.IO.Path]::GetFileName($packageUrl)
        #Remove the filename from the $packageURL
        $baseURL = $packageUrl -replace $packageName, "" 

        # Download the package XML
        $packageXMLPath = "$DriversFolder\$packageName"
        WriteLog "Downloading $category package XML $packageUrl to $packageXMLPath"
        If ((Start-BitsTransferWithRetry -Source $packageUrl -Destination $packageXMLPath) -eq $false) {
            Write-Output "Failed to download $category package XML: $packageXMLPath"
            WriteLog "Failed to download $category package XML: $packageXMLPath"
            continue
        }

        # Load the package XML content
        $packageXmlContent = [xml](Get-Content -Path $packageXMLPath)
        $packageType = $packageXmlContent.Package.PackageType.type
        $packageTitle = $packageXmlContent.Package.title.InnerText

        # Fix the name for drivers that contain illegal characters for folder name purposes
        $packageTitle = $packageTitle -replace '[\\\/\:\*\?\"\<\>\|]', '_'

        # If ' - ' is in the package title, truncate the title to the first part of the string.
        $packageTitle = $packageTitle -replace ' - .*', ''

        #Check if packagetype = 2. If packagetype is not 2, skip the package. $packageType is a System.Xml.XmlElement.
        #This filters out Firmware, BIOS, and other non-INF drivers
        if ($packageType -ne 2) {
            Remove-Item -Path $packageXMLPath -Force
            continue
        }

        # Extract the driver file name and the extract command
        $driverFileName = $packageXmlContent.Package.Files.Installer.File.Name
        $extractCommand = $packageXmlContent.Package.ExtractCommand

        #if extract command is empty/missing, skip the package
        if (!($extractCommand)) {
            Remove-Item -Path $packageXMLPath -Force
            continue
        }

        # Create the download URL and folder structure
        $driverUrl = $baseUrl + $driverFileName
        $downloadFolder = "$DriversFolder\$Model\$Category\$packageTitle"
        $driverFilePath = Join-Path -Path $downloadFolder -ChildPath $driverFileName

        # Check if file has already been downloaded
        if (Test-Path -Path $driverFilePath) {
            Write-Output "Driver already downloaded: $driverFilePath skipping"
            WriteLog "Driver already downloaded: $driverFilePath skipping"
            continue
        }

        if (-not (Test-Path -Path $downloadFolder)) {
            WriteLog "Creating download folder: $downloadFolder"
            New-Item -Path $downloadFolder -ItemType Directory -Force | Out-Null
            WriteLog "Download folder created"
        }

        # Download the driver with retry
        WriteLog "Downloading driver: $driverUrl to $driverFilePath"
        Start-BitsTransferWithRetry -Source $driverUrl -Destination $driverFilePath
        WriteLog "Driver downloaded"

        # Make folder for extraction
        $extractFolder = $downloadFolder + "\" + $driverFileName.TrimEnd($driverFileName[-4..-1])
        WriteLog "Creating extract folder: $extractFolder"
        New-Item -Path $extractFolder -ItemType Directory -Force | Out-Null
        WriteLog "Extract folder created"

        # Modify the extract command
        $modifiedExtractCommand = $extractCommand -replace '%PACKAGEPATH%', "`"$extractFolder`""

        # Extract the driver
        # Start-Process -FilePath $driverFilePath -ArgumentList $modifiedExtractCommand -Wait -NoNewWindow
        WriteLog "Extracting driver: $driverFilePath to $extractFolder"
        Invoke-Process -FilePath $driverFilePath -ArgumentList $modifiedExtractCommand
        WriteLog "Driver extracted"

        # Delete the .exe driver file after extraction
        WriteLog "Deleting driver installation file: $driverFilePath"
        Remove-Item -Path $driverFilePath -Force
        WriteLog "Driver installation file deleted: $driverFilePath"

        # Delete the package XML file after extraction
        WriteLog "Deleting package XML file: $packageXMLPath"
        Remove-Item -Path $packageXMLPath -Force
        WriteLog "Package XML file deleted"
    }

    #Delete the catalog XML file after processing
    WriteLog "Deleting catalog XML file: $LenovoCatalogXML"
    Remove-Item -Path $LenovoCatalogXML -Force
    WriteLog "Catalog XML file deleted"
}

#### DELL DRIVERS #############################################################
function Get-DellDrivers {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Model,
        [Parameter(Mandatory = $true)]
        [ValidateSet("x64", "x86", "ARM64")]
        [string]$WindowsArch
    )

    $catalogUrl = "http://downloads.dell.com/catalog/CatalogPC.cab"
    if (-not (Test-Url -Url $catalogUrl)) {
        WriteLog "Dell Catalog cab URL is not accessible: $catalogUrl Exiting"
        if ($VerbosePreference -ne 'Continue') {
            Write-Host "Dell Catalog cab URL is not accessible: $catalogUrl Exiting"
        }
        exit
    }

    if (-not (Test-Path -Path $DriversFolder)) {
        WriteLog "Creating Drivers folder: $DriversFolder"
        New-Item -Path $DriversFolder -ItemType Directory -Force | Out-Null
        WriteLog "Drivers folder created"
    }

    $DriversFolder = "$DriversFolder\$Make"
    WriteLog "Creating Dell Drivers folder: $DriversFolder"
    New-Item -Path $DriversFolder -ItemType Directory -Force | Out-Null
    WriteLog "Dell Drivers folder created"

    $DellCabFile = "$DriversFolder\CatalogPC.cab"
    WriteLog "Downloading Dell Catalog cab file: $catalogUrl to $DellCabFile"
    Start-BitsTransferWithRetry -Source $catalogUrl -Destination $DellCabFile
    WriteLog "Dell Catalog cab file downloaded"

    $DellCatalogXML = "$DriversFolder\CatalogPC.XML"
    WriteLog "Extracting Dell Catalog cab file to $DellCatalogXML"
    Invoke-Process -FilePath Expand.exe -ArgumentList "$DellCabFile $DellCatalogXML"
    WriteLog "Dell Catalog cab file extracted"

    $xmlContent = [xml](Get-Content -Path $DellCatalogXML)
    $baseLocation = "https://" + $xmlContent.manifest.baseLocation + "/"
    $latestDrivers = @{}

    $softwareComponents = $xmlContent.Manifest.SoftwareComponent | Where-Object { $_.ComponentType.value -eq "DRVR" }
    foreach ($component in $softwareComponents) {
        $models = $component.SupportedSystems.Brand.Model
        foreach ($item in $models) {
            if ($item.Display.'#cdata-section' -match $Model) {
                $validOS = $component.SupportedOperatingSystems.OperatingSystem | Where-Object { $_.osArch -eq $WindowsArch }
                if ($validOS) {
                    $driverPath = $component.path
                    $downloadUrl = $baseLocation + $driverPath
                    $driverFileName = [System.IO.Path]::GetFileName($driverPath)
                    $name = $component.Name.Display.'#cdata-section'
                    $name = $name -replace '[\\\/\:\*\?\"\<\>\|]', '_'
                    $name = $name -replace '[\,]', '-'
                    $category = $component.Category.Display.'#cdata-section'
                    $version = [version]$component.vendorVersion
                    $namePrefix = ($name -split '-')[0]

                    # Use hash table to store the latest driver for each category to prevent downloading older driver versions
                    if ($latestDrivers[$category]) {
                        if ($latestDrivers[$category][$namePrefix]) {
                            if ($latestDrivers[$category][$namePrefix].Version -lt $version) {
                                $latestDrivers[$category][$namePrefix] = [PSCustomObject]@{
                                    Name = $name; 
                                    DownloadUrl = $downloadUrl; 
                                    DriverFileName = $driverFileName; 
                                    Version = $version; 
                                    Category = $category 
                                }
                            }
                        }
                        else {
                            $latestDrivers[$category][$namePrefix] = [PSCustomObject]@{
                                Name = $name; 
                                DownloadUrl = $downloadUrl; 
                                DriverFileName = $driverFileName; 
                                Version = $version; 
                                Category = $category 
                            }
                        }
                    }
                    else {
                        $latestDrivers[$category] = @{}
                        $latestDrivers[$category][$namePrefix] = [PSCustomObject]@{
                            Name = $name; 
                            DownloadUrl = $downloadUrl; 
                            DriverFileName = $driverFileName; 
                            Version = $version; 
                            Category = $category 
                        }
                    }
                }
            }
        }
    }

    foreach ($category in $latestDrivers.Keys) {
        foreach ($driver in $latestDrivers[$category].Values) {
            $downloadFolder = "$DriversFolder\$Model\$($driver.Category)"
            $driverFilePath = Join-Path -Path $downloadFolder -ChildPath $driver.DriverFileName
            
            if (Test-Path -Path $driverFilePath) {
                Write-Output "Driver already downloaded: $driverFilePath skipping"
                continue
            }

            WriteLog "Downloading driver: $($driver.Name)"
            if (-not (Test-Path -Path $downloadFolder)) {
                WriteLog "Creating download folder: $downloadFolder"
                New-Item -Path $downloadFolder -ItemType Directory -Force | Out-Null
                WriteLog "Download folder created"
            }

            WriteLog "Downloading driver: $($driver.DownloadUrl) to $driverFilePath"
            try{
                Start-BitsTransferWithRetry -Source $driver.DownloadUrl -Destination $driverFilePath
                WriteLog "Driver downloaded"
            }catch{
                WriteLog "Failed to download driver: $($driver.DownloadUrl) to $driverFilePath"
                continue
            }
            

            $extractFolder = $downloadFolder + "\" + $driver.DriverFileName.TrimEnd($driver.DriverFileName[-4..-1])
            WriteLog "Creating extraction folder: $extractFolder"
            New-Item -Path $extractFolder -ItemType Directory -Force | Out-Null
            WriteLog "Extraction folder created"

            $arguments = "/s /e=`"$extractFolder`""
            WriteLog "Extracting driver: $driverFilePath to $extractFolder"
            Invoke-Process -FilePath $driverFilePath -ArgumentList $arguments
            WriteLog "Driver extracted"

            WriteLog "Deleting driver file: $driverFilePath"
            Remove-Item -Path $driverFilePath -Force
            WriteLog "Driver file deleted"
        }
    }
}

# Credit: https://github.com/Microsoft/vsts-tasks/blob/d052c35e5abfe5400341323a50826b9ca795166c/Tasks/Common/TlsHelper_/TlsHelper_.psm1
function Add-Tls12InSession {
    [CmdletBinding()]
    param()

    try {
        if ([Net.ServicePointManager]::SecurityProtocol.ToString().Split(',').Trim() -notcontains 'Tls12') {
            $securityProtocol=@()
            $securityProtocol+=[Net.ServicePointManager]::SecurityProtocol
            $securityProtocol+=[Net.SecurityProtocolType]3072
            [Net.ServicePointManager]::SecurityProtocol=$securityProtocol

            Write-Host "TLS12AddedInSession succeeded"
        }
        else {
            Write-Verbose 'TLS 1.2 already present in session.'
        }
    }
    catch {
        Write-Host "UnableToAddTls12InSession $($_.Exception.Message)"
    }
}


#### MAIN DRIVER INSTALL CODE #################################################
#[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13
$tsenv = New-Object -ComObject Microsoft.SMS.TSEnvironment
$OSDisk = $tsenv.Value("OSDisk")
$OSDiskPath = $OSDisk + "\"
$WinDir = $OSDiskPath + "Windows"
$DriversFolder=$OSDiskPath + "Drivers"
$DriverLog = $DriversFolder + "drivers-dism.log"

# TLS 1.2 for downloading from external hosts
Add-Tls12InSession -Verbose

If(!(Test-Path -PathType Container $DriversFolder))
{
      New-Item -ItemType Directory -Path $DriversFolder -Force
}
Write-Output "Listing Variables"
Get-Variable

$deviceMake=(Get-WmiObject -Class Win32_ComputerSystem -Property Manufacturer).Manufacturer
$deviceModel=(Get-WmiObject -Class Win32_ComputerSystem -Property Model).Model
$arch="x64"
$release=11
$version="22h2"

if ($deviceMake.ToLower() -like '*dell*') {
	Get-DellDrivers -Model $deviceModel -WindowsArch $arch
} elseif ($deviceMake.ToLower() -like '*lenovo*') { 
	Get-LenovoDrivers -Model $deviceModel -WindowsArch $arch -WindowsRelease $release
} elseif ($deviceMake.ToLower() -like '*microsoft*') { 
	Get-MicrosoftDrivers -Model $deviceModel -Make $deviceMake -WindowsRelease $release
} else {
	Get-HPDrivers -Model $deviceModel -Make $deviceMake -WindowsArch $arch -WindowsRelease $release -WindowsVersion $version
}

try {
	Add-WindowsDriver -Path "$OSDiskPath" -Driver "$DriversFolder" -Recurse -ErrorAction SilentlyContinue | Out-null
}
catch {
	WriteLog 'Some drivers failed to be added to the offline system. This can be expected. Continuing.'
}

