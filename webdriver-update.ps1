param (
    $registryRoot        = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths",
    $edgeRegistryPath    = "$registryRoot\msedge.exe",
    $chromeRegistryPath  = "$registryRoot\chrome.exe",
    $webDriversPath      = "C:\Program Files\LEAPWORK\Agent",
    $edgeDriverPath      = "$($webDriversPath)\msedgedriver.exe",
    $chromeDriverPath    = "$($webDriversPath)\chromedriver.exe",
    $chromeDriverWebsite = "https://chromedriver.chromium.org/downloads",
    $chromeDriverUrlBase = "https://chromedriver.storage.googleapis.com",
    $chromeDriverUrlEnd  = "chromedriver_win32.zip",
    $edgeDriverWebsite   = "https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/" 
)
#endregion INITIALIZATION

#region FUNCTIONS
# function checking driver version using the -v switch of each driver
function Get-LocalDriverVersion{
    param(
        $pathToDriver                                               # direct path to the driver
    )
    $processInfo = New-Object System.Diagnostics.ProcessStartInfo   # need to pass the switch & catch the output, hence ProcessStartInfo is used

    $processInfo.FileName               = $pathToDriver
    $processInfo.RedirectStandardOutput = $true                     # need to catch the output - the version
    $processInfo.Arguments              = "-v"
    $processInfo.UseShellExecute        = $false                    # hide execution

    $process = New-Object System.Diagnostics.Process

    $process.StartInfo  = $processInfo
    $process.Start()    | Out-Null
    $process.WaitForExit()                                          # run synchronously, we need to wait for result
    $processStOutput    = $process.StandardOutput.ReadToEnd()

    if ($pathToDriver.Contains("msedgedriver")){
        return ($processStOutput -split " ")[3]                     # MS Edge returns version on 4th place in the output (be carefulm in old versions it was on 1st as well)... 
    }
    else {
        return ($processStOutput -split " ")[1]                     # ... while Chrome on 2nd place
    }
}

# function evaluating a need for update
function Confirm-NeedForUpdate{
    param(
        $v1,                                                                                 # version 1 to compare
        $v2                                                                                  # version 2 to compare
    )
    return $v1.Substring(0, $v1.LastIndexOf(".")) -ne $v2.Substring(0, $v2.LastIndexOf(".")) # return true if update is needed, otherwise false. Ignore last minor version - it's not so important and can be skipped
}
#endregion FUNCTIONS

#region MAIN SCRIPT
# firstly check which browser versions are installed (from registry)
$edgeVersion   = (Get-Item (Get-ItemProperty $edgeRegistryPath).'(Default)').VersionInfo.ProductVersion
$chromeVersion = (Get-Item (Get-ItemProperty $chromeRegistryPath).'(Default)').VersionInfo.ProductVersion   

# check which driver versions are installed
$edgeDriverVersion   = Get-LocalDriverVersion -pathToDriver $edgeDriverPath
$chromeDriverVersion = Get-LocalDriverVersion -pathToDriver $chromeDriverPath

# download new MS Edge driver if neccessary
if (Confirm-NeedForUpdate $edgeVersion $edgeDriverVersion){
    # find exact matching version
    $edgeDriverAvailableVersions = (Invoke-RestMethod $edgeDriverWebsite) -split " " | where {$_ -like "*href=*win64*"} | % {$_.replace("href=","").replace('"','')}
    $downloadLink                = $edgeDriverAvailableVersions | where {$_ -like "*/$edgeVersion/*"}

    # if cannot find (e.g. it's too new to have a web driver), look for relevant major version
    if (!$downloadLink){
        $browserMajorVersion = $edgeVersion.Substring(0, $edgeVersion.IndexOf("."))
        $downloadLink        = $edgeDriverAvailableVersions | where {$_ -like "*/$browserMajorVersion*"}
    }

    # in case of multiple links, take the first only
    if ($downloadLink.Count -gt 1) {
        $downloadLink = $downloadLink[0]
    }

    # download the file
    Invoke-WebRequest $downloadLink -OutFile "edgeNewDriver.zip"

    # epand archive and replace the old file
    Expand-Archive "edgeNewDriver.zip"              -DestinationPath "edgeNewDriver\"                      -Force
    Move-Item      "edgeNewDriver/msedgedriver.exe" -Destination     "$($webDriversPath)\msedgedriver.exe" -Force

    # clean-up
    Remove-Item "edgeNewDriver.zip" -Force
    Remove-Item "edgeNewDriver"     -Recurse -Force
}                   

# download new Chrome driver if neccessary
if (Confirm-NeedForUpdate $chromeVersion $chromeDriverVersion){
    # find exact matching version
    $chromeDriverAvailableVersions = (Invoke-RestMethod $chromeDriverWebsite) -split " " | where {$_ -like "*href=*?path=*"} | % {$_.replace("href=","").replace('"','')}
    $versionLink                   = $chromeDriverAvailableVersions | where {$_ -like "*$chromeVersion/*"}
    
    # if cannot find (e.g. it's too new to have a web driver), look for relevant major version
    if (!$versionLink){
        $browserMajorVersion = $chromeVersion.Substring(0, $chromeVersion.IndexOf("."))
        $versionLink         = $chromeDriverAvailableVersions | where {$_ -like "*$browserMajorVersion.*"}
    }

    # in case of multiple links, take the first only
    if ($versionLink.Count -gt 1){
        $versionLink = $versionLink[0]
    }

    # build tge download URL according to found version and download URL schema
    $version      = ($versionLink -split"=" | where {$_ -like "*.*.*.*/"}).Replace('/','')
    $downloadLink = "$chromeDriverUrlBase/$version/$chromeDriverUrlEnd"

    # download the file
    Invoke-WebRequest $downloadLink -OutFile "chromeNewDriver.zip"

    # epand archive and replace the old file
    Expand-Archive "chromeNewDriver.zip"              -DestinationPath "chromeNewDriver\"                    -Force
    Move-Item      "chromeNewDriver/chromedriver.exe" -Destination     "$($webDriversPath)\chromedriver.exe" -Force

    # clean-up
    Remove-Item "chromeNewDriver.zip" -Force
    Remove-Item "chromeNewDriver" -Recurse -Force
}