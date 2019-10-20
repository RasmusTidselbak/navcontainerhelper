<# 
 .Synopsis
  Invoke Newman tests in Container
 .Description
  Invoke the Postman Collection within the build project folder, with Newman inside the Container. Install-NavContainerNodeNewman should be run first, to ensure Newman availability.
 .Parameter containerName
  Name of the container in which you want to invoke the tests
 .Parameter tenant
  Name of the tenant in which context you want to invoke the tests
 .Parameter Credential
  Credentials for the user invoking the tests
 .Parameter appProjectFolder
  Location of the project. This folder (or any of its parents) needs to be shared with the container.
 .Parameter JUnitResultFileName
  Filename where the function should place a JUnit compatible result file
 .Example
 Invoke-NavContainerNewmanTest -containerName $containerName -credential $credential -appProjectFolder "C:\Users\raaen\Documents\AL\Test" -
 .Example
 Invoke-NavContainerNewmanTest -containerName $containerName -credential $credential -appProjectFolder "C:\Users\raaen\Documents\AL\Test" -JUnitResultFileName "c:\ProgramData\NavContainerHelper\$containername.results.xml"
#>
function Invoke-NavContainerNewmanTest {
    Param(
        [Parameter(Mandatory = $true)]
        [string] $containerName, 
        [Parameter(Mandatory = $false)]
        [pscredential] $tenant,
        [Parameter(Mandatory = $true)]
        [pscredential] $credential,
        [Parameter(Mandatory = $true)]
        [string] $appProjectFolder,
        [Parameter(Mandatory = $false)]
        [string] $JUnitResultFileName
    )

    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($credential.Password)
    $UnsecurePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

    $containerTestResultsFile = ""
    if ($JUnitResultFileName) {
        $containerTestResultsFile = Get-NavContainerPath -containerName $containerName -path $JUnitResultFileName
        if ("$containerTestResultsFile" -eq "") {
            throw "The path for TestResultsFile ($JUnitResultFileName) is not shared with the container."
        }
    }

    $customConfig = Get-NavContainerServerConfiguration -ContainerName $containerName

    # $parameters = @{}
    # if ($customConfig.ClientServicesCredentialType -eq "Windows") {
    #     $parameters += @{ "usedefaultcredential" = $true }
    # }
    # else {
    #     if (!($credential)) {
    #         throw "You need to specify credentials when you are not using Windows Authentication"
    #     }
    #     $parameters += @{ "credential" = $credential }
    # }

    if ($customConfig.ODataServicesSSLEnabled -eq "true") {
        $protocol = "https://"
    }
    else {
        $protocol = "http://"
    }

    $ip = Get-NavContainerIpAddress -containerName $containerName
    if ($ip) {
        $BaseUrl = "${protocol}${ip}:$($customConfig.ODataServicesPort)/$($customConfig.ServerInstance)"
    }
    else {
        $BaseUrl = $customconfig.PublicODataBaseUrl.Replace("/OData", "")
    }

    Get-ChildItem -Path $appProjectFolder -Filter "*postman_collection.json" -Recurse | ForEach-Object {
        Invoke-ScriptInNavContainer -containerName $containerName -scriptblock { 
            Param($username, $UnsecurePassword, $PostmanCollection, $JUnitResultFileName, $BaseUrl)

            Remove-Item -Path .\newman\ -Force -Recurse -ErrorAction SilentlyContinue

            newman run $PostmanCollection --env-var "BaseUrl=$BaseUrl" --env-var "username=$($UserName)" --env-var "password=$UnsecurePassword" --reporters cli, junit
        
            Get-ChildItem .\newman\ | % { Move-Item -Path $_.FullName -Destination $JUnitResultFileName }
        } -ArgumentList @($credential.UserName, $UnsecurePassword, $_.FullName, $containerTestResultsFile, $BaseUrl)
    } 
}