<# 
 .Synopsis
  Install Node and Newman in container
 .Description
  Install a specified version of Node and the latest version of Newman (Postman collection runner), in the container.
 .Parameter containerName
  Name of the container in which you want to install Node and Newman
 .Parameter NodeVersion
  The version of Node which should be installed
 .Example
  Install-NavContainerNodeNewman -containerName $containerName
#>
function Install-NavContainerNodeNewman {
    Param(
        [Parameter(Mandatory = $true)]
        [string] $containerName,
        [string] $NodeVersion = "8.12.0"
    )

    Invoke-ScriptInNavContainer -containerName $containerName -scriptblock { 
        Param($NodeVersion)
        Invoke-WebRequest $('https://nodejs.org/dist/v{0}/node-v{0}-x64.msi' -f $NodeVersion) -OutFile 'node.msi'
        Start-Process "msiexec.exe" -ArgumentList @("/q", "/i", "node.msi") -Wait
        npm install -g newman
    } -ArgumentList $NodeVersion
}