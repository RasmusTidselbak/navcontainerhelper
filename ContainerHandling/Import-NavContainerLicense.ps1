﻿<# 
 .Synopsis
  Import License file to a NAV/BC Container
 .Description
  Import a license from a file or a url to a container
 .Parameter containerName
  Name of the container in which you want to import a license
 .Parameter licenseFile
  Path or secure url to the licensefile to upload
 .Example
  Import-NavContainerLicense -containerName test -licenseFile c:\temp\mylicense.flf
 .Example
  Import-NavContainerLicense -containerName test -licenseFile "https://www.dropbox.com/s/fhwfwjfjwhff/license.flf?dl=1"
#>
function Import-NavContainerLicense {
    Param (
        [string] $containerName = "navserver", 
        [Parameter(Mandatory=$true)]
        [string] $licenseFile
    )

    if ($licensefile.StartsWith("https://", "OrdinalIgnoreCase") -or $licensefile.StartsWith("http://", "OrdinalIgnoreCase")) {
        $containerLicenseFile = $licensefile
    } else {
        $containerLicenseFile = Get-NavContainerPath -containerName $containerName -path $licenseFile
        $copied = $false
        if ("$containerLicenseFile" -eq "") {
            $containerLicenseFile = Join-Path "c:\run" ([System.IO.Path]::GetFileName($licensefile))
            Copy-FileToNavContainer -containerName $containerName -localPath $licensefile -containerPath $containerLicenseFile
            $copied = $true
        }
    }

    Invoke-ScriptInNavContainer -containerName $containerName -ScriptBlock { Param($licensefile)

        if ($licensefile.StartsWith("https://") -or $licensefile.StartsWith("http://"))
        {
            $licensefileurl = $licensefile
            $licensefile = (Join-Path $runPath "license.flf")
            Write-Host "Downloading license file '$licensefileurl'"
            (New-Object System.Net.WebClient).DownloadFile($licensefileurl, $licensefile)

            $bytes = [System.IO.File]::ReadAllBytes($licenseFile)
            $text = [System.Text.Encoding]::ASCII.GetString($bytes, 0, 100)
            if (!($text.StartsWith("Microsoft Software License Information"))) {
                Remove-Item -Path $licenseFile -Force
                Write-Error "Specified license file Uri isn't a direct download Uri"
            }
        }
    
        Write-Host "Import License $licensefile"
        Import-NAVServerLicense -LicenseFile $licensefile -ServerInstance $ServerInstance -Database NavDatabase -WarningAction SilentlyContinue
    
    }  -ArgumentList $containerLicenseFile
}
Set-Alias -Name Import-BCContainerLicense -Value Import-NavContainerLicense
Export-ModuleMember -Function Import-NavContainerLicense -Alias Import-BCContainerLicense
