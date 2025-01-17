﻿#requires -Version 5.0
using namespace Microsoft.Dynamics.Framework.UI.Client
using namespace Microsoft.Dynamics.Framework.UI.Client.Interactions

class ClientContext {

    $events = @()
    $clientSession = $null
    $culture = ""
    $caughtForm = $null
    $debugMode = $false
    $addressUri = $null

    ClientContext([string] $serviceUrl, [string] $accessToken, [timespan] $interactionTimeout, [string] $culture) {
        $this.Initialize($serviceUrl, ([AuthenticationScheme]::AzureActiveDirectory), (New-Object Microsoft.Dynamics.Framework.UI.Client.TokenCredential -ArgumentList $accessToken), $interactionTimeout, $culture)
    }

    ClientContext([string] $serviceUrl, [string] $accessToken) {
        $this.Initialize($serviceUrl, ([AuthenticationScheme]::AzureActiveDirectory), (New-Object Microsoft.Dynamics.Framework.UI.Client.TokenCredential -ArgumentList $accessToken), ([timespan]::FromMinutes(10)), 'en-US')
    }

    ClientContext([string] $serviceUrl, [pscredential] $credential, [timespan] $interactionTimeout, [string] $culture) {
        $this.Initialize($serviceUrl, ([AuthenticationScheme]::UserNamePassword), (New-Object System.Net.NetworkCredential -ArgumentList $credential.UserName, $credential.Password), $interactionTimeout, $culture)
    }

    ClientContext([string] $serviceUrl, [pscredential] $credential) {
        $this.Initialize($serviceUrl, ([AuthenticationScheme]::UserNamePassword), (New-Object System.Net.NetworkCredential -ArgumentList $credential.UserName, $credential.Password), ([timespan]::FromMinutes(10)), 'en-US')
    }

    ClientContext([string] $serviceUrl, [timespan] $interactionTimeout, [string] $culture) {
        $this.Initialize($serviceUrl, ([AuthenticationScheme]::Windows), $null, $interactionTimeout, $culture)
    }
    
    ClientContext([string] $serviceUrl) {
        $this.Initialize($serviceUrl, ([AuthenticationScheme]::Windows), $null, ([timespan]::FromMinutes(10)), 'en-US')
    }
    
    Initialize([string] $serviceUrl, [AuthenticationScheme] $authenticationScheme, [System.Net.ICredentials] $credential, [timespan] $interactionTimeout, [string] $culture) {
        $this.addressUri = New-Object System.Uri -ArgumentList $serviceUrl
        $this.addressUri = [ServiceAddressProvider]::ServiceAddress($this.addressUri)
        $jsonClient = New-Object JsonHttpClient -ArgumentList $this.addressUri, $credential, $authenticationScheme
        $httpClient = ($jsonClient.GetType().GetField("httpClient", [Reflection.BindingFlags]::NonPublic -bor [Reflection.BindingFlags]::Instance)).GetValue($jsonClient)
        $httpClient.Timeout = $interactionTimeout
        $this.clientSession = New-Object ClientSession -ArgumentList $jsonClient, (New-Object NonDispatcher), (New-Object 'TimerFactory[TaskTimer]')
        $this.culture = $culture
        $this.OpenSession()
    }

    OpenSession() {
        $clientSessionParameters = New-Object ClientSessionParameters
        $clientSessionParameters.CultureId = $this.culture
        $clientSessionParameters.UICultureId = $this.culture
        $clientSessionParameters.AdditionalSettings.Add("IncludeControlIdentifier", $true)
    
        $this.events += @(Register-ObjectEvent -InputObject $this.clientSession -EventName MessageToShow -Action {
            Write-Host -ForegroundColor Yellow "Message : $($EventArgs.Message)"
            if ($this.debugMode) {
                $this.GetAllForms() | ForEach-Object {
                    $formInfo = $this.GetFormInfo($_)
                    if ($formInfo) {
                        Write-Host -ForegroundColor Yellow "Title: $($formInfo.title)"
                        Write-Host -ForegroundColor Yellow "Title: $($formInfo.identifier)"
                        $formInfo.controls | ConvertTo-Json -Depth 99 | Out-Host
                    }
                }
            }
        })
        $this.events += @(Register-ObjectEvent -InputObject $this.clientSession -EventName CommunicationError -Action {
            Write-Host -ForegroundColor Red "CommunicationError : $($EventArgs.Exception.Message)"
            if ($null -ne $EventArgs.Exception.InnerException) {
                Write-Host -ForegroundColor Red "CommunicationError InnerException : $($EventArgs.Exception.InnerException)"    
            }
            if ($this.debugMode) {
                $this.GetAllForms() | ForEach-Object {
                    $formInfo = $this.GetFormInfo($_)
                    if ($formInfo) {
                        Write-Host -ForegroundColor Yellow "Title: $($formInfo.title)"
                        Write-Host -ForegroundColor Yellow "Title: $($formInfo.identifier)"
                        $formInfo.controls | ConvertTo-Json -Depth 99 | Out-Host
                    }
                }
            }
            Remove-ClientSession
        })
        $this.events += @(Register-ObjectEvent -InputObject $this.clientSession -EventName UnhandledException -Action {
            Write-Host -ForegroundColor Red "UnhandledException : $($EventArgs.Exception.Message)"
            if ($null -ne $EventArgs.Exception.InnerException) {
                Write-Host -ForegroundColor Red "UnhandledException InnerException : $($EventArgs.Exception.InnerException)"    
            }
            if ($this.debugMode) {
                $this.GetAllForms() | ForEach-Object {
                    $formInfo = $this.GetFormInfo($_)
                    if ($formInfo) {
                        Write-Host -ForegroundColor Yellow "Title: $($formInfo.title)"
                        Write-Host -ForegroundColor Yellow "Title: $($formInfo.identifier)"
                        $formInfo.controls | ConvertTo-Json -Depth 99 | Out-Host
                    }
                }
            }
            Remove-ClientSession
        })
        $this.events += @(Register-ObjectEvent -InputObject $this.clientSession -EventName InvalidCredentialsError -Action {
            Write-Host -ForegroundColor Red "InvalidCredentialsError"
            Remove-ClientSession
        })
        $this.events += @(Register-ObjectEvent -InputObject $this.clientSession -EventName UriToShow -Action {
            Write-Host -ForegroundColor Yellow "UriToShow : $($EventArgs.UriToShow)"
        })
        $this.events += @(Register-ObjectEvent -InputObject $this.clientSession -EventName LookupFormToShow -Action { 
            Write-Host -ForegroundColor Yellow "Open Lookup form"
        })
        $this.events += @(Register-ObjectEvent -InputObject $this.clientSession -EventName DialogToShow -Action {
            $form = $EventArgs.DialogToShow
            if ($this.debugMode) {
                Write-Host -ForegroundColor Yellow "Show dialog $($form.ControlIdentifier)"
                $formInfo = $this.GetFormInfo($form)
                if ($formInfo) {
                    Write-Host -ForegroundColor Yellow "Title: $($formInfo.title)"
                    $formInfo.controls | ConvertTo-Json -Depth 99 | Out-Host
                }
            }
            if ( $form.ControlIdentifier -eq "00000000-0000-0000-0800-0000836bd2d2" ) {
                $errorControl = $form.ContainedControls | Where-Object { $_ -is [ClientStaticStringControl] } | Select-Object -First 1                
                Write-Host -ForegroundColor Red "ERROR: $($errorControl.StringValue)"
                $this.CloseForm($form)
            }
            elseif ( $form.ControlIdentifier -eq "00000000-0000-0000-0300-0000836bd2d2" ) {
                $errorControl = $form.ContainedControls | Where-Object { $_ -is [ClientStaticStringControl] } | Select-Object -First 1                
                Write-Host -ForegroundColor Yellow "WARNING: $($errorControl.StringValue)"
                $this.CloseForm($form)
            }
        })
    
        $this.clientSession.OpenSessionAsync($clientSessionParameters)
        $this.Awaitstate([ClientSessionState]::Ready)
    }
    #
    
    Dispose() {
        $this.events | ForEach-Object { Unregister-Event $_.Name }
        $this.events = @()
    
        try {
            if ($this.clientSession -and ($this.clientSession.State -ne ([ClientSessionState]::Closed))) {
                $this.clientSession.CloseSessionAsync()
                $this.AwaitState([ClientSessionState]::Closed)
            }
        }
        catch {
        }
    }
    
    AwaitState([ClientSessionState] $state) {
        While ($this.clientSession.State -ne $state) {
            Start-Sleep -Milliseconds 100
            if ($this.clientSession.State -eq [ClientSessionState]::InError) {
                throw "ClientSession in Error"
            }
            if ($this.clientSession.State -eq [ClientSessionState]::TimedOut) {
                throw "ClientSession timed out"
            }
            if ($this.clientSession.State -eq [ClientSessionState]::Uninitialized) {
                throw "ClientSession is Uninitialized"
            }
        }
    }
    
    InvokeInteraction([ClientInteraction] $interaction) {
        $this.clientSession.InvokeInteractionAsync($interaction)
        $this.AwaitState([ClientSessionState]::Ready)
    }
    
    [ClientLogicalForm] InvokeInteractionAndCatchForm([ClientInteraction] $interaction) {
        $Global:PsTestRunnerCaughtForm = $null
        $formToShowEvent = Register-ObjectEvent -InputObject $this.clientSession -EventName FormToShow -Action { 
            $Global:PsTestRunnerCaughtForm = $EventArgs.FormToShow
        }
        try {
            $this.InvokeInteraction($interaction)
            if (!($Global:PsTestRunnerCaughtForm)) {
                $this.CloseAllWarningForms()
            }
        } finally {
            Unregister-Event -SourceIdentifier $formToShowEvent.Name
        }
        $form = $Global:PsTestRunnerCaughtForm
        Remove-Variable PsTestRunnerCaughtForm -Scope Global
        return $form
    }
    
    [ClientLogicalForm] OpenForm([int] $page) {
        $interaction = New-Object OpenFormInteraction
        $interaction.Page = $page
        return $this.InvokeInteractionAndCatchForm($interaction)
    }
    
    CloseForm([ClientLogicalControl] $form) {
        $this.InvokeInteraction((New-Object CloseFormInteraction -ArgumentList $form))
    }
    
    [ClientLogicalForm[]]GetAllForms() {
        $forms = @()
        $this.clientSession.OpenedForms.GetEnumerator() | ForEach-Object { $forms += $_ }
        return $forms
    }
    
    [string]GetErrorFromErrorForm() {
        $errorText = ""
        $this.clientSession.OpenedForms.GetEnumerator() | ForEach-Object {
            $form = $_
            if ( $form.ControlIdentifier -eq "00000000-0000-0000-0800-0000836bd2d2" ) {
                $form.ContainedControls | Where-Object { $_ -is [ClientStaticStringControl] } | ForEach-Object {
                    $errorText = $_.StringValue
                }
            }
        }
        return $errorText
    }
    
    [string]GetWarningFromWarningForm() {
        $warningText = ""
        $this.clientSession.OpenedForms.GetEnumerator() | ForEach-Object {
            $form = $_
            if ( $form.ControlIdentifier -eq "00000000-0000-0000-0300-0000836bd2d2" ) {
                $form.ContainedControls | Where-Object { $_ -is [ClientStaticStringControl] } | ForEach-Object {
                    $warningText = $_.StringValue
                }
            }
        }
        return $warningText
    }

    [Hashtable]GetFormInfo([ClientLogicalForm] $form) {
    
        function Dump-RowControl {
            Param(
                [ClientLogicalControl] $control
            )
            @{
                "$($control.Name)" = $control.ObjectValue
            }
        }
    
        function Dump-Control {
            Param(
                [ClientLogicalControl] $control
            )
    
            $output = @{
                "name" = $control.Name
                "type" = $control.GetType().Name
                "identifier" = $control.ControlIdentifier
            }
            if ($control -is [ClientGroupControl]) {
                $output += @{
                    "caption" = $control.Caption
                    "mappingHint" = $control.MappingHint
                    "children" = @($control.Children | ForEach-Object { Dump-Control -control $_ })
                }
            } elseif ($control -is [ClientStaticStringControl]) {
                $output += @{
                    "value" = $control.StringValue
                }
            } elseif ($control -is [ClientInt32Control]) {
                $output += @{
                    "value" = $control.ObjectValue
                }
            } elseif ($control -is [ClientStringControl]) {
                $output += @{
                    "value" = $control.stringValue
                }
            } elseif ($control -is [ClientActionControl]) {
                $output += @{
                    "caption" = $control.Caption
                }
            } elseif ($control -is [ClientFilterLogicalControl]) {
            } elseif ($control -is [ClientRepeaterControl]) {
                $output += @{
                    "$($control.name)" = @()
                }
                $index = 0
                while ($true) {
                    if ($index -ge ($control.Offset + $control.DefaultViewport.Count)) {
                        break
                    }
                    $rowIndex = $index - $control.Offset
                    if ($rowIndex -ge $control.DefaultViewport.Count) {
                        break 
                    }
                    $row = $control.DefaultViewport[$rowIndex]
                    $rowoutput = @{}
                    $row.Children | ForEach-Object { $rowoutput += Dump-RowControl -control $_ }
                    $output[$control.name] += $rowoutput
                    $index++
                }
            }
            else {
            }
            $output
        }
    
        return @{
            "title" = "$($form.Name) $($form.Caption)"
            "identifier" = $form.ControlIdentifier
            "controls" = $form.Children | ForEach-Object { Dump-Control -control $_ }
        }
    }
    
    CloseAllForms() {
        $this.GetAllForms() | ForEach-Object { $this.CloseForm($_) }
    }

    CloseAllErrorForms() {
        $this.GetAllForms() | ForEach-Object {
            if ($_.ControlIdentifier -eq "00000000-0000-0000-0800-0000836bd2d2") {
                $this.CloseForm($_)
            }
        }
    }

    CloseAllWarningForms() {
        $this.GetAllForms() | ForEach-Object {
            if ($_.ControlIdentifier -eq "00000000-0000-0000-0300-0000836bd2d2") {
                $this.CloseForm($_)
            }
        }
    }
    
    [ClientLogicalControl]GetControlByCaption([ClientLogicalControl] $control, [string] $caption) {
        return $control.ContainedControls | Where-Object { $_.Caption.Replace("&","") -eq $caption } | Select-Object -First 1
    }
    
    [ClientLogicalControl]GetControlByName([ClientLogicalControl] $control, [string] $name) {
        $result = $control.ContainedControls | Where-Object { $_.Name -eq $name } | Select-Object -First 1
        if (-not $result) {
            $result = $control.ContainedControls | Where-Object { $_.Caption -eq $name } | Select-Object -First 1
        }
        return $result
    }
    
    [ClientLogicalControl]GetControlByType([ClientLogicalControl] $control, [Type] $type) {
        return $control.ContainedControls | Where-Object { $_ -is $type } | Select-Object -First 1
    }
    
    SaveValue([ClientLogicalControl] $control, [string] $newValue) {
        $this.InvokeInteraction((New-Object SaveValueInteraction -ArgumentList $control, $newValue))
    }
    
    ScrollRepeater([ClientRepeaterControl] $repeater, [int] $by) {
        $this.InvokeInteraction((New-Object ScrollRepeaterInteraction -ArgumentList $repeater, $by))
    }
    
    ActivateControl([ClientLogicalControl] $control) {
        $this.InvokeInteraction((New-Object ActivateControlInteraction -ArgumentList $control))
    }
    
    [ClientActionControl]GetActionByCaption([ClientLogicalControl] $control, [string] $caption) {
        return $control.ContainedControls | Where-Object { ($_ -is [ClientActionControl]) -and ($_.Caption.Replace("&","") -eq $caption) } | Select-Object -First 1
    }
    
    [ClientActionControl]GetActionByName([ClientLogicalControl] $control, [string] $name) {
        $result = $control.ContainedControls | Where-Object { ($_ -is [ClientActionControl]) -and ($_.Name -eq $name) } | Select-Object -First 1
        if (-not $result) {
            $result = $control.ContainedControls | Where-Object { ($_ -is [ClientActionControl]) -and ($_.Caption -eq $name) } | Select-Object -First 1
        }
        return $result
    }
    
    InvokeAction([ClientActionControl] $action) {
        $this.InvokeInteraction((New-Object InvokeActionInteraction -ArgumentList $action))
    }
    
    [ClientLogicalForm]InvokeActionAndCatchForm([ClientActionControl] $action) {
        return $this.InvokeInteractionAndCatchForm((New-Object InvokeActionInteraction -ArgumentList $action))
    }
}
