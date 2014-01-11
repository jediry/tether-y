Param([string] $Action,
      [string] $ProxyIP = $Null,
      [string] $ProxyPort = $Null,
      $ExcludeAddresses = $Null,
      [switch] $WhatIf = $False)


$Script:ProxyRegKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings"

Function Start-ProxyEngine([string] $ProxyIP,
                           [string] $ProxyPort,
                           [switch] $WhatIf) {
    $ProxyString = "$IPAddress`:$Port"

    If (!$WhatIf) {
        $CurrentProxyString = (Get-ItemProperty -Path $Script:ProxyRegKey ProxyServer -ErrorAction:SilentlyContinue).ProxyServer
        If ($CurrentProxyString -ne $ProxyString) {
            Set-ItemProperty -Path $Script:ProxyRegKey ProxyServer $ProxyString -ErrorAction:SilentlyContinue
            If ($?) {
                Write-Host -ForegroundColor Green "[SUCCESS]"
            } Else {
                Write-Host -ForegroundColor Red "[FAILURE]"
                Throw "Failed to set proxy address"
            }
        } Else {
            Write-Host -ForegroundColor Yellow "[NO CHANGE]"
        }
    } Else {
        Write-Host -ForegroundColor Cyan "[SKIPPED]"
    }

    Write-Host -NoNewline "Enabling proxy..."
    If (!$WhatIf) {
        $CurrentProxyEnable = (Get-ItemProperty -Path $Script:ProxyRegKey ProxyEnable -ErrorAction:SilentlyContinue).ProxyEnable
        If ($CurrentProxyEnable -ne 1) {
            Set-ItemProperty -Path $Script:ProxyRegKey ProxyEnable 1 -ErrorAction:SilentlyContinue
            If ($?) {
                Write-Host -ForegroundColor Green "[SUCCESS]"
            } Else {
                Write-Host -ForegroundColor Red "[FAILURE]"
                Throw "Failed to enable proxy"
            }
        } Else {
            Write-Host -ForegroundColor Yellow "[NO CHANGE]"
        }
    } Else {
        Write-Host -ForegroundColor Cyan "[SKIPPED]"
    }

    [System.Net.WebRequest]::DefaultWebProxy = [System.Net.WebProxy]$ProxyString
}


Function Stop-ProxyEngine([switch] $WhatIf) {
    $CurrentProxyEnable = (Get-ItemProperty -Path $Script:ProxyRegKey ProxyEnable -ErrorAction:SilentlyContinue).ProxyEnable
    If ($CurrentProxyEnable -ne 0) {
        Set-ItemProperty -Path $Script:ProxyRegKey ProxyEnable 0 -ErrorAction:SilentlyContinue
        If ($?) {
            Write-Host -ForegroundColor Green "[SUCCESS]"
        } Else {
            Write-Host -ForegroundColor Red "[FAILURE]"
            Throw "Failed to disable proxy"
        }
    } Else {
        Write-Host -ForegroundColor Yellow "[NO CHANGE]"
    }

    Write-Host -NoNewline "Clearing proxy server address..."
    $CurrentProxyString = (Get-ItemProperty -Path $Script:ProxyRegKey ProxyServer -ErrorAction:SilentlyContinue).ProxyServer
    If (![String]::IsNullOrEmpty($CurrentProxyString)) {
        Clear-ItemProperty -Path $Script:ProxyRegKey ProxyServer -ErrorAction:SilentlyContinue
        If ($?) {
            Write-Host -ForegroundColor Green "[SUCCESS]"
        } Else {
            Write-Host -ForegroundColor Red "[FAILURE]"
            Throw "Failed to clear proxy address"
        }
    } Else {
        Write-Host -ForegroundColor Yellow "[NO CHANGE]"
    }

    [System.Net.WebRequest]::DefaultWebProxy = $Null
}


Switch ($Action) {
    "start"     { Return (Start-ProxyEngine -ProxyIP $ProxyIP -ProxyPort $ProxyPort -WhatIf:$WhatIf) }
    "stop"      { Return (Stop-ProxyEngine -WhatIf:$WhatIf) }
    "detect"    { Return $True }
    "describe"  { Return "Windows built-in HTTP proxy support" }
    Default     { Write-Host "Unhandled action `"$Action`"" }
}
