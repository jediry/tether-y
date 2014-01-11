Param([string] $Action,
      [string] $ProxyIP = $Null,
      [string] $ProxyPort = $Null,
      $ExcludeAddresses = $Null,
      [switch] $WhatIf = $False)


Function Get-ProxifierPath {
    $ProgramFiles_x86 = (Get-Item "Env:\ProgramFiles(x86)").Value
    $Path = "$ProgramFiles_x86\Proxifier\Proxifier.exe"
    If (Get-Item -LiteralPath $Path -ErrorAction:SilentlyContinue) {
        Return $Path
    }

    $Path = "$($Env:ProgramFiles)\Proxifier\Proxifier.exe"
    If (Get-Item -LiteralPath $Path -ErrorAction:SilentlyContinue) {
        Return $Path
    }

    Write-Host "Proxifier is not installed"
    Return $Null
}


Function Get-ProxifierProfilePath {
    Return "$($Env:TEMP)\tether-y.ppx"
}


Function Write-ProxifierProfile([string] $ProxyIP,
                                [string] $ProxyPort,
                                $ExcludeAddresses,
                                [string] $Path) {
    $XML = @"
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <ProxifierProfile version="101" platform="Windows" product_id="0" product_minver="310">
          <Options>
            <Resolve>
              <AutoModeDetection enabled="true" />
              <ViaProxy enabled="true">
                <TryLocalDnsFirst enabled="false" />
              </ViaProxy>
              <ExclusionList>%ComputerName%; localhost; *.local</ExclusionList>
            </Resolve>
            <Encryption mode="basic" />
            <HttpProxiesSupport enabled="false" />
            <HandleDirectConnections enabled="false" />
            <ConnectionLoopDetection enabled="true" />
            <ProcessServices enabled="false" />
            <ProcessOtherUsers enabled="false" />
          </Options>
          <ProxyList>
            <Proxy id="100" type="HTTPS">
              <Address>$ProxyIP</Address>
              <Port>$ProxyPort</Port>
              <Options>48</Options>
            </Proxy>
          </ProxyList>
          <ChainList />
          <RuleList>
            <Rule enabled="true">
              <Name>Localhost</Name>
              <Targets>$($Excludes -join "; ")</Targets>
              <Action type="Direct" />
            </Rule>
            <Rule enabled="true">
              <Name>Default</Name>
              <Action type="Proxy">100</Action>
            </Rule>
          </RuleList>
        </ProxifierProfile>
"@
    $XML = $XML -Replace "(?m)^        ",""

    Set-Content -Path $Path -Value $XML
}


Function Start-ProxyEngine([string] $ProxyIP,
                           [string] $ProxyPort,
                           $ExcludeAddresses,
                           [switch] $WhatIf) {

    $Profile = Get-ProxifierProfilePath
    Write-ProxifierProfile -Path $Profile -ProxyIP $ProxyIP -ProxyPort $ProxyPort -ExcludeAddresses $ExcludeAddresses
    If (!$WhatIf) {
        $Proxifier = Get-ProxifierPath
        & $Proxifier $Profile silent-load
        Write-Host -ForegroundColor Green "[SUCCESS]"
    } Else {
        Write-Host -ForegroundColor Cyan "[SKIPPED]"
    }
}


Function Stop-ProxyEngine([switch] $WhatIf) {
    If (!$WhatIf) {
        Try {
            $Proxifier = Get-Process -Name "Proxifier" -ErrorAction:SilentlyContinue
            If ($Proxifier -ne $Null) {
                Stop-Process $Proxifier
                Write-Host -ForegroundColor Green "[SUCCESS]"
            } Else {
                Write-Host -ForegroundColor Yellow "[NO CHANGE]"
            }
        } Catch {
            Write-Host -ForegroundColor Red "[FAILURE]"
            Write-Host "Error: $_"
        }
    } Else {
        Write-Host -ForegroundColor Cyan "[SKIPPED]"
    }
}


Switch ($Action) {
    "start"     { Return (Start-ProxyEngine -ProxyIP $ProxyIP -ProxyPort $ProxyPort -ExcludeAddresses $ExcludeAddresses -WhatIf:$WhatIf) }
    "stop"      { Return (Stop-ProxyEngine -WhatIf:$WhatIf) }
    "detect"    { Return (Get-ProxifierPath -ne $Null) }
    "describe"  { Return 'Proxifier: Winsock plugin; proxies all kinds of network traffic ($40 at www.proxifier.com)' }
    Default     { Write-Host "Unhandled action `"$Action`"" }
}
