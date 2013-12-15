# equires -RunAsAdministrator


$SSID = "Tether-X-$($Env:COMPUTERNAME)"
$Password = "12345678"
$ProxyMAC = "4C:25:78:52:5B:B4"
$ProxyPort = "8080"


Function Write-WaitSpinner($Block, $Interval, [Switch] $NoNewline=$False) {
    $SpinnerCount = 0
    $Spinner = "-\|/"
    Write-Host -NoNewline -ForegroundColor Magenta "-"
    While ($True) {

        # Run the block, and return if it succeeded
        $Result = &$Block
        If ($Result -ne $Null) {
            # Remove the spinner and append a newline unless told to suppress it
            Write-Host -NoNewline:$NoNewline "`b"
            Return $Result
        }

        # Check for user input, and throw if we were told to abort
        If ([Console]::KeyAvailable) {
            $Info = [Console]::readKey($True)
            If ($Info.Key -eq [ConsoleKey]::Escape) {
                Write-Host
                Throw "Escape"
            }

            If (($Info.Modifiers -band [ConsoleModifiers]"Control") -and ($Info.key -eq 'C')) {
                Write-Host
                Throw "Ctrl-C"
            }
        }

        # Draw the spinner
        $Char = $Spinner.Chars($SpinnerCount % $Spinner.Length)
        Write-Host -NoNewline -ForegroundColor Magenta "`b$Char"
        $SpinnerCount += 1

        # Sleep for the prescribed amount of time
        Start-Sleep -Milliseconds $Interval
    }
}


Function Start-HostedNetwork($SSID, $Key) {
    Write-Host -NoNewline -ForegroundColor Cyan "Enabling "
    Write-Host -NoNewline "the hosted wifi network "
    Write-Host -NoNewline -ForegroundColor Yellow $SSID
    Write-Host -NoNewline "..."

    netsh wlan set hosted mode=allow ssid=$SSID key=$Key > $Nul
    If (-not $?) {
        Write-Host -ForegroundColor Red "[FAILURE]"
        Throw
    }

    netsh wlan start hosted > $Nul
    If (-not $?) {
        Write-Host -ForegroundColor Red "[FAILURE]"
        Throw
    }

    $Adapter = Get-NetAdapter -InterfaceDescription "Microsoft Hosted Network Virtual Adapter" -ErrorAction:SilentlyContinue
    If ($Adapter -eq $Null) {
        Write-Host -ForegroundColor Red "[FAILURE]"
    }

    Write-Host -ForegroundColor Green "[SUCCESS]"

    # Clear old ARP cache entries for this interface
    Remove-NetNeighbor -Confirm:$False -InterfaceIndex $Adapter.ifIndex -State "Stale", "Unreachable", "Incomplete", "Probe", "Delay", "Reachable" -ErrorAction:SilentlyContinue
    Return $Adapter
}


Function Stop-HostedNetwork {
    Write-Host -NoNewline -ForegroundColor Cyan "Disabling "
    Write-Host -NoNewline "the hosted wifi network..."

    netsh wlan stop hosted > $Nul
    If (-not $?) {
        Write-Host -ForegroundColor Red "[FAILURE]"
        Throw
    }

    Write-Host -ForegroundColor Green "[SUCCESS]"
}


$Script:ProxyRegKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings"

Function Set-Proxy($IPAddress, $Port) {

    $ProxyString = "$IPAddress`:$Port"
    Write-Host -NoNewline "Setting proxy server address "
    Write-Host -NoNewline -ForegroundColor Yellow $ProxyString
    Write-Host -NoNewline "..."
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

    Write-Host -NoNewline "Enabling proxy..."
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

    [System.Net.WebRequest]::DefaultWebProxy = [System.Net.WebProxy]$ProxyString
}


Function Clear-Proxy {
    Write-Host -NoNewline "Disabling proxy..."
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


Function Get-IPAddressForProxy($MAC) {
    $Entry = Get-NetNeighbor -LinkLayerAddress $MAC.Replace(":", "") -ErrorAction:SilentlyContinue
    If ($Entry -eq $Null) {
        Return $Null
    }

    If ($Entry.State -eq "Reachable") {
        Return $Entry.IPAddress
    } ElseIf ($Entry.State -eq "Stale") {
#        Write-Host "Sending ping"
        ping $Entry.IPAddress -n 1 > $Nul
    } ElseIf ($Entry.State -ne "Probe") {
        Write-Host "State: $($Entry.State)"
#    } Else {
#        Write-Host "State: $($Entry.State)"
    }

    Return $Null # Keep waiting
}


Function Wait-Proxy($ProxyMAC, $Interval) {
    Write-Host -NoNewline "Waiting for Tether-X proxy (MAC address "
    Write-Host -NoNewline -ForegroundColor Yellow $ProxyMAC
    Write-Host -NoNewline ") to connect..."
    $IPAddress = (Write-WaitSpinner { Get-IPAddressForProxy $ProxyMAC } -Interval $Interval -NoNewline)
    Write-Host -ForegroundColor Cyan $IPAddress
    Return $IPAddress
}


Function Wait-Internet($Interval) {
    Write-Host -NoNewline "Tether-X Proxy set. Checking Internet connectivity..."
    Write-WaitSpinner -Interval $Interval {
        Try {
            $Request = [System.Net.WebRequest]::Create("http://internetbeacon.msedge.net")
            $Response = $Request.GetResponse()
            If ($Response.StatusCode -eq 200) {
                Return $True
            } Else {
                Write-Host "Response: $($Response.StatusDescription)"
            }
        } Catch {
#            Write-Host "Ack! $_"
        }

        Return $Null
    } > $Nul
}


Function Wait-Disconnect($ProxyMAC, $Interval) {
    Write-Host -NoNewline "Tethered link established...press "
    Write-Host -NoNewline -ForegroundColor Cyan "Esc "
    Write-Host -NoNewline "to disconnect..."
    Write-WaitSpinner -Interval $Interval {
        Return $Null
#        $IP = Get-IPAddressForProxy $ProxyMAC
#        If ($IP -ne $Null) { Return $Null }
#        Else { Return 1 }
    } > $Nul
}



#### Main Script ####

# Save this so we can restore it later
$SaveTreatControlCAsInput = [Console]::TreatControlCAsInput
[Console]::TreatControlCAsInput = $True

Try {

    # Check the status of the hosted wireless network, and start if it's not running
    $Adapter = Start-HostedNetwork -SSID $SSID -Key $Password

    # Wait for the proxy to connect to the hosted network (based on the MAC address configured above)
    $ProxyIP = Wait-Proxy $ProxyMAC -Interval 150

    # Set this as our proxy server
    Set-Proxy -IPAddress $ProxyIP -Port $ProxyPort

    # Verify that the Internet is reachable
    Wait-Internet -Interval 1000

    # Wait for the proxy to disappear, or for the user to tell us to disconnect
    Wait-Disconnect -Interval 5000 $ProxyMAC

} Catch [System.Management.Automation.RuntimeException] {

    Write-Host -NoNewline -ForegroundColor Magenta "[$_]"
    Write-Host "...Disconnecting"

} Finally {

    # Clear the proxy if we set it
    Clear-Proxy

    # Shut down the hosted network
    Stop-HostedNetwork

    # Put Ctrl-C handling back the way we found it
    [Console]::TreatControlCAsInput = $SaveTreatControlCAsInput

}
