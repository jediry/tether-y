# equires -RunAsAdministrator

Param([Switch] $WhatIf=$False,
      [String] $ConfigFile="$($Env:USERPROFILE)\.tether-y.xml")


$Script:KeyBuffer = $Null
Function Get-Key {
    $Key = $Null

    # If we previously pushed a key back, that's the key we want to return
    If ($Script:KeyBuffer -ne $Null) {
        $Key = $Script:KeyBuffer
        $Script:KeyBuffer = $Null
    } ElseIf ([Console]::KeyAvailable) {
        $Key = [Console]::readKey($True)
    }

    Return $Key
}


# Push a keystroke previous read by Get-Key back onto the stream, so that it will be read again by the next call to
# Get-Key.
Function Push-Key($Key) {
    If ($Script:KeyBuffer -ne $Null) {
        Write-Warning "Pushing multiple keystrokes back onto the stream...stuff will be lost!"
    }

    $Script:KeyBuffer = $Key
}


Function Write-WaitSpinner($Block,
                           [int] $Interval,
                           [int] $RepeatCount=0,
                           [Switch] $NoNewline=$False,
                           [Switch] $WhatIf=$False) {
    $SpinnerCount = 0
    $Spinner = "-\|/"
    Write-Host -NoNewline -ForegroundColor Magenta "-"
    $Iteration = 0
    While ($RepeatCount -eq 0 -or $Iteration -lt $RepeatCount) {

        # Run the block, and return if it succeeded
        $Result = &$Block -WhatIf:$WhatIf
        If ($Result -ne $Null) {
            # Remove the spinner (and append a newline unless told to suppress it)
            Write-Host -NoNewline:$NoNewline "`b `b"
            Return $Result
        }

        # Check for user input, and throw if we were told to abort
        $Info = Get-Key
        If ($Info -ne $Null) {
            If ($Info.Key -eq [ConsoleKey]::Escape) {
                Write-Host -ForegroundColor Red "`b[CANCEL]"
                Throw "Escape"
            }

            If ($Info.Key -eq [ConsoleKey]::Delete) {
                Write-Host -ForegroundColor Magenta "`b[RECONFIGURE]"
                Throw "Delete"
            }

            If (($Info.Modifiers -band [ConsoleModifiers]"Control") -and ($Info.key -eq 'C')) {
                Write-Host -ForegroundColor Red "`b[CANCEL]"
                Throw "Ctrl-C"
            }
        }

        # Draw the spinner
        $Char = $Spinner.Chars($SpinnerCount % $Spinner.Length)
        Write-Host -NoNewline -ForegroundColor Magenta "`b$Char"
        $SpinnerCount += 1

        # Sleep for the prescribed amount of time
        Start-Sleep -Milliseconds $Interval

        $Iteration++
    }

    Write-Host -NoNewline "`b"
    Return $Null
}


Function Read-Config([String] $Path) {

    $Config = @{
        "ProxyMAC"      = $Null;
        "ProxyPort"     = 8080;
        "ProxyEngine"   = $Null;
        "NetworkSSID"   = "TETHER-X-$($Env:COMPUTERNAME)";
        "NetworkKey"    = "1234567890";
        "Loaded"        = $False;
    }

    Write-Host -NoNewline "Reading configuration file "
    Write-Host -NoNewline -ForegroundColor Yellow $Path
    Write-Host -NoNewline "..."

    If (Test-Path -PathType "Leaf" -Path $Path) {
        Try {
            [XML] $xml = Get-Content $Path

            $Incomplete = $False

            $ProxyMAC = $xml."tether-y-config"."proxy-mac"
            If (![string]::IsNullOrEmpty($ProxyMAC)) {
                $Config.ProxyMAC = $ProxyMAC
            } Else {
                $Incomplete = $True
            }

            $ProxyPort = $xml."tether-y-config"."proxy-port"
            If (![string]::IsNullOrEmpty($ProxyPort)) {
                $Config.ProxyPort = $ProxyPort
            } Else {
                $Incomplete = $True
            }

            $ProxyEngine = $xml."tether-y-config"."proxy-engine"
            If (![string]::IsNullOrEmpty($ProxyEngine)) {
                $Config.ProxyEngine = $ProxyEngine
            } Else {
                $Incomplete = $True
            }

            $NetworkSSID = $xml."tether-y-config"."network-ssid"
            If (![string]::IsNullOrEmpty($NetworkSSID)) {
                $Config.NetworkSSID = $NetworkSSID
            } Else {
                $Incomplete = $True
            }

            $NetworkKey = $xml."tether-y-config"."network-key"
            If (![string]::IsNullOrEmpty($NetworkKey)) {
                $Config.NetworkKey = $NetworkKey
            } Else {
                $Incomplete = $True
            }

            If ($Incomplete) {
                Write-Host -ForegroundColor Cyan "config file is missing some information..."
            } Else {
                $Config.Loaded = $True
                Write-Host -ForegroundColor Green "[SUCCESS]"
            }
        } Catch {
            Write-Host -ForegroundColor Red "[FAILURE]"
        }
    } Else {
        Write-Host -ForegroundColor Cyan "[NOT PRESENT]"
    }

    Return $Config
}


Function Write-Config($Config,
                      [String] $Path,
                      [Switch] $WhatIf=$False) {
    Write-Host -NoNewline "Writing config file "
    Write-Host -NoNewline -ForegroundColor Yellow $Path
    Write-Host -NoNewline "..."

    If (!$WhatIf) {
        Try {
            $XML = @"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <tether-y-config>
                  <network-ssid>$($Config.NetworkSSID)</network-ssid>
                  <network-key>$($Config.NetworkKey)</network-key>
                  <proxy-mac>$($Config.ProxyMAC)</proxy-mac>
                  <proxy-port>$($Config.ProxyPort)</proxy-port>
                  <proxy-engine>$($Config.ProxyEngine)</proxy-engine>
                </tether-y-config>
"@
            $XML = $XML -Replace "(?m)^                ",""
            Set-Content -Path $Path -Value $XML
            Write-Host -ForegroundColor Green "[SUCCESS]"
        } Catch {
            Write-Host -ForegroundColor Red "[FAILURE]"
        }
    } Else {
        Write-Host -ForegroundColor Cyan "[SKIPPED]"
    }
}


# Ask the user for the value of a property, possibly offering a default
Function Get-UserValue([String] $Text, [String] $DefaultValue=$Null) {
    Write-Host -NoNewline $Text
    If (![String]::IsNullOrEmpty($DefaultValue)) {
        Write-Host -NoNewline " ["
        Write-Host -NoNewline -ForegroundColor Green $DefaultValue
        Write-Host -NoNewline "]"
    }
    Write-Host -NoNewline ": "
    $Input = Read-Host
    If ([String]::IsNullOrEmpty($Input)) {
        $Input = $DefaultValue
    }
    Return $Input
}


Function Get-NetworkSSID($Config) {
    If (!$Config.Loaded -or [String]::IsNullOrEmpty($Config.NetworkSSID)) {
        Do {
            $Config.NetworkSSID = Get-UserValue "Enter SSID for the Tether-X wireless network" -DefaultValue $Config.NetworkSSID
        } While ([String]::IsNullOrEmpty($Config.NetworkSSID))
    }

    Return $Config.NetworkSSID
}


Function Get-NetworkKey($Config) {
    If (!$Config.Loaded -or [String]::IsNullOrEmpty($Config.NetworkKey)) {
        Do {
            $Config.NetworkKey = Get-UserValue "Enter key for the Tether-X wireless network" -DefaultValue $Config.NetworkKey
        } While ([String]::IsNullOrEmpty($Config.NetworkKey))
    }

    Return $Config.NetworkKey
}


Function Get-ProxyIP($Config,
                     $Adapter,
                     [Switch] $WhatIf=$False) {
    If (!$Config.Loaded -or [String]::IsNullOrEmpty($Config.ProxyMAC)) {
        Write-Host -NoNewline "Searching for devices on the "
        Write-Host -NoNewline -ForegroundColor Green $Config.NetworkSSID
        Write-Host " network. Select the one that is your Tether-X device."

        If ($WhatIf) {
            $Config.ProxyMAC = "DEADBEEF0000"
            Write-Host "Proxy device is $($Config.ProxyMAC)"
            Return "0.0.0.0"
        }

        $Neighbors = @()
        Do {
            # Check (again) to see if there are any new devices
            Get-NetNeighbor -InterfaceIndex $Adapter.ifIndex -AddressFamily IPv4 | ForEach-Object {
                # Skip over some MAC addresses that couldn't possibly be the Tether-X device
                If ($_.LinkLayerAddress -eq "ffffffffffff" -or $_.LinkLayerAddress -eq "000000000000") {
                    # Ignore
                } ElseIf ($_.State -eq "Permanent") {
                    # Ignore
                } Else {
                    # OK, this looks like a possibility. Only add it to our list of options if it's not already in there
                    $Found = $False
                    ForEach ($N in $Neighbors) {
                        If ($N.LinkLayerAddress -eq $_.LinkLayerAddress) {
                            $Found = $True
                            Break
                        }
                    }

                    If (!$Found) {
                        If ($Neighbors.Count -lt 9) {
                            Write-Host "`t[$($Neighbors.Count)]: $($_.IPAddress)`t$($_.LinkLayerAddress)`t$($_.State)"
                            $Neighbors += $_
                        } Else {
                            Write-Host "Uh oh...more than 10 neighbors on the hosted network? This code won't work..."
                        }
                    }
                }
            }

            # Give the user a chance to select one
            $Proxy = Write-WaitSpinner -Interval 100 -RepeatCount 10 -NoNewline {
                $Key = Get-Key
                If ($Key -ne $Null) {
                    If ([char]::IsDigit($Key.KeyChar)) {
                        $Index = [int32]::parse($Key.KeyChar)
                        If ($Index -lt $Neighbors.Count) {
                            Return $Neighbors[ $Index ]
                        }
                    } Else {
                        Push-Key $Key
                    }
                }
            }

            If ($Proxy -ne $Null) {
                $Config.ProxyMAC = $Proxy.LinkLayerAddress
                Write-Host "Proxy is $($Config.ProxyMAC)"
                Return $Proxy.IPAddress
            }

        } While ($True)
    }

    Write-Host -NoNewline "Waiting for Tether-X proxy (MAC address "
    Write-Host -NoNewline -ForegroundColor Yellow $Config.ProxyMAC
    Write-Host -NoNewline ") to connect..."
    $IPAddress = Write-WaitSpinner -Interval $Interval -NoNewline -WhatIf:$WhatIf {
        Param([Switch] $WhatIf)

        If (!$WhatIf) {
            Get-IPAddressForProxy $Config.ProxyMAC
        } Else {
            Return "0.0.0.0"
        }
    }
    Write-Host -ForegroundColor Cyan $IPAddress

    Return $IPAddress
}


Function Get-ProxyPort($Config) {
    If (!$Config.Loaded -or [String]::IsNullOrEmpty($Config.ProxyPort)) {
        Do {
            $Config.ProxyPort = Get-UserValue "Enter network port for the the Tether-X proxy" -DefaultValue $Config.ProxyPort
        } While ([String]::IsNullOrEmpty($Config.ProxyPort))
    }

    Return $Config.ProxyPort
}


Function Get-ProxyEngine($Config) {
    If (!$Config.Loaded -or [String]::IsNullOrEmpty($Config.ProxyEngine)) {
        $Engines = @( )
        $DefaultIndex = -1
        Get-Item -Path ProxyEngines\*.ps1 | ForEach-Object {
            $Description = . $_ -Action "describe"
            $Available = . $_ -Action "detect"
            $Name = ($_ -Replace '\.ps1','') -Replace '^.*\\',''
            If ($Name -eq $Config.ProxyEngine) {
                If ($Available) {
                    Write-Host -NoNewline -ForegroundColor Green "`t[$($Engines.Count)] "
                    $DefaultIndex = $Engines.Count
                } Else {
                    # Hmmm...default engine is now unavailable?
                    Write-Host -NoNewline "`t[$($Engines.Count)] "
                    $Config.ProxyEngine = $Null
                }
            } Else {
                Write-Host -NoNewline "`t[$($Engines.Count)] "
            }

            If ($Available) {
                Write-Host -ForegroundColor White $Description
            } Else {
                Write-Host -ForegroundColor DarkGray $Description
            }
            $Engines += @{ "Name" = $Name; "Description" = $Description; "Available" = $Available; }
        }

        Do {
            $Index = Get-UserValue "Select which proxy engine you wish to use for routing traffic through Tether-X" -DefaultValue $DefaultIndex
            If ( $Index -ne $Null) {
                $Index = [int] $Index
                If ($Index -ge 0 -and $Index -lt $Engines.Count) {
                    $Engine = $Engines[ $Index ]
                    If ($Engine.Available) {
                        $Config.ProxyEngine = $Engines[ $Index ].Name
                    } Else {
                        # TODO: provide URL to acquire
                        Write-Host "This proxy engine is not currently installed on your system."
                    }
                }
            }
        } While ([String]::IsNullOrEmpty($Config.ProxyEngine))
    }

    Return $Config.ProxyEngine
}


Function Start-HostedNetwork($SSID, $Key,
                             [Switch] $WhatIf=$False) {
    Write-Host -NoNewline -ForegroundColor Cyan "Enabling "
    Write-Host -NoNewline "the hosted wifi network "
    Write-Host -NoNewline -ForegroundColor Yellow $SSID
    Write-Host -NoNewline " with key "
    Write-Host -NoNewline -ForegroundColor Yellow $Key
    Write-Host -NoNewline "..."

    If (!$WhatIf) {
        netsh wlan set hosted mode=allow ssid=$SSID key=$Key > $Nul
        If (-not $?) {
            Write-Host -ForegroundColor Red "[FAILURE]"
            Throw
        }
    }

    If (!$WhatIf) {
        netsh wlan start hosted > $Nul
        If (-not $?) {
            Write-Host -ForegroundColor Red "[FAILURE]"
            Throw
        }
    }

    If (!$WhatIf) {
        $Adapter = Get-NetAdapter -InterfaceDescription "Microsoft Hosted Network Virtual Adapter" -ErrorAction:SilentlyContinue
        If ($Adapter -eq $Null) {
            Write-Host -ForegroundColor Red "[FAILURE]"
        }

        # Clear old ARP cache entries for this interface
        Remove-NetNeighbor -Confirm:$False -InterfaceIndex $Adapter.ifIndex -State "Stale", "Unreachable", "Incomplete", "Probe", "Delay", "Reachable" -ErrorAction:SilentlyContinue
    } Else {
        $Adapter = $True
    }

    If (!$WhatIf) {
        Write-Host -ForegroundColor Green "[SUCCESS]"
    } Else {
        Write-Host -ForegroundColor Cyan "[SKIPPED]"
    }

    Return $Adapter
}


Function Stop-HostedNetwork([Switch] $WhatIf=$False) {
    Write-Host -NoNewline -ForegroundColor Cyan "Disabling "
    Write-Host -NoNewline "the hosted wifi network..."

    If (!$WhatIf) {
        netsh wlan stop hosted > $Nul
        If (-not $?) {
            Write-Host -ForegroundColor Red "[FAILURE]"
            Throw
        }

        Write-Host -ForegroundColor Green "[SUCCESS]"
    } Else {
        Write-Host -ForegroundColor Cyan "[SKIPPED]"
    }
}


Function Start-Proxy([string] $IPAddress,
                     [string] $Port,
                     [string] $Engine,
                     [Switch] $WhatIf=$False) {

    Write-Host -NoNewline "Setting proxy server address "
    Write-Host -NoNewline -ForegroundColor Yellow $ProxyString
    Write-Host -NoNewline "..."

    # Addresses to exclude from proxying
    $Excludes = @( "localhost", "127.0.0.1", "$Env:ComputerName" )

    . ProxyEngines\$Engine.ps1 -Action "start" -ProxyIP $IPAddress -ProxyPort $Port -ExcludeAddresses $Excludes -WhatIf:$WhatIf
}


Function Stop-Proxy([string] $Engine,
                    [switch] $WhatIf=$False) {
    Write-Host -NoNewline "Disabling proxy..."
    . ProxyEngines\$Engine.ps1 -Action "stop" -WhatIf:$WhatIf
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


Function Wait-Proxy($ProxyMAC, $Interval, [Switch] $WhatIf=$False) {
    Write-Host -NoNewline "Waiting for Tether-X proxy (MAC address "
    Write-Host -NoNewline -ForegroundColor Yellow $ProxyMAC
    Write-Host -NoNewline ") to connect..."
    $IPAddress = Write-WaitSpinner -Interval $Interval -NoNewline -WhatIf:$WhatIf {
        Param([Switch] $WhatIf)

        If (!$WhatIf) {
            Get-IPAddressForProxy $ProxyMAC
        } Else {
            Return "0.0.0.0"
        }
    }
    Write-Host -ForegroundColor Cyan $IPAddress
    Return $IPAddress
}


Function Wait-Internet($Interval, [Switch] $WhatIf=$False) {
    Write-Host -NoNewline "Tether-X Proxy set. Checking Internet connectivity..."
    $Result = Write-WaitSpinner -Interval $Interval -WhatIf:$WhatIf -NoNewline {
        Param([Switch] $WhatIf)

        If (!$WhatIf) {
            Try {
                $Request = [System.Net.WebRequest]::Create("http://internetbeacon.msedge.net")
                $Response = $Request.GetResponse()
                If ($Response.StatusCode -eq 200) {
                    Return $True
                } Else {
                    Write-Host "Response: $($Response.StatusDescription)"
                }
            } Catch {
#                Write-Host "Ack! $_"
            }

            Return $Null

        } Else {
            Return $True
        }
    }

    If ($Result) {
        Write-Host -ForegroundColor Green "[SUCCESS]"
    }
}


Function Wait-Disconnect($ProxyMAC, $Interval, [Switch] $WhatIf=$False) {
    Write-Host -NoNewline "Tethered link established...press "
    Write-Host -NoNewline -ForegroundColor Cyan "Esc "
    Write-Host -NoNewline "to disconnect..."
    Write-WaitSpinner -Interval $Interval -WhatIf:$WhatIf -NoNewline {
        Param([Switch] $WhatIf)

        If (!$WhatIf) {
            Return $Null
#            $IP = Get-IPAddressForProxy $ProxyMAC
#            If ($IP -ne $Null) { Return $Null }
#            Else { Return 1 }
        } Else {
            Return $True
        }
    } > $Nul

    Write-Host -ForegroundColor Cyan "[DISCONNECT]"
}



#### Main Script ####

# Save this so we can restore it later
If (!$WhatIf) {
    $SaveTreatControlCAsInput = [Console]::TreatControlCAsInput
    [Console]::TreatControlCAsInput = $True
}

# Load any existing configuration file (or defaults if the file is not present) 
$Config = Read-Config $ConfigFile

$ProxyEngine = $Null
$RetrySetup = $True
While ($RetrySetup) {
    Try {
        If ($Config.Loaded) {
            Write-Host -NoNewline "Using configuration from "
            Write-Host -NoNewline -ForegroundColor Yellow $ConfigFile
            Write-Host -NoNewline ". Press "
            Write-Host -NoNewline -ForegroundColor Magenta "Delete"
            Write-Host " to cancel and reconfigure.`n"
        } Else {
            Write-Host -NoNewline "`nReconfiguring. Press "
            Write-Host -NoNewline -ForegroundColor Cyan "Enter"
            Write-Host " to accept defaults.`n"
        }

        # Figure out what SSID and privacy key to use
        $NetworkSSID = Get-NetworkSSID -Config $Config
        $NetworkKey = Get-NetworkKey -Config $Config

        # Check the status of the hosted wireless network, and start if it's not running
        $Adapter = Start-HostedNetwork -SSID $NetworkSSID -Key $NetworkKey -WhatIf:$WhatIf

        # Decide which proxy engine we're using
        $ProxyEngine = Get-ProxyEngine -Config $Config

        # Wait for the proxy to connect to the hosted network (based on the MAC address configured above)
        $ProxyIP = Get-ProxyIP -Config $Config -Adapter $Adapter -WhatIf:$WhatIf

        $ProxyPort = Get-ProxyPort -Config $Config

        # Set this as our proxy server
        Start-Proxy -IPAddress $ProxyIP -Port $ProxyPort -Engine $ProxyEngine -WhatIf:$WhatIf

        # Verify that the Internet is reachable
        Wait-Internet -Interval 1000 -WhatIf:$WhatIf

        # Wait for the proxy to disappear, or for the user to tell us to disconnect
        Wait-Disconnect -Interval 5000 $Config.ProxyMAC -WhatIf:$WhatIf

    } Catch [System.Management.Automation.RuntimeException] {

        $Key = $_.ToString()
        If ($Key -eq "Delete") {
            Write-Host -ForegroundColor Cyan "`nReconfiguring"
            $Config.Loaded = $False
        } ElseIf ($Key -eq "Escape" -or $Key -eq "Ctrl-C") {
            Write-Host -ForegroundColor Magenta "`nShutting down"
            $RetrySetup = $False
        } Else {
            Write-Host -ForegroundColor Red "`nUnexpected input [$Key]...shutting down"
            $RetrySetup = $False
        }

    } Finally {

        # Clear the proxy if we set it
        If ($ProxyEngine -ne $Null) {
            Stop-Proxy -Engine $ProxyEngine -WhatIf:$WhatIf
            $ProxyEngine = $Null
        }

        # Shut down the hosted network
        Stop-HostedNetwork -WhatIf:$WhatIf
    }
}

# Write the current config to the config file
Write-Config -Path $ConfigFile -Config $Config -WhatIf:$WhatIf

# Restore Ctrl-C handling
If (!$WhatIf) {
    # Put Ctrl-C handling back the way we found it
    [Console]::TreatControlCAsInput = $SaveTreatControlCAsInput
}
