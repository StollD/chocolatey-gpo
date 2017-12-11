# chocolatey-gpo - Linking the chocolatey package manager with Active Directory Group Policies
# Copyright (c) 2017 Dorian Stoll
# Licensed under the Terms of the MIT License

function Get-RegValue([String] $KeyPath, [String] $ValueName) {
    (Get-ItemProperty -LiteralPath $KeyPath -Name $ValueName).$ValueName
}
function Get-RegValues([String] $KeyPath) {
    $RegKey = (Get-ItemProperty $KeyPath)
    $RegKey.PSObject.Properties | 
        Where-Object { $_.Name -ne "PSPath" -and $_.Name -ne "PSParentPath" -and $_.Name -ne "PSChildName" -and $_.Name -ne "PSDrive" -and $_.Name -ne "PSProvider" } | 
        ForEach-Object {
        $_.Name
    }
}
function Call-CMD([String] $command) {
    cmd.exe /C $command
}

#region Install Boxstarter & Chocolatey

Invoke-Expression ((New-Object System.Net.WebClient).DownloadString("http://boxstarter.org/bootstrapper.ps1")); Get-Boxstarter -Force
if (Test-Path "C:\Users\Public\Desktop\Boxstarter Shell.lnk") { 
    Remove-Item -Path "C:\Users\Public\Desktop\Boxstarter Shell.lnk" -Force
}

#endregion

#region Install Updates

Get-RegValues "HKLM:\SOFTWARE\Policies\Chocolatey" | ForEach-Object {
    if ($_ -eq "installUpdates") {
        $mode = Get-RegValue "HKLM:\SOFTWARE\Policies\Chocolatey" $_
        if ($mode -eq 1) {
            Install-WindowsUpdate -Full -AcceptEula -SupressReboots
        }
    }
}
Update-ExecutionPolicy Unrestricted

#endregion

#region Apply Config

Get-RegValues "HKLM:\SOFTWARE\Policies\Chocolatey\Config" | ForEach-Object {
    if (!($_ -match ".+?Mode$")) {
        $value = Get-RegValue "HKLM:\SOFTWARE\Policies\Chocolatey\Config" $_
        $mode = Get-RegValue "HKLM:\SOFTWARE\Policies\Chocolatey\Config" ($_ + "Mode")
        
        if ($mode -eq 1) {
            # set
            Call-CMD "choco config set --name $_ --value $value"
        }
        else {
            # unset
            Call-CMD "choco config unset --name $_"
        }
    }
}

#endregion

#region Apply Features

Get-RegValues "HKLM:\SOFTWARE\Policies\Chocolatey\Features" | ForEach-Object { 
    $mode = Get-RegValue "HKLM:\SOFTWARE\Policies\Chocolatey\Features" $_
    
    if ($mode -eq 1) {
        # enable
        Call-CMD "choco feature enable --name $_"
    }
    elseif ($mode -eq 0) {
        # disable
        Call-CMD "choco feature disable --name $_"
    }
}

#endregion

#region Package Sources

Get-RegValues "HKLM:\SOFTWARE\Policies\Chocolatey\Sources" | ForEach-Object {
    $value = Get-RegValue "HKLM:\SOFTWARE\Policies\Chocolatey\Sources" $_
    if ($value -eq "remove") {
        Call-CMD "choco sources remove -n $_"
    }
    else {
        Call-CMD "choco sources add -n $_ -s $value"
    }
}

#endregion

#region Install Packages

Get-RegValues "HKLM:\SOFTWARE\Policies\Chocolatey\Packages" | ForEach-Object {
    $params = Get-RegValue "HKLM:\SOFTWARE\Policies\Chocolatey\Packages" $_
    Write-Host $params
    if ($params -eq "remove") {
        Call-CMD "cuninst $_ -y"
    }
    else {
        Call-CMD "choco install $_ -y $params"
        Call-CMD "choco upgrade $_ -y $params"
    }
}

#endregion