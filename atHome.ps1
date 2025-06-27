#region Helper functions

function Is-VirtualMachine {
    $cs = Get-CimInstance -ClassName Win32_ComputerSystem
    $manufacturer = $cs.Manufacturer
    $model = $cs.Model

    $vmIndicators = @("Microsoft Corporation", "VMware", "VirtualBox", "QEMU", "Xen", "KVM", "Parallels")

    foreach ($vendor in $vmIndicators) {
        if ($manufacturer -like "*$vendor*" -or $model -like "*$vendor*") {
            return $true
        }
    }

    return $false
}

function Is-DomainJoined {
    $computerSystem = Get-CimInstance -Class Win32_ComputerSystem
    return $computerSystem.PartOfDomain
}

function Is-Laptop {
    # Check for internal battery
    $battery = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue
    if ($battery) {
        return $true
    }

    # Check chassis type for laptop indicators
    $chassis = Get-CimInstance -ClassName Win32_SystemEnclosure
    $laptopTypes = @(8, 9, 10, 14)  # Portable, Laptop, Notebook, SubNotebook

    foreach ($type in $chassis.ChassisTypes) {
        if ($laptopTypes -contains $type) {
            return $true
        }
    }

    return $false
}

function Get-WindowsMajorVersion {
    $version = (Get-CimInstance -ClassName Win32_OperatingSystem).Version
    $build = [int]($version.Split('.')[2])

    if ($build -ge 22000) {
        return 11
    } else {
        return 10
    }
}

function Get-WindowsEdition {
    $releaseId = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name "DisplayVersion" -ErrorAction SilentlyContinue
    if (-not $releaseId) {
        # Fallback for older Windows 10 versions
        $releaseId = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name "ReleaseId" -ErrorAction SilentlyContinue
    }

    if ($releaseId) {
        return $releaseId
    } else {
        return "Unknown"
    }
}


#endregion

#region Main Functions
function Set-KeyboardLanguages {
    # Define desired layouts
    $desiredLayouts = @("0409:00000409", "0418:00010418") # English (US), Romanian Standard

    # Get current input methods
    $currentLayouts = Get-WinUserLanguageList

    # Remove all existing input methods
    $currentLayouts.Clear()

    # Add English (US) and Romanian (Standard)
    $enUS = New-WinUserLanguageList en-US
    $roRO = New-WinUserLanguageList ro-RO
    $enUS[0].InputMethodTips.Clear()
    $roRO[0].InputMethodTips.Clear()
    $enUS[0].InputMethodTips.Add("0409:00000409")
    $roRO[0].InputMethodTips.Add("0418:00010418")

    # Combine and set
    $finalList = $enUS + $roRO
    Set-WinUserLanguageList $finalList -Force

    # Set English (US) as default input method
    Set-WinUILanguageOverride -Language "en-US"
    Set-WinUserLanguageList -LanguageList $finalList -Force
    Set-WinSystemLocale en-US
    Set-Culture en-US
    Set-WinHomeLocation -GeoId 244 # United States
}

function Remove-TaskbarCrap {
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "SearchboxTaskbarMode" -PropertyType DWord -Value 0 -Force | Out-Null

    New-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowTaskViewButton" -PropertyType DWord -Value 0 -Force | Out-Null
}

function Set-DarkModeOn {
    # Set Windows mode and App mode to Dark
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" `
                    -Name "AppsUseLightTheme" -Type DWord -Value 0 -Force | Out-Null

    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" `
                    -Name "SystemUsesLightTheme" -Type DWord -Value 0 -Force | Out-Null
}

function Set-WindowsAccentColorDefaultBlue {
    $accentColor = 0xD77800  # Default blue in BGR

    # Set accent color values
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\DWM" `
                     -Name "ColorizationColor" -Type DWord -Value $accentColor -Force

    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Accent" `
                     -Name "AccentColorMenu" -Type DWord -Value $accentColor -Force

    # Ensure accent is NOT applied to Start/taskbar and title bars
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" `
                     -Name "ColorPrevalence" -Type DWord -Value 0 -Force

    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\DWM" `
                     -Name "ColorPrevalence" -Type DWord -Value 0 -Force
}

function Center-StartButton {
    $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    $name = "TaskbarAl"
    
    # Only proceed if Windows 11 (build 22000+)
    $osVersion = [System.Environment]::OSVersion.Version
    if ($osVersion.Major -eq 10 -and $osVersion.Build -ge 22000) {
        $current = Get-ItemProperty -Path $regPath -Name $name -ErrorAction SilentlyContinue

        if ($null -ne $current -and $current.$name -eq 0) {
            Set-ItemProperty -Path $regPath -Name $name -Value 1 -Force
            Stop-Process -Name explorer -Force
        }
        Write-Host "Start button - done"
    }
    else {
        Write-Host "Start button - skipped"
    }
}

function Enable-ShowExtensionsAndHiddenFiles {
    $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"

    # Show file extensions
    Set-ItemProperty -Path $regPath -Name "HideFileExt" -Value 0 -Force

    # Show hidden files and folders
    Set-ItemProperty -Path $regPath -Name "Hidden" -Value 1 -Force
    
}

function Remove-MusicAndVideosFromQuickAccess {
    $quickAccessPins = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Quick Access"

    # Known Folder GUIDs
    $foldersToRemove = @(
        "{4BD8D571-6D19-48D3-BE97-422220080E43}", # Music
        "{18989B1D-99B5-455B-841C-AB7C74E4DDFC}"  # Videos
    )

    foreach ($guid in $foldersToRemove) {
        $path = Join-Path $quickAccessPins $guid
        if (Test-Path $path) {
            Remove-Item -Path $path -Recurse -Force
        }
    }

    # Clear Quick Access cache (optional, forces rebuild)
    Remove-Item "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations\f01b4d95cf55d32a.automaticDestinations-ms" -Force -ErrorAction SilentlyContinue
}

function Set-ExplorerToOpenThisPC {
    $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    
    # 0 = Quick access, 1 = This PC
    Set-ItemProperty -Path $regPath -Name "LaunchTo" -Value 1 -Force
}

function Disable-NotepadPlusPlusRememberLastSession {
    $configPath = "$env:APPDATA\Notepad++\config.xml"

    if (-not (Test-Path $configPath)) {
        Write-Host "Notepad++ - skipped"
        return
    }

    [xml]$xml = Get-Content $configPath

    $node = $xml.SelectSingleNode("/NotepadPlus/GUIConfigs/GUIConfig[@name='RememberLastSession']")

    if ($node -ne $null) {
        $node.InnerText = "no"
    } else {
        # Create <GUIConfig name="RememberLastSession">no</GUIConfig>
        $guiConfigsNode = $xml.SelectSingleNode("/NotepadPlus/GUIConfigs")
        if (-not $guiConfigsNode) {
            $guiConfigsNode = $xml.CreateElement("GUIConfigs")
            $xml.DocumentElement.AppendChild($guiConfigsNode) | Out-Null
        }

        $newNode = $xml.CreateElement("GUIConfig")
        $newNode.SetAttribute("name", "RememberLastSession")
        $newNode.InnerText = "no"
        $guiConfigsNode.AppendChild($newNode) | Out-Null
    }

    $xml.Save($configPath)
    Write-Host "Notepad++ - done"
}

function Disable-KeyboardCrap {
    # Registry path
    $basePath = "HKCU:\Control Panel\Accessibility"

    # Disable Sticky Keys
    Set-ItemProperty "$basePath\StickyKeys" -Name "Flags" -Value 506
    Set-ItemProperty "$basePath\StickyKeys" -Name "HotKeyActive" -Value 0
    Set-ItemProperty "$basePath\StickyKeys" -Name "HotkeySound" -Value 0
    Set-ItemProperty "$basePath\StickyKeys" -Name "ConfirmActivation" -Value 0

    # Disable Toggle Keys
    Set-ItemProperty "$basePath\ToggleKeys" -Name "Flags" -Value 58
    Set-ItemProperty "$basePath\ToggleKeys" -Name "HotKeyActive" -Value 0
    Set-ItemProperty "$basePath\ToggleKeys" -Name "HotkeySound" -Value 0
    Set-ItemProperty "$basePath\ToggleKeys" -Name "ConfirmActivation" -Value 0

    # Ensure both features are off in the master Accessibility section
    Set-ItemProperty "$basePath" -Name "StickyKeys" -Value "0"
    Set-ItemProperty "$basePath" -Name "ToggleKeys" -Value "0"

}

function Set-TaskbarCombineAlways {
    $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    
    Set-ItemProperty -Path $regPath -Name "TaskbarGlomLevel" -Value 0 -Force
}

function Set-Win11StartMenuPreferences {
    # 1. More pins layout
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" `
        -Name "Start_Layout" -Type DWord -Value 1 -Force

    # 2. Disable recommended/recent items
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" `
        -Name "Start_TrackDocs" -Type DWord -Value 0 -Force

    # 3. Disable tips/shortcuts/new apps
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" `
        -Name "SubscribedContent-338393Enabled" -Type DWord -Value 0 -Force

    # 4. Enable recently added apps
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" `
        -Name "Start_TrackProgs" -Type DWord -Value 1 -Force

    # 5. Show Downloads and Settings next to power button
    $path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"
    if (-not (Test-Path $path)) {
        New-Item -Path $path -Force | Out-Null
    }

    $downloadsGuid = '{374DE290-123F-4565-9164-39C4925E467B}'
    $settingsGuid  = '{F1B32785-6FBA-4FCF-9D55-7B8E7F157091}'

    Set-ItemProperty -Path $path -Name $downloadsGuid -Type DWord -Value 0 -Force
    Set-ItemProperty -Path $path -Name $settingsGuid -Type DWord -Value 0 -Force

    Write-Host "Win11 Start Menu - done"

}

function Enable-ClipboardHistory {
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Clipboard" `
                     -Name "EnableClipboardHistory" `
                     -Type DWord -Value 1 -Force
}

function Disable-WidgetsButton {
    # Check if the OS is Windows 11
    $osVersion = (Get-CimInstance -ClassName Win32_OperatingSystem).Version
    $buildNumber = [int]($osVersion.Split('.')[2])

    if ($buildNumber -lt 22000) {
        Write-Output "Widgets - skipped"
        return
    }

    # Registry path for Widgets taskbar setting
    $regPath = "HKLM:\SOFTWARE\Policies\Microsoft\Dsh"

    # Create key if it doesn't exist
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }

    # Set the value to disable widgets
    New-ItemProperty -Path $regPath -Name "AllowNewsAndInterests" -PropertyType DWord -Value 0 -Force | Out-Null

    Write-Output "Widgets - done"
}

function Disable-Spotlight {
    $basePath = "HKCU:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"

    if (-not (Test-Path $basePath)) {
        New-Item -Path $basePath -Force | Out-Null
    }

    # Set GPO-equivalent values to disable Spotlight
    $settings = @{
        "DisableWindowsSpotlightFeatures"          = 1  # Master switch
        "DisableWindowsSpotlightOnActionCenter"    = 1
        "DisableWindowsSpotlightOnLockScreen"      = 1
        "DisableWindowsSpotlightSuggestions"       = 1
    }

    foreach ($name in $settings.Keys) {
        New-ItemProperty -Path $basePath -Name $name -Value $settings[$name] -PropertyType DWord -Force | Out-Null
    }

    Write-Output "Spotlight - done"
}

function Remove-BloatwareForCurrentUser {
    $appsToRemove = @(
        "Microsoft.WindowsFeedbackHub",
        "Microsoft.ZuneMusic",              # Media Player
        "Microsoft.Clipchamp",
        "MicrosofWt.BingNews",
        "Microsoft.QuickAssist",
        "Microsoft.MicrosoftSolitaireCollection",
        "Microsoft.MicrosoftStickyNotes",
        "Microsoft.BingWeather"
    )

    foreach ($app in $appsToRemove) {
        $installed = Get-AppxPackage -Name $app
        if ($installed) {
            Remove-AppxPackage -Package $installed.PackageFullName -ErrorAction SilentlyContinue
        }
    }

    Write-Host "Bloatware - done"
}

function Set-SolidBlackWallpaper {
    # Set registry values
    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "Wallpaper" -Value ""
    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "WallpaperStyle" -Value "0"
    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "TileWallpaper" -Value "0"
    Set-ItemProperty -Path "HKCU:\Control Panel\Colors"   -Name "Background" -Value "0 0 0"

    # Apply changes immediately using user32.dll
    Add-Type @"
using System.Runtime.InteropServices;
public class NativeMethods {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
"@

    # 20 = SPI_SETDESKWALLPAPER, 3 = SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE
    [NativeMethods]::SystemParametersInfo(20, 0, "", 3) | Out-Null

    Write-Output "Wallpaper - done"
}


function Remove-StartMenuRecommendations {
    # Disable Recommendations in Start Menu for current user
    $regPath = "HKCU:\Software\Policies\Microsoft\Windows\Explorer"
    
    # Create key if missing
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
        # Set policy: HideRecommendedSection = 1
        New-ItemProperty -Path $regPath -Name "HideRecommendedSection" -Value 1 -PropertyType DWord -Force | Out-Null
        
        Write-Output "Start recommendations - done"
    
    }
}


function Add-StartMenuFolders {
    $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Start"
    $valueName = "VisiblePlaces"

    # Hex value for enabling Settings and Downloads:
    # 86,08,73,52,aa,51,43,42,9f,7b,27,76,58,46,59,d4
    $hex = "86,08,73,52,AA,51,43,42,9F,7B,27,76,58,46,59,D4"
    $bytes = $hex -split ',' | ForEach-Object { [byte]"0x$_" }

    # Create key if missing
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }

    # Set the binary value
    Set-ItemProperty -Path $regPath -Name $valueName -Value ([byte[]]$bytes) -Force

    Write-Output "Start folders - done"
}




#endregion

#region Execution


Write-Host "Start script execution"

Set-KeyboardLanguages # set keyboard to En-US & Ro-RO
Write-Host "Keyboard languages - done"

Remove-TaskbarCrap # gets rid of spotlight & task view
Set-TaskbarCombineAlways
Write-Host "Taskbar - done"

Set-DarkModeOn # sets dark mode
Write-Host "Dark mode - done"

Set-WindowsAccentColorDefaultBlue # sets blue as the accent color
Write-Host "Accent color - done"

Center-StartButton # centers start button
Enable-ShowExtensionsAndHiddenFiles # changes file explorer behavior to show extensions & hidden stuff
Set-ExplorerToOpenThisPC # opens file explorer to This PC instead of quick access
Write-Host "File explorer - done"

Disable-NotepadPlusPlusRememberLastSession # disables npp persistent sessions

Disable-KeyboardCrap # disables sticky keys & toggle keys
Write-Host "Sticky keys - done"

Enable-ClipboardHistory # enables clipboard history
Write-Host "Clipboard history - done"

Disable-WidgetsButton # gets rid of widgets

Set-SolidBlackWallpaper # simple black wallpaper

# OS dependent tasks
if ((Get-WindowsMajorVersion) -lt 11) {
    Remove-MusicAndVideosFromQuickAccess # gets rid of Music & Videos from quick access (win10 only)
    Write-Host "Quick access - done"
    Write-Host "Win11 Start menu - skipped"
    Write-Host "Start folders - skipped"
}
else {
    Write-Host "Quick access - skipped"
    Set-Win11StartMenuPreferences # gets rid of start menu bloat
    Add-StartMenuFolders # adds settings & downloads to start
}

# domain state dependent tasks
if (Is-DomainJoined) {
    Write-Host "Spotlight - skipped"
    Write-Host "Bloatware - skipped"
    if ((Get-WindowsMajorversion) -ge 11) {
        Remove-StartMenuRecommendations # removes recommendations completely
    }
}
else {
    Disable-Spotlight # gets rid of spotlight
    Remove-BloatwareForCurrentUser # removes bloat
}

Stop-Process -Name explorer -Force # to apply changes

#endregion
