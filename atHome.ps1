#-----------------------------------------------------------------------------HEADER-------------------------------------------------------------------------------
# File : atHome.ps1
# Author : zorrouraganu	
# Version : 2.0.1
# Description : Feel @home script that turns any Windows 11 machine into my own 
#
# Version History:
#
# Version	Date				By								Comments
# -----------------------------------------------------------------------------------------------------------------------------------------------------------------
# 1.0		June 2025			zorrouraganu        		Creation of script
# 2.0		April 2026			zorrouraganu        		Complete rework
# 2.0.1     April 2026          zorrouraganu                Dealt with new Category view mode in Start Menu in 25H2
# 
#-----------------------------------------------------------------------------HEADER-------------------------------------------------------------------------------

#requires -Version 5.1

#region Configuration

$Config = [ordered]@{
    WorkingFolder              = Join-Path $env:LOCALAPPDATA 'Temp\FeelAtHome'
    LogNamePrefix              = 'FeelAtHome-Win11'
    RestartExplorerAtEnd       = $true

    DisplayLanguage            = 'en-US'
    SecondaryLanguage          = 'ro-RO'
    PrimaryInputTip            = '0409:00000409'   # English (United States) - US
    SecondaryInputTip          = '0418:00010418'   # Romanian (Standard)

    AccentColorHex             = 'D77800'          # RGB
    RemoveBloatOnStandalone    = $true
    ConfigureTaskbarPins       = $true
}

#endregion

#region Bootstrap

$null = New-Item -Path $Config.WorkingFolder -ItemType Directory -Force -ErrorAction SilentlyContinue

$script:LogPath          = Join-Path $Config.WorkingFolder ('{0}-{1}.log' -f $Config.LogNamePrefix, (Get-Date -Format 'yyyyMMdd-HHmmss'))
$script:WarningCount     = 0
$script:ErrorCount       = 0
$script:RequiresSignOut  = $false
$script:RequiresRestart  = $false

try {
    Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class NativeMethods {
    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern bool SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
"@ -ErrorAction SilentlyContinue | Out-Null
}
catch {}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO'
    )

    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    $line | Out-File -FilePath $script:LogPath -Append -Encoding utf8

    switch ($Level) {
        'WARN'  { $script:WarningCount++ }
        'ERROR' { $script:ErrorCount++ }
    }

    Write-Host $line
}

function Test-IsAdmin {
    $current = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($current)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-IsDomainJoined {
    try {
        return [bool](Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop).PartOfDomain
    }
    catch {
        Write-Log "Could not determine domain join state: $($_.Exception.Message)" 'WARN'
        return $false
    }
}

function Ensure-Key {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        $null = New-Item -Path $Path -Force
    }
}

function Set-RegistryValue {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][AllowEmptyString()][AllowNull()][object]$Value,
        [ValidateSet('String','ExpandString','DWord','QWord','Binary')][string]$Type = 'DWord'
    )

    try {
        Ensure-Key -Path $Path

        $exists = $false
        try {
            $null = Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
            $exists = $true
        }
        catch {}

        if ($exists) {
            Set-ItemProperty -Path $Path -Name $Name -Value $Value -ErrorAction Stop
        }
        else {
            $null = New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $Type -Force -ErrorAction Stop
        }

        Write-Log "Set $Path\$Name = $Value"
    }
    catch {
        Write-Log "Failed to set $Path\$Name : $($_.Exception.Message)" 'ERROR'
    }
}

function Remove-RegistryValue {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Name
    )

    try {
        if ((Test-Path -LiteralPath $Path) -and ($null -ne (Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue))) {
            Remove-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
            Write-Log "Removed $Path\$Name"
        }
    }
    catch {
        Write-Log "Failed to remove $Path\$Name : $($_.Exception.Message)" 'WARN'
    }
}

function Get-ColorRefFromRgbHex {
    param([Parameter(Mandatory = $true)][string]$Hex)

    $hex = $Hex.Trim().TrimStart('#')
    if ($hex.StartsWith('0x')) { $hex = $hex.Substring(2) }
    if ($hex.Length -ne 6) { throw "AccentColorHex must be 6 hex digits." }

    $r = [Convert]::ToUInt32($hex.Substring(0,2), 16)
    $g = [Convert]::ToUInt32($hex.Substring(2,2), 16)
    $b = [Convert]::ToUInt32($hex.Substring(4,2), 16)

    return [uint32](($b * 65536) + ($g * 256) + $r)
}

function Get-DwmColorFromRgbHex {
    param([Parameter(Mandatory = $true)][string]$Hex)

    $colorRef = Get-ColorRefFromRgbHex -Hex $Hex
    return [uint32]::Parse(('FF{0:X6}' -f $colorRef), [System.Globalization.NumberStyles]::HexNumber)
}

function Restart-Explorer {
    try {
        Write-Log 'Restarting Explorer'
        Get-Process explorer -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        #Start-Process explorer.exe
    }
    catch {
        Write-Log "Explorer restart failed: $($_.Exception.Message)" 'WARN'
    }
}

#endregion

#region Taskbar

function Set-TaskbarPreferences {
    Write-Log 'Configuring taskbar preferences'

    $advanced = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
    $search   = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Search'

    Disable-WidgetsPolicy   # Widgets off
    Set-RegistryValue -Path $advanced -Name 'ShowTaskViewButton'  -Value 0 -Type DWord   # Task View off
    Set-RegistryValue -Path $advanced -Name 'TaskbarAl'           -Value 1 -Type DWord   # Center
    Set-RegistryValue -Path $advanced -Name 'TaskbarGlomLevel'    -Value 0 -Type DWord   # Always combine
    Set-RegistryValue -Path $advanced -Name 'ShowCopilotButton'   -Value 0 -Type DWord   # Best effort

    Set-RegistryValue -Path $search   -Name 'SearchBoxTaskbarMode'      -Value 0 -Type DWord
    Set-RegistryValue -Path $search   -Name 'SearchboxTaskbarMode'      -Value 0 -Type DWord
    Set-RegistryValue -Path $search   -Name 'SearchBoxTaskbarModeCache' -Value 0 -Type DWord
}

function Disable-WidgetsPolicy {
    Write-Log 'Disabling Widgets via policy (AllowNewsAndInterests=0)'

    if (-not (Test-IsAdmin)) {
        Write-Log 'Skipping Widgets policy because the session is not elevated' 'WARN'
        return
    }

    Set-RegistryValue -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Dsh' -Name 'AllowNewsAndInterests' -Value 0 -Type DWord
}

function Set-TaskbarPins {
    if (-not $Config.ConfigureTaskbarPins) {
        Write-Log 'Taskbar pin configuration disabled in config'
        return
    }

    Write-Log 'Staging taskbar layout XML for Explorer + Edge only'

    $xmlPath = Join-Path $Config.WorkingFolder 'TaskbarLayout.xml'
    $xml = @'
<?xml version="1.0" encoding="utf-8"?>
<LayoutModificationTemplate
    xmlns="http://schemas.microsoft.com/Start/2014/LayoutModification"
    xmlns:defaultlayout="http://schemas.microsoft.com/Start/2014/FullDefaultLayout"
    xmlns:start="http://schemas.microsoft.com/Start/2014/StartLayout"
    xmlns:taskbar="http://schemas.microsoft.com/Start/2014/TaskbarLayout"
    Version="1">
  <CustomTaskbarLayoutCollection PinListPlacement="Replace">
    <defaultlayout:TaskbarLayout>
      <taskbar:TaskbarPinList>
        <taskbar:DesktopApp DesktopApplicationID="Microsoft.Windows.Explorer" />
        <taskbar:DesktopApp DesktopApplicationID="MSEdge" />
      </taskbar:TaskbarPinList>
    </defaultlayout:TaskbarLayout>
  </CustomTaskbarLayoutCollection>
</LayoutModificationTemplate>
'@

    try {
        $xml | Out-File -FilePath $xmlPath -Encoding utf8 -Force
        Write-Log "Wrote taskbar layout XML to $xmlPath"
    }
    catch {
        Write-Log "Failed to write taskbar layout XML: $($_.Exception.Message)" 'ERROR'
        return
    }

    Set-RegistryValue -Path 'HKCU:\Software\Policies\Microsoft\Windows\Explorer' -Name 'StartLayoutFile'   -Value $xmlPath -Type String
    Set-RegistryValue -Path 'HKCU:\Software\Policies\Microsoft\Windows\Explorer' -Name 'LockedStartLayout' -Value 1        -Type DWord

    if (Test-IsAdmin) {
        Set-RegistryValue -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer' -Name 'LayoutXMLPath' -Value $xmlPath -Type String
    }

    $script:RequiresSignOut = $true
}

#endregion

#region Appearance

function Set-SolidBlackWallpaper {
    Write-Log 'Setting solid black wallpaper'

    Set-RegistryValue -Path 'HKCU:\Control Panel\Colors'  -Name 'Background'     -Value '0 0 0' -Type String
    Set-RegistryValue -Path 'HKCU:\Control Panel\Desktop' -Name 'Wallpaper'      -Value ''      -Type String
    Set-RegistryValue -Path 'HKCU:\Control Panel\Desktop' -Name 'WallpaperStyle' -Value '0'     -Type String
    Set-RegistryValue -Path 'HKCU:\Control Panel\Desktop' -Name 'TileWallpaper'  -Value '0'     -Type String

    try {
        [void][NativeMethods]::SystemParametersInfo(20, 0, '', 3)
        Write-Log 'Wallpaper refresh requested'
    }
    catch {
        Write-Log "Wallpaper refresh call failed: $($_.Exception.Message)" 'WARN'
    }
}

function Set-DarkModeAndAccent {
    Write-Log 'Configuring dark mode and accent color'

    $colorRef = Get-ColorRefFromRgbHex -Hex $Config.AccentColorHex
    $dwmColor = Get-DwmColorFromRgbHex -Hex $Config.AccentColorHex

    Set-RegistryValue -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name 'AppsUseLightTheme'   -Value 0 -Type DWord
    Set-RegistryValue -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name 'SystemUsesLightTheme' -Value 0 -Type DWord

    # Keep accent color defined, but do NOT show it on Start/taskbar
    Set-RegistryValue -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name 'ColorPrevalence'      -Value 0 -Type DWord
    Set-RegistryValue -Path 'HKCU:\Control Panel\Desktop'                                           -Name 'AutoColorization'    -Value 0 -Type DWord

    # Accent color itself
    Set-RegistryValue -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Accent'       -Name 'AccentColorMenu'    -Value $colorRef -Type DWord
    Set-RegistryValue -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Accent'       -Name 'StartColorMenu'     -Value $colorRef -Type DWord
    Set-RegistryValue -Path 'HKCU:\Software\Microsoft\Windows\DWM'                                  -Name 'AccentColor'        -Value $dwmColor -Type DWord
    Set-RegistryValue -Path 'HKCU:\Software\Microsoft\Windows\DWM'                                  -Name 'ColorizationColor'  -Value $dwmColor -Type DWord

    # Do NOT show accent on title bars / window borders
    Set-RegistryValue -Path 'HKCU:\Software\Microsoft\Windows\DWM'                                  -Name 'ColorPrevalence'    -Value 0 -Type DWord
}

#endregion

#region Explorer

function Set-ExplorerPreferences {
    Write-Log 'Configuring File Explorer preferences'

    $advanced = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'

    Set-RegistryValue -Path $advanced -Name 'HideFileExt' -Value 0 -Type DWord
    Set-RegistryValue -Path $advanced -Name 'Hidden'      -Value 1 -Type DWord
    Set-RegistryValue -Path $advanced -Name 'LaunchTo'    -Value 1 -Type DWord
}

#endregion

#region Language and input

function Set-LanguageAndInput {
    Write-Log 'Configuring language, display language and keyboard layouts'

    Import-Module International -ErrorAction SilentlyContinue | Out-Null
    Import-Module LanguagePackManagement -ErrorAction SilentlyContinue | Out-Null

    $languageAlreadyInstalled = $false

    if (Get-Command -Name Get-InstalledLanguage -ErrorAction SilentlyContinue) {
        try {
            $installedLang = Get-InstalledLanguage -Language $Config.DisplayLanguage -ErrorAction SilentlyContinue
            if ($installedLang) {
                $languageAlreadyInstalled = $true
                Write-Log "Language already installed: $($Config.DisplayLanguage)"
            }
        }
        catch {
            Write-Log "Get-InstalledLanguage check failed: $($_.Exception.Message)" 'WARN'
        }
    }

    if ((-not $languageAlreadyInstalled) -and (Get-Command -Name Install-Language -ErrorAction SilentlyContinue)) {
        if (Test-IsAdmin) {
            try {
                Install-Language -Language $Config.DisplayLanguage -CopyToSettings -ErrorAction Stop | Out-Null
                Write-Log "Installed language pack: $($Config.DisplayLanguage)"
            }
            catch {
                Write-Log "Install-Language for $($Config.DisplayLanguage) failed: $($_.Exception.Message)" 'WARN'
            }
        }
        else {
            Write-Log 'Install-Language skipped because the session is not elevated' 'WARN'
        }
    }
    elseif ($languageAlreadyInstalled) {
        Write-Log "Skipped Install-Language because $($Config.DisplayLanguage) is already installed"
    }

    try {
        $currentUiOverride = Get-WinUILanguageOverride -ErrorAction SilentlyContinue
        if ($currentUiOverride -ne $Config.DisplayLanguage) {
            Set-WinUILanguageOverride -Language $Config.DisplayLanguage
            Write-Log "UI language override set to $($Config.DisplayLanguage)"
            $script:RequiresSignOut = $true
        }
        else {
            Write-Log "UI language override already set to $($Config.DisplayLanguage)"
        }
    }
    catch {
        Write-Log "Failed to set UI language override: $($_.Exception.Message)" 'ERROR'
    }

    try {
        $langList = New-WinUserLanguageList $Config.DisplayLanguage
        $langList[0].InputMethodTips.Clear()
        [void]$langList[0].InputMethodTips.Add($Config.PrimaryInputTip)

        $roList = New-WinUserLanguageList $Config.SecondaryLanguage
        $roList[0].InputMethodTips.Clear()
        [void]$roList[0].InputMethodTips.Add($Config.SecondaryInputTip)

        [void]$langList.Add($roList[0])

        Set-WinUserLanguageList -LanguageList $langList -Force
        Write-Log 'User language list updated'
    }
    catch {
        Write-Log "Failed to set WinUserLanguageList: $($_.Exception.Message)" 'ERROR'
    }

    try {
        Set-Culture -CultureInfo $Config.DisplayLanguage
        Write-Log "Culture set to $($Config.DisplayLanguage)"
    }
    catch {
        Write-Log "Failed to set culture: $($_.Exception.Message)" 'WARN'
    }

    if (Get-Command -Name Set-WinDefaultInputMethodOverride -ErrorAction SilentlyContinue) {
        try {
            Set-WinDefaultInputMethodOverride -InputTip $Config.PrimaryInputTip
            Write-Log "Default input method set to $($Config.PrimaryInputTip)"
        }
        catch {
            Write-Log "Failed to set default input method override: $($_.Exception.Message)" 'WARN'
        }
    }

    if (Test-IsAdmin) {
        try {
            Set-WinSystemLocale -SystemLocale $Config.DisplayLanguage
            Write-Log "System locale set to $($Config.DisplayLanguage)"
            $script:RequiresRestart = $true
        }
        catch {
            Write-Log "Failed to set system locale: $($_.Exception.Message)" 'WARN'
        }

        if (Get-Command -Name Copy-UserInternationalSettingsToSystem -ErrorAction SilentlyContinue) {
            try {
                Copy-UserInternationalSettingsToSystem -WelcomeScreen $true -NewUser $true
                Write-Log 'Copied user international settings to Welcome screen and new users'
                $script:RequiresRestart = $true
            }
            catch {
                Write-Log "Copy-UserInternationalSettingsToSystem failed: $($_.Exception.Message)" 'WARN'
            }
        }
    }
    else {
        Write-Log 'System locale copy skipped because the session is not elevated' 'WARN'
    }
}

#endregion

#region Accessibility and clipboard

function Set-AccessibilityPreferences {
    Write-Log 'Disabling Sticky Keys and Toggle Keys'

    Set-RegistryValue -Path 'HKCU:\Control Panel\Accessibility\StickyKeys' -Name 'Flags' -Value '506' -Type String
    Set-RegistryValue -Path 'HKCU:\Control Panel\Accessibility\ToggleKeys' -Name 'Flags' -Value '58'  -Type String
}

function Set-ClipboardHistory {
    Write-Log 'Enabling clipboard history'

    Set-RegistryValue -Path 'HKCU:\Software\Microsoft\Clipboard' -Name 'EnableClipboardHistory' -Value 1 -Type DWord

    if (Test-IsAdmin) {
        Set-RegistryValue -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System' -Name 'AllowClipboardHistory' -Value 1 -Type DWord
    }
    else {
        Write-Log 'Machine policy for clipboard history skipped because the session is not elevated' 'WARN'
    }
}

#endregion

#region Start menu

function Set-StartVisiblePlaces {
    Write-Log 'Configuring Start folders via HKCU VisiblePlaces'

    $startKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Start'

    $foldersToShow = @(
        [Guid]'E367B32F-89DE-4355-BFCE-61F37B18A937'  # Downloads
        [Guid]'52730886-51AA-4243-9F7B-2776584659D4'  # Settings
    )

    $bytes = New-Object System.Collections.Generic.List[byte]
    foreach ($folder in $foldersToShow) {
        $bytes.AddRange($folder.ToByteArray())
    }

    Set-RegistryValue -Path $startKey -Name 'VisiblePlaces' -Value $bytes.ToArray() -Type Binary
}

function Set-StartMenuPreferences {
    Write-Log 'Configuring Start menu preferences'

    $advanced   = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
    $startKey   = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Start'
    $startPol   = 'HKCU:\Software\Policies\Microsoft\Windows\Explorer'

    Set-RegistryValue -Path $advanced -Name 'Start_Layout'              -Value 1 -Type DWord   # More pins
    Set-RegistryValue -Path $advanced -Name 'Start_IrisRecommendations' -Value 0 -Type DWord   # Tips off
    Set-RegistryValue -Path $advanced -Name 'Start_TrackDocs'           -Value 0 -Type DWord   # Recommended files off

    Set-RegistryValue -Path $startKey -Name 'ShowRecentList' -Value 1 -Type DWord               # Recently added apps on

    Remove-RegistryValue -Path $startPol -Name 'HideRecentlyAddedApps'
    Set-RegistryValue -Path $startPol -Name 'ShowOrHideMostUsedApps'        -Value 2 -Type DWord # Most used off
    Set-RegistryValue -Path $startPol -Name 'HideRecommendedSection'        -Value 1 -Type DWord # Recommended off
    Set-RegistryValue -Path $startPol -Name 'HideRecommendedPersonalizedSites' -Value 1 -Type DWord

    Set-RegistryValue -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Start' -Name 'AllAppsViewMode' -Value 2 -Type DWord # Sets apps view mode to List

    Set-StartVisiblePlaces
}

#endregion

#region Notepad++

function Disable-NotepadPlusPlusSession {
    Write-Log 'Checking for Notepad++'

    $roots = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    $app = $null
    foreach ($root in $roots) {
        $app = Get-ItemProperty -Path $root -ErrorAction SilentlyContinue |
               Where-Object { $_.DisplayName -like 'Notepad++*' } |
               Select-Object -First 1
        if ($app) { break }
    }

    if (-not $app) {
        Write-Log 'Notepad++ not detected'
        return
    }

    Write-Log 'Notepad++ detected'

    $candidates = @(
        (Join-Path $env:APPDATA 'Notepad++\config.xml')
    )

    if ($app.InstallLocation) {
        $candidates += (Join-Path $app.InstallLocation 'config.xml')
    }

    $candidates = $candidates | Select-Object -Unique

    foreach ($path in $candidates) {
        if (-not (Test-Path -LiteralPath $path)) { continue }

        try {
            $raw = Get-Content -LiteralPath $path -Raw -Encoding UTF8

            $updated = $raw
            $updated = [regex]::Replace($updated, 'rememberLastSession="yes"', 'rememberLastSession="no"', 'IgnoreCase')
            $updated = [regex]::Replace($updated, 'rememberLastSession="true"', 'rememberLastSession="false"', 'IgnoreCase')
            $updated = [regex]::Replace($updated, '(<GUIConfig\s+name="RememberCurrentSessionForNextLaunch"\s*>)(yes|true|1)(</GUIConfig>)', '$1no$3', 'IgnoreCase')
            $updated = [regex]::Replace($updated, '(<GUIConfig\s+name="RememberLastSession"\s*>)(yes|true|1)(</GUIConfig>)', '$1no$3', 'IgnoreCase')

            if ($updated -ne $raw) {
                $updated | Out-File -LiteralPath $path -Encoding UTF8 -Force
                Write-Log "Updated Notepad++ config: $path"
            }
            else {
                Write-Log "No matching Notepad++ session setting found in $path" 'WARN'
            }
        }
        catch {
            Write-Log "Failed to update $path : $($_.Exception.Message)" 'WARN'
        }
    }
}

#endregion

#region Spotlight and bloat

function Disable-SpotlightIfStandalone {
    param([bool]$IsDomainJoined)

    if ($IsDomainJoined) {
        Write-Log 'Domain joined machine detected, Spotlight policy left untouched'
        return
    }

    Write-Log 'Standalone machine detected, disabling Spotlight'

    Set-RegistryValue -Path 'HKCU:\Software\Policies\Microsoft\Windows\CloudContent' -Name 'DisableWindowsSpotlightFeatures'   -Value 1 -Type DWord
    Set-RegistryValue -Path 'HKCU:\Software\Policies\Microsoft\Windows\CloudContent' -Name 'DisableSpotlightCollectionOnDesktop' -Value 1 -Type DWord

    if (Test-IsAdmin) {
        Set-RegistryValue -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent' -Name 'DisableWindowsSpotlightFeatures'   -Value 1 -Type DWord
        Set-RegistryValue -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent' -Name 'DisableSpotlightCollectionOnDesktop' -Value 1 -Type DWord
    }
}

function Remove-AppxPattern {
    param([Parameter(Mandatory = $true)][string]$Pattern)

    try {
        $installed = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | Where-Object { $_.Name -like $Pattern }
        foreach ($pkg in $installed) {
            try {
                Remove-AppxPackage -Package $pkg.PackageFullName -AllUsers -ErrorAction Stop
                Write-Log "Removed installed package: $($pkg.Name)"
            }
            catch {
                Write-Log "Failed to remove installed package $($pkg.Name): $($_.Exception.Message)" 'WARN'
            }
        }

        $provisioned = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like $Pattern }
        foreach ($pkg in $provisioned) {
            try {
                $null = Remove-AppxProvisionedPackage -Online -PackageName $pkg.PackageName -ErrorAction Stop
                Write-Log "Removed provisioned package: $($pkg.DisplayName)"
            }
            catch {
                Write-Log "Failed to remove provisioned package $($pkg.DisplayName): $($_.Exception.Message)" 'WARN'
            }
        }
    }
    catch {
        Write-Log "Package removal failed for pattern $Pattern : $($_.Exception.Message)" 'WARN'
    }
}

function Remove-BloatIfStandalone {
    param([bool]$IsDomainJoined)

    if ($IsDomainJoined) {
        Write-Log 'Domain joined machine detected, bloat removal skipped'
        return
    }

    if (-not $Config.RemoveBloatOnStandalone) {
        Write-Log 'Standalone bloat removal disabled in config'
        return
    }

    if (-not (Test-IsAdmin)) {
        Write-Log 'Bloat removal skipped because the session is not elevated' 'WARN'
        return
    }

    Write-Log 'Standalone machine detected, removing requested inbox apps'

    $patterns = @(
        'Microsoft.WindowsFeedbackHub',
        'Microsoft.ZuneMusic',
        'Clipchamp.Clipchamp',
        'Microsoft.BingNews',
        'MicrosoftCorporationII.QuickAssist',
        'Microsoft.MicrosoftSolitaireCollection',
        'Microsoft.MicrosoftStickyNotes',
        'Microsoft.BingWeather'
    )

    foreach ($pattern in $patterns) {
        Remove-AppxPattern -Pattern $pattern
    }
}

#endregion

#region Main

Write-Log '--- Script started ---'
Write-Log "Log file: $script:LogPath"

$IsAdmin       = Test-IsAdmin
$IsDomainJoined = Test-IsDomainJoined

Write-Log ("Elevated session: {0}" -f $IsAdmin)
Write-Log ("Domain joined: {0}" -f $IsDomainJoined)

Set-SolidBlackWallpaper
Set-TaskbarPreferences
Set-TaskbarPins
Set-LanguageAndInput
Set-DarkModeAndAccent
Set-ExplorerPreferences
Disable-NotepadPlusPlusSession
Set-AccessibilityPreferences
Set-StartMenuPreferences
Set-ClipboardHistory
Disable-SpotlightIfStandalone -IsDomainJoined:$IsDomainJoined
Remove-BloatIfStandalone -IsDomainJoined:$IsDomainJoined

if ($Config.RestartExplorerAtEnd) {
    Restart-Explorer
}

Write-Log '--- Script finished ---'

if ($script:RequiresSignOut) {
    Write-Log 'A sign out / sign in is required for all changes to fully apply' 'WARN'
}

if ($script:RequiresRestart) {
    Write-Log 'A restart is recommended for all language settings to fully apply' 'WARN'
}

Write-Log ("Warnings: {0} | Errors: {1}" -f $script:WarningCount, $script:ErrorCount)

Write-Host ''
Write-Host ('Done. Log: {0}' -f $script:LogPath)
if ($script:RequiresSignOut) { Write-Host 'Sign out/in required.' }
if ($script:RequiresRestart) { Write-Host 'Restart recommended.' }

#endregion
