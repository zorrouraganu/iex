#region Functions
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
    $win11 = [System.Environment]::OSVersion.Version.Major -eq 10 -and `
             [System.Environment]::OSVersion.Version.Build -ge 22000

    if (-not $win11) { 
        Write-Host "Win11 Start Menu - skipped"
        return
        }

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

function Remove-PinnedTaskbarIcons {
    $appsToRemove = @(
        "Microsoft.WindowsStore_8wekyb3d8bbwe!App",
        "Microsoft.MicrosoftOfficeHub_8wekyb3d8bbwe!Microsoft.MicrosoftOfficeHub"
    )

    foreach ($appId in $appsToRemove) {
        $command = "explorer.exe shell:Appsfolder\$appId"
        $verb = "unpin from taskbar"

        # Get the shell application object
        $shellApp = New-Object -ComObject Shell.Application
        $folder = $shellApp.Namespace("shell:Appsfolder")

        foreach ($item in $folder.Items()) {
            if ($item.Path -eq $command) {
                foreach ($itemVerb in $item.Verbs()) {
                    if ($itemVerb.Name.Replace('&','').ToLower() -eq $verb) {
                        $itemVerb.DoIt()
                        Write-Output "Unpinned: $appId"
                    }
                }
            }
        }
    }
}


#endregion

#region Execution

Write-Host "Start script execution"
Set-KeyboardLanguages
Write-Host "Keboard languages - done"
Remove-TaskbarCrap
Set-TaskbarCombineAlways
Write-Host "Taskbar - done"
Set-DarkModeOn
Write-Host "Dark mode - done"
Set-WindowsAccentColorDefaultBlue
Write-Host "Accent color - done"
Center-StartButton
Enable-ShowExtensionsAndHiddenFiles
Set-ExplorerToOpenThisPC
Write-Host "File explorer - done"
Remove-MusicAndVideosFromQuickAccess
Write-Host "Quick access - done"
Disable-NotepadPlusPlusRememberLastSession
Disable-KeyboardCrap
Write-Host "Sticky keys - done"
Set-Win11StartMenuPreferences
Enable-ClipboardHistory
Write-Host "Clipboard history - done"
Disable-WidgetsButton
Remove-PinnedTaskbarIcons

Stop-Process -Name explorer -Force # to apply changes

#endregion
