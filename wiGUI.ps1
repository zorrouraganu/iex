#-----------------------------------------------------------------------------HEADER-------------------------------------------------------------------------------
# File : wiGUI.ps1
# Author : zorrouraganu	
# Version : 2.2.0
# Description : WinGet GUI that allows easy install of multiple apps 
#
# Version History:
#
# Version	Date				By								Comments
# -----------------------------------------------------------------------------------------------------------------------------------------------------------------
# 1.0		June 2025			zorrouraganu        		Creation of script
# 2.0		April 2026			zorrouraganu        		Complete rework
# 2.1       April 2026          zorrouraganu                Bug fixes & improvements
# 2.2       April 2026          zorrouraganu                Async busy operations so progress spinner animates
#                                                           UI improvements
# 
#-----------------------------------------------------------------------------HEADER-------------------------------------------------------------------------------

#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Configuration
$AppTitle                 = 'wiGUI'
$WindowWidth              = 1320
$WindowHeight             = 860
$SearchResultLimit        = 100
$DefaultSilentInstall     = $true
$LogFolderName            = 'wiGUI'
$DefaultWingetSource      = 'winget'
$RepositorySearchHint     = 'Search one package by name or ID'
$EnableMicaBackdrop       = $true
$IncludeUpdateAllButton   = $true

$CommonApps = @(
    @{ Name = 'Google Chrome';         Id = 'Google.Chrome';                 Source = 'winget' }
    @{ Name = 'Mozilla Firefox';       Id = 'Mozilla.Firefox';               Source = 'winget' }
    @{ Name = 'Microsoft Edge';        Id = 'Microsoft.Edge';                Source = 'winget' }
    @{ Name = 'Visual Studio Code';    Id = 'Microsoft.VisualStudioCode';    Source = 'winget' }
    @{ Name = '7-Zip';                 Id = '7zip.7zip';                     Source = 'winget' }
    @{ Name = 'Notepad++';             Id = 'Notepad++.Notepad++';           Source = 'winget' }
    @{ Name = 'VLC media player';      Id = 'VideoLAN.VLC';                  Source = 'winget' }
    @{ Name = 'PowerToys';             Id = 'Microsoft.PowerToys';           Source = 'winget' }
    @{ Name = 'Git';                   Id = 'Git.Git';                       Source = 'winget' }
    @{ Name = 'Python 3';              Id = 'Python.Python.3.12';            Source = 'winget' }
    @{ Name = 'Docker Desktop';        Id = 'Docker.DockerDesktop';          Source = 'winget' }
    @{ Name = 'OBS Studio';            Id = 'OBSProject.OBSStudio';          Source = 'winget' }
    @{ Name = 'Spotify';               Id = 'Spotify.Spotify';               Source = 'winget' }
    @{ Name = 'Discord';               Id = 'Discord.Discord';               Source = 'winget' }
)

#endregion

#region Bootstrap
Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase
Add-Type -AssemblyName System.Xaml

$script:LogRoot = Join-Path -Path $env:LOCALAPPDATA -ChildPath ("Temp\{0}" -f $LogFolderName)
$null = New-Item -Path $script:LogRoot -ItemType Directory -Force
$script:SessionStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$script:LogPath = Join-Path -Path $script:LogRoot -ChildPath ("wiGUI_{0}.log" -f $script:SessionStamp)
$script:IsBusy = $false
$script:ActiveWingetOperations = New-Object System.Collections.ArrayList

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS', 'DEBUG')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $line = '[{0}] [{1}] {2}' -f $timestamp, $Level, $Message

    Add-Content -Path $script:LogPath -Value $line -Encoding UTF8

    switch ($Level) {
        'WARN'    { Write-Host $line -ForegroundColor Yellow }
        'ERROR'   { Write-Host $line -ForegroundColor Red }
        'SUCCESS' { Write-Host $line -ForegroundColor Green }
        'DEBUG'   { Write-Host $line -ForegroundColor DarkGray }
        default   { Write-Host $line -ForegroundColor Gray }
    }
}

function Set-UiBusy {
    param(
        [bool]$Busy,
        [string]$StatusText
    )

    $script:IsBusy = $Busy

    if ($StatusText) {
        $script:LblStatus.Text = $StatusText
        if ($script:BusyOverlayMessage) {
            $script:BusyOverlayMessage.Text = $StatusText
        }
    }
    elseif ($script:BusyOverlayMessage) {
        $script:BusyOverlayMessage.Text = 'Working...'
    }

    $script:Window.Cursor = if ($Busy) {
        [System.Windows.Input.Cursors]::Wait
    }
    else {
        [System.Windows.Input.Cursors]::Arrow
    }

    $script:ProgressBar.IsIndeterminate = $Busy

    if ($script:BusyOverlay) {
        $script:BusyOverlay.Visibility = if ($Busy) { 'Visible' } else { 'Collapsed' }
        $script:BusyOverlay.IsHitTestVisible = $Busy
    }

    $script:Window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Render)
}

function Set-BusyMessage {
    param([string]$Message)

    if ([string]::IsNullOrWhiteSpace($Message)) {
        return
    }

    $script:LblStatus.Text = $Message

    if ($script:BusyOverlayMessage) {
        $script:BusyOverlayMessage.Text = $Message
    }

    $script:Window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Render)
}

function Invoke-LogRetentionCleanup {
    param([int]$RetentionDays = 7)

    if (-not (Test-Path -LiteralPath $script:LogRoot)) {
        return
    }

    $cutoff = (Get-Date).AddDays(-1 * [Math]::Abs($RetentionDays))
    $removedCount = 0

    try {
        $staleLogFiles = Get-ChildItem -LiteralPath $script:LogRoot -File -Filter '*.log' -ErrorAction Stop |
            Where-Object { $_.LastWriteTime -lt $cutoff -and $_.FullName -ne $script:LogPath }
    }
    catch {
        Write-Log ("Failed to enumerate old log files in '{0}': {1}" -f $script:LogRoot, $_.Exception.Message) 'WARN'
        return
    }

    foreach ($staleLogFile in $staleLogFiles) {
        try {
            Remove-Item -LiteralPath $staleLogFile.FullName -Force -ErrorAction Stop
            $removedCount++
        }
        catch {
            Write-Log ("Failed to remove old log file '{0}': {1}" -f $staleLogFile.FullName, $_.Exception.Message) 'WARN'
        }
    }

    if ($removedCount -gt 0) {
        Write-Log ("Removed {0} old log file(s) older than {1} day(s) from {2}" -f $removedCount, [Math]::Abs($RetentionDays), $script:LogRoot) 'INFO'
    }
}

Write-Log "Session started. Log file: $script:LogPath"
Invoke-LogRetentionCleanup -RetentionDays 7
#endregion

#region Native styling
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class NativeDwm
{
    [DllImport("dwmapi.dll")]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int pvAttribute, int cbAttribute);
}
"@

function Get-WindowsBuild {
    try {
        return [int](Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name CurrentBuildNumber).CurrentBuildNumber
    }
    catch {
        return 0
    }
}

function Set-ModernWindowChrome {
    param([Parameter(Mandatory = $true)][System.Windows.Window]$Window)

    try {
        $interop = New-Object System.Windows.Interop.WindowInteropHelper($Window)
        $handle = $interop.EnsureHandle()

        $darkMode = 1
        $roundPreference = 2
        [void][NativeDwm]::DwmSetWindowAttribute($handle, 20, [ref]$darkMode, 4)
        [void][NativeDwm]::DwmSetWindowAttribute($handle, 33, [ref]$roundPreference, 4)

        if ($EnableMicaBackdrop -and (Get-WindowsBuild) -ge 22621) {
            $mica = 2
            [void][NativeDwm]::DwmSetWindowAttribute($handle, 38, [ref]$mica, 4)
        }
    }
    catch {
        Write-Log "Failed to apply Windows 11 chrome styling: $($_.Exception.Message)" 'WARN'
    }
}
#endregion

#region Data model
if (-not ('WingetPackageItem' -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.ComponentModel;

public class WingetPackageItem : INotifyPropertyChanged
{
    private bool _selected;

    public bool Selected
    {
        get { return _selected; }
        set
        {
            if (_selected != value)
            {
                _selected = value;
                OnPropertyChanged("Selected");
            }
        }
    }

    public string Name { get; set; }
    public string Id { get; set; }
    public string Version { get; set; }
    public string Source { get; set; }
    public string Origin { get; set; }

    public event PropertyChangedEventHandler PropertyChanged;

    protected void OnPropertyChanged(string propertyName)
    {
        var handler = PropertyChanged;
        if (handler != null)
        {
            handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
"@
}

function New-PackageItem {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [string]$Id,
        [string]$Version,
        [string]$Source = $DefaultWingetSource,
        [ValidateSet('Startup', 'Search', 'Selected')][string]$Origin = 'Search'
    )

    $item = New-Object WingetPackageItem
    $item.Selected = $false
    $item.Name = $Name
    $item.Id = $Id
    $item.Version = $Version
    $item.Source = $Source
    $item.Origin = $Origin
    return $item
}

$script:ResultPackages = New-Object 'System.Collections.ObjectModel.ObservableCollection[WingetPackageItem]'
$script:SelectedPackages = New-Object 'System.Collections.ObjectModel.ObservableCollection[WingetPackageItem]'

function Get-PackageKey {
    param([Parameter(Mandatory = $true)][WingetPackageItem]$Package)
    return ('{0}|{1}|{2}' -f $Package.Source, $Package.Id, $Package.Name)
}

function Get-StartupCatalogPackages {
    $items = New-Object System.Collections.Generic.List[WingetPackageItem]
    foreach ($app in $CommonApps) {
        $source = if ($app.ContainsKey('Source') -and -not [string]::IsNullOrWhiteSpace([string]$app.Source)) { [string]$app.Source } else { $DefaultWingetSource }
        [void]$items.Add((New-PackageItem -Name $app.Name -Id $app.Id -Source $source -Origin 'Startup'))
    }
    return @($items.ToArray())
}

function Find-SelectedPackageByKey {
    param([Parameter(Mandatory = $true)][string]$Key)

    foreach ($item in $script:SelectedPackages) {
        if ((Get-PackageKey -Package $item) -eq $Key) {
            return $item
        }
    }

    return $null
}

function Set-ResultPackages {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$Items
    )

    $script:ResultPackages.Clear()

    foreach ($incoming in @($Items)) {
        $key = Get-PackageKey -Package $incoming
        $selectedItem = Find-SelectedPackageByKey -Key $key

        if ($selectedItem) {
            $selectedItem.Name = $incoming.Name
            $selectedItem.Id = $incoming.Id
            $selectedItem.Version = $incoming.Version
            $selectedItem.Source = $incoming.Source
            $selectedItem.Origin = $incoming.Origin
            $script:ResultPackages.Add($selectedItem)
        }
        else {
            $script:ResultPackages.Add($incoming)
        }
    }
}

function Sync-SelectedPackages {
    $selectedLookup = @{}
    foreach ($item in @($script:SelectedPackages)) {
        $selectedLookup[(Get-PackageKey -Package $item)] = $item
    }

    foreach ($item in @($script:ResultPackages)) {
        $key = Get-PackageKey -Package $item
        if ($item.Selected -and -not $selectedLookup.ContainsKey($key)) {
            $script:SelectedPackages.Add($item)
            $selectedLookup[$key] = $item
        }
        elseif ((-not $item.Selected) -and $selectedLookup.ContainsKey($key)) {
            [void]$script:SelectedPackages.Remove($selectedLookup[$key])
            $selectedLookup.Remove($key)
        }
    }

    foreach ($item in @($script:SelectedPackages)) {
        if (-not $item.Selected) {
            [void]$script:SelectedPackages.Remove($item)
        }
    }
}

function Get-SelectedPackages {
    Sync-SelectedPackages
    return @($script:SelectedPackages | Where-Object { $_.Selected })
}
#endregion

#region WinGet helpers
function Test-WingetAvailable {
    return [bool](Get-Command winget -ErrorAction SilentlyContinue)
}

function Ensure-Winget {
    if (Test-WingetAvailable) {
        Write-Log 'winget is available.' 'DEBUG'
        return $true
    }

    Write-Log 'winget was not found on PATH.' 'WARN'

    $message = @(
        'winget (App Installer) was not found.',
        '',
        'The script can try Microsoft''s supported repair/bootstrap path first.',
        'If that fails, you can install or repair App Installer and then relaunch the script.',
        '',
        'Proceed?'
    ) -join [Environment]::NewLine

    $choice = [System.Windows.MessageBox]::Show(
        $message,
        $AppTitle,
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )

    if ($choice -ne [System.Windows.MessageBoxResult]::Yes) {
        Write-Log 'User declined winget bootstrap/repair attempt.' 'WARN'
        return $false
    }

    try {
        Write-Log 'Attempting App Installer registration...' 'INFO'
        Add-AppxPackage -RegisterByFamilyName -MainPackage 'Microsoft.DesktopAppInstaller_8wekyb3d8bbwe'
        Start-Sleep -Seconds 2
    }
    catch {
        Write-Log "App Installer registration attempt failed: $($_.Exception.Message)" 'WARN'
    }

    if (Test-WingetAvailable) {
        Write-Log 'winget became available after App Installer registration.' 'SUCCESS'
        return $true
    }

    try {
        Write-Log 'Attempting WinGet repair bootstrap via Microsoft.WinGet.Client...' 'INFO'
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        Install-PackageProvider -Name NuGet -Force | Out-Null
        Install-Module -Name Microsoft.WinGet.Client -Force -Repository PSGallery -Scope CurrentUser | Out-Null
        Import-Module Microsoft.WinGet.Client -Force
        Repair-WinGetPackageManager | Out-Null
        Start-Sleep -Seconds 2
    }
    catch {
        Write-Log "Repair-WinGetPackageManager bootstrap failed: $($_.Exception.Message)" 'ERROR'
    }

    if (Test-WingetAvailable) {
        Write-Log 'winget became available after repair bootstrap.' 'SUCCESS'
        return $true
    }

    [System.Windows.MessageBox]::Show(
        "winget is still unavailable. Install or repair 'App Installer', then relaunch the script.`n`nLog: $script:LogPath",
        $AppTitle,
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Error
    ) | Out-Null

    return $false
}

function New-WingetLogPath {
    param([Parameter(Mandatory = $true)][string]$Prefix)

    $safePrefix = ($Prefix -replace '[^a-zA-Z0-9_-]', '_')
    return Join-Path -Path $script:LogRoot -ChildPath ('{0}_{1}.log' -f $safePrefix, (Get-Date -Format 'yyyyMMdd_HHmmss_fff'))
}

function ConvertTo-CommandLineArgumentString {
    param([Parameter(Mandatory = $true)][string[]]$Arguments)

    $quotedArguments = foreach ($argument in $Arguments) {
        if ($null -eq $argument) { continue }

        if ($argument -eq '') {
            '""'
        }
        elseif ($argument -match '[\s"]') {
            '"{0}"' -f ($argument -replace '"', '\"')
        }
        else {
            $argument
        }
    }

    return [string]::Join(' ', $quotedArguments)
}

function Get-WingetProcessResult {
    param(
        [Parameter(Mandatory = $true)][System.Diagnostics.Process]$Process,
        [Parameter(Mandatory = $true)][string]$StdOutPath,
        [Parameter(Mandatory = $true)][string]$StdErrPath,
        [switch]$StreamToLog
    )

    try {
        $null = $Process.WaitForExit()
    }
    catch {}

    try {
        $Process.Refresh()
    }
    catch {}

    $stdOut = if (Test-Path -LiteralPath $StdOutPath) { [System.IO.File]::ReadAllText($StdOutPath) } else { '' }
    $stdErr = if (Test-Path -LiteralPath $StdErrPath) { [System.IO.File]::ReadAllText($StdErrPath) } else { '' }

    if ($StreamToLog) {
        foreach ($line in ($stdOut -split "`r?`n")) {
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                Write-Log $line.Trim() 'INFO'
            }
        }

        foreach ($line in ($stdErr -split "`r?`n")) {
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                Write-Log $line.Trim() 'WARN'
            }
        }
    }
    elseif ($stdErr) {
        foreach ($line in ($stdErr -split "`r?`n")) {
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                Write-Log $line.Trim() 'DEBUG'
            }
        }
    }

    $exitCode = $null

    try {
        $exitCode = [int]$Process.ExitCode
    }
    catch {
        try {
            $Process.Refresh()
            $exitCode = [int]$Process.ExitCode
        }
        catch {}
    }

    try { Remove-Item -LiteralPath $StdOutPath -Force -ErrorAction SilentlyContinue } catch {}
    try { Remove-Item -LiteralPath $StdErrPath -Force -ErrorAction SilentlyContinue } catch {}

    return [pscustomobject]@{
        ExitCode = $exitCode
        StdOut   = $stdOut
        StdErr   = $stdErr
    }
}

function Start-WingetProcess {
    param(
        [Parameter(Mandatory = $true)][string[]]$Arguments,
        [switch]$CaptureOutput,
        [switch]$StreamToLog,
        [string]$FriendlyName = 'winget'
    )

    $argumentString = ConvertTo-CommandLineArgumentString -Arguments $Arguments
    $stdOutPath = New-WingetLogPath -Prefix 'stdout'
    $stdErrPath = New-WingetLogPath -Prefix 'stderr'

    Write-Log ("Starting {0}: winget {1}" -f $FriendlyName, $argumentString) 'INFO'

    try {
        $process = Start-Process `
            -FilePath 'winget.exe' `
            -ArgumentList $argumentString `
            -RedirectStandardOutput $stdOutPath `
            -RedirectStandardError $stdErrPath `
            -PassThru `
            -Wait `
            -WindowStyle Hidden
    }
    catch {
        throw "Failed to launch winget.exe: $($_.Exception.Message)"
    }

    return Get-WingetProcessResult -Process $process -StdOutPath $stdOutPath -StdErrPath $stdErrPath -StreamToLog:$StreamToLog
}

function Start-WingetProcessAsync {
    param(
        [Parameter(Mandatory = $true)][string[]]$Arguments,
        [switch]$CaptureOutput,
        [switch]$StreamToLog,
        [string]$FriendlyName = 'winget',
        [Parameter(Mandatory = $true)][scriptblock]$OnCompleted
    )

    $argumentString = ConvertTo-CommandLineArgumentString -Arguments $Arguments
    $stdOutPath = New-WingetLogPath -Prefix 'stdout'
    $stdErrPath = New-WingetLogPath -Prefix 'stderr'
    $activeWingetOperations = $script:ActiveWingetOperations
    $writeLogFn = (Get-Command -Name 'Write-Log' -CommandType Function -ErrorAction Stop).ScriptBlock
    $getWingetProcessResultFn = (Get-Command -Name 'Get-WingetProcessResult' -CommandType Function -ErrorAction Stop).ScriptBlock

    & $writeLogFn ("Starting {0}: winget {1}" -f $FriendlyName, $argumentString) 'INFO'

    try {
        $process = Start-Process `
            -FilePath 'winget.exe' `
            -ArgumentList $argumentString `
            -RedirectStandardOutput $stdOutPath `
            -RedirectStandardError $stdErrPath `
            -PassThru `
            -WindowStyle Hidden
    }
    catch {
        & $OnCompleted $null $_.Exception
        return
    }

    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(250)

    $operation = [pscustomobject]@{
        Process      = $process
        Timer        = $timer
        StdOutPath   = $stdOutPath
        StdErrPath   = $stdErrPath
        StreamToLog  = [bool]$StreamToLog
        FriendlyName = $FriendlyName
        OnCompleted  = $OnCompleted
    }

    if ($activeWingetOperations) {
        [void]$activeWingetOperations.Add($operation)
    }

    $timer.Add_Tick({
        $isExited = $false

        try {
            $isExited = $operation.Process.HasExited
        }
        catch {
            $isExited = $true
        }

        if (-not $isExited) {
            return
        }

        $operation.Timer.Stop()

        $result = $null
        $callbackError = $null

        try {
            $result = & $getWingetProcessResultFn -Process $operation.Process -StdOutPath $operation.StdOutPath -StdErrPath $operation.StdErrPath -StreamToLog:$operation.StreamToLog
        }
        catch {
            $callbackError = $_.Exception
        }

        try {
            & $operation.OnCompleted $result $callbackError
        }
        catch {
            & $writeLogFn ("Unhandled async callback failure in {0}: {1}" -f $operation.FriendlyName, $_.Exception.Message) 'ERROR'
        }
        finally {
            if ($activeWingetOperations) {
                [void]$activeWingetOperations.Remove($operation)
            }
        }
    }.GetNewClosure())

    $timer.Start()
}

function ConvertFrom-WingetSearchOutput {
    param([Parameter(Mandatory = $true)][string]$Text)

    $cleanText = $Text
    $cleanText = $cleanText -replace "`r(?!`n)", "`n"
    $cleanText = [regex]::Replace($cleanText, '\x1B\[[0-9;?]*[ -/]*[@-~]', '')

    $lines = @(
        $cleanText -split "`r?`n" |
        ForEach-Object { $_.TrimEnd() } |
        Where-Object { $_ -and $_.Trim() }
    )

    if (-not $lines.Count) {
        return @()
    }

    $separatorIndex = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^\s*-{3,}\s*$') {
            $separatorIndex = $i
            break
        }
    }

    if ($separatorIndex -lt 1) {
        return @()
    }

    $header = $lines[$separatorIndex - 1]
    $columnMatches = [regex]::Matches($header, '\S+')

    if ($columnMatches.Count -lt 4) {
        return @()
    }

    $columnStarts = @()
    foreach ($match in $columnMatches) {
        $columnStarts += $match.Index
    }

    $dataLines = @($lines[($separatorIndex + 1)..($lines.Count - 1)])
    $items = New-Object System.Collections.Generic.List[object]

    foreach ($line in $dataLines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        if ($line -match '^\s*-{3,}\s*$') { continue }
        if ($line -match '(?i)^No package found') { continue }

        $padded = $line.PadRight([Math]::Max($line.Length, $header.Length))
        $values = @()

        for ($index = 0; $index -lt $columnStarts.Count; $index++) {
            $start = $columnStarts[$index]
            if ($start -ge $padded.Length) {
                $values += ''
                continue
            }

            if ($index -lt ($columnStarts.Count - 1)) {
                $length = $columnStarts[$index + 1] - $start
                if ($length -lt 0) { $length = 0 }
                $values += $padded.Substring($start, [Math]::Min($length, $padded.Length - $start)).Trim()
            }
            else {
                $values += $padded.Substring($start).Trim()
            }
        }

        if ($values.Count -lt 4) { continue }

        $name = $values[0]
        $id = $values[1]
        $version = $values[2]

        if ([string]::IsNullOrWhiteSpace($name) -or [string]::IsNullOrWhiteSpace($id)) {
            continue
        }

        # Search is already restricted to $DefaultWingetSource, so do not trust parsed "source" text.
        [void]$items.Add((New-PackageItem -Name $name -Id $id -Version $version -Source $DefaultWingetSource -Origin 'Search'))
    }

    return @($items.ToArray())
}

function Search-WingetRepository {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Query,

        [int]$Count = $SearchResultLimit
    )

    $trimmedQuery = $Query.Trim()

    if ([string]::IsNullOrWhiteSpace($trimmedQuery)) {
        throw 'Search query cannot be empty.'
    }

    $args = @(
        'search',
        '--source', $DefaultWingetSource,
        '--accept-source-agreements',
        '--disable-interactivity',
        '--count', [string]$Count,
        '--query', $trimmedQuery
    )

    $result = Start-WingetProcess -Arguments $args -CaptureOutput -FriendlyName 'repository search'
    $allText = $result.StdOut + "`n" + $result.StdErr
    $items = @(ConvertFrom-WingetSearchOutput -Text $allText)

    if ($result.ExitCode -ne 0 -and $items.Count -eq 0 -and ($allText -notmatch '(?i)No package found')) {
        throw "winget search failed with exit code $($result.ExitCode)."
    }

    return @($items)
}

function Resolve-PackageInstallIdentity {
    param([Parameter(Mandatory = $true)]$Package)

    if ($Package.Id -and $Package.Id -notmatch '[…]{1}|\.\.\.$') {
        return [pscustomobject]@{ Mode = 'Id'; Value = $Package.Id }
    }

    try {
        $matches = @(Search-WingetRepository -Query $Package.Name | Where-Object { $_.Name -eq $Package.Name })

        if ($matches.Count -eq 1 -and $matches[0].Id -and $matches[0].Id -notmatch '[…]{1}|\.\.\.$') {
            return [pscustomobject]@{ Mode = 'Id'; Value = $matches[0].Id }
        }
    }
    catch {
        Write-Log "Search-based identity resolution failed for '$($Package.Name)': $($_.Exception.Message)" 'WARN'
    }

    return [pscustomobject]@{ Mode = 'Name'; Value = $Package.Name }
}

function Install-PackageSelection {
    param(
        [Parameter(Mandatory = $true)]$Package,
        [bool]$Silent
    )

    $resolved = Resolve-PackageInstallIdentity -Package $Package
    $wingetLog = New-WingetLogPath -Prefix 'install'

    $args = @(
        'install',
        '--source', ($(if ([string]::IsNullOrWhiteSpace([string]$Package.Source)) { $DefaultWingetSource } else { [string]$Package.Source })),
        '--accept-source-agreements',
        '--accept-package-agreements',
        '--log', $wingetLog
    )

    if ($resolved.Mode -eq 'Id') {
        $args += @('--id', $resolved.Value, '--exact')
    }
    else {
        $args += @('--name', $resolved.Value, '--exact')
    }

    if ($Silent) {
        $args += '--silent'
    }
    else {
        $args += '--interactive'
    }

    Write-Log ("Installing '{0}' using {1} '{2}'. Native winget log: {3}" -f $Package.Name, $resolved.Mode, $resolved.Value, $wingetLog)
    $result = Start-WingetProcess -Arguments $args -StreamToLog -FriendlyName ("install {0}" -f $Package.Name)

    if ($result.ExitCode -eq 0) {
        Write-Log ("Installed '{0}' successfully." -f $Package.Name) 'SUCCESS'
        return $true
    }

    if (Test-WingetInstallNoOpSuccess -ExitCode $result.ExitCode -StdOut $result.StdOut -StdErr $result.StdErr) {
        Write-Log ("'{0}' is already installed and no applicable newer version is available. Treating as success." -f $Package.Name) 'SUCCESS'
        return $true
    }

    Write-Log ("Failed to install '{0}'. Exit code: {1}" -f $Package.Name, $result.ExitCode) 'ERROR'
    return $false
}

function ConvertFrom-WingetUpgradeListOutput {
    param([Parameter(Mandatory = $true)][string]$Text)

    $cleanText = $Text
    $cleanText = $cleanText -replace "`r(?!`n)", "`n"
    $cleanText = [regex]::Replace($cleanText, '\x1B\[[0-9;?]*[ -/]*[@-~]', '')

    $lines = @(
        $cleanText -split "`r?`n" |
        ForEach-Object { $_.TrimEnd() } |
        Where-Object { $_ -and $_.Trim() }
    )

    if (-not $lines.Count) {
        return @()
    }

    $separatorIndex = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^\s*-{3,}\s*$') {
            $separatorIndex = $i
            break
        }
    }

    if ($separatorIndex -lt 1) {
        return @()
    }

    $header = $lines[$separatorIndex - 1]
    $columnMatches = [regex]::Matches($header, '\S+')

    if ($columnMatches.Count -lt 4) {
        return @()
    }

    $columnStarts = @()
    foreach ($match in $columnMatches) {
        $columnStarts += $match.Index
    }

    $dataLines = @($lines[($separatorIndex + 1)..($lines.Count - 1)])
    $items = New-Object System.Collections.Generic.List[object]

    foreach ($line in $dataLines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        if ($line -match '^\s*-{3,}\s*$') { continue }
        if ($line -match '(?i)^No package found') { continue }
        if ($line -match '(?i)^\d+\s+upgrades?\s+available') { continue }
        if ($line -match '(?i)^\d+\s+package\(s\)\s+have\s+pins?') { continue }
        if ($line -match '(?i)^The following packages have') { continue }
        if ($line -match '(?i)^Name\s+Id\s+Version\s+Available') { continue }

        $padded = $line.PadRight([Math]::Max($line.Length, $header.Length))
        $values = @()

        for ($index = 0; $index -lt $columnStarts.Count; $index++) {
            $start = $columnStarts[$index]
            if ($start -ge $padded.Length) {
                $values += ''
                continue
            }

            if ($index -lt ($columnStarts.Count - 1)) {
                $length = $columnStarts[$index + 1] - $start
                if ($length -lt 0) { $length = 0 }
                $values += $padded.Substring($start, [Math]::Min($length, $padded.Length - $start)).Trim()
            }
            else {
                $values += $padded.Substring($start).Trim()
            }
        }

        if ($values.Count -lt 4) { continue }

        $name = $values[0]
        $id = $values[1]
        $version = $values[2]
        $available = $values[3]
        $source = if ($values.Count -ge 5) { $values[$values.Count - 1] } else { '' }

        if ([string]::IsNullOrWhiteSpace($name) -or [string]::IsNullOrWhiteSpace($id)) {
            continue
        }

        [void]$items.Add([pscustomobject]@{
            Name      = $name
            Id        = $id
            Version   = $version
            Available = $available
            Source    = $source
        })
    }

    return @($items.ToArray())
}

function Get-WingetUpgradeablePackages {
    $args = @(
        'list',
        '--upgrade-available',
        '--accept-source-agreements',
        '--disable-interactivity'
    )

    Write-Log 'Enumerating installed packages with available upgrades.'
    $result = Start-WingetProcess -Arguments $args -CaptureOutput -FriendlyName 'list upgrade-available'
    $items = ConvertFrom-WingetUpgradeListOutput -Text ($result.StdOut + "`n" + $result.StdErr)

    if ($result.ExitCode -ne 0 -and $items.Count -eq 0 -and (($result.StdOut + $result.StdErr) -notmatch '(?i)(No package found|No installed package found|No applicable upgrade found|No installed package matching input criteria)')) {
        throw "winget list --upgrade-available failed with exit code $($result.ExitCode)."
    }

    return $items
}

function Invoke-WingetUpdateAll {
    param(
        [bool]$Silent,
        [scriptblock]$ProgressCallback
    )

    $packages = @(Get-WingetUpgradeablePackages)
    Write-Log ("Upgradeable package count: {0}" -f $packages.Count)

    if (-not $packages.Count) {
        Write-Log 'No upgrades are currently available.' 'INFO'
        return $true
    }

    $successCount = 0
    $failureCount = 0

    foreach ($package in $packages) {
        if ($ProgressCallback) {
            & $ProgressCallback ("Upgrading {0}..." -f $package.Name)
        }

        $wingetLog = New-WingetLogPath -Prefix 'upgrade'
        $args = @(
            'upgrade',
            '--id', $package.Id,
            '--exact',
            '--accept-source-agreements',
            '--accept-package-agreements',
            '--log', $wingetLog
        )

        if (-not [string]::IsNullOrWhiteSpace([string]$package.Source)) {
            $args += @('--source', [string]$package.Source)
        }

        if ($Silent) {
            $args += '--silent'
        }
        else {
            $args += '--interactive'
        }

        Write-Log ("Upgrading '{0}' ({1}) from {2} to {3}. Native winget log: {4}" -f $package.Name, $package.Id, $package.Version, $package.Available, $wingetLog)
        $result = Start-WingetProcess -Arguments $args -StreamToLog -FriendlyName ("upgrade {0}" -f $package.Name)

        if ($result.ExitCode -eq 0) {
            $successCount++
            Write-Log ("Upgraded '{0}' successfully." -f $package.Name) 'SUCCESS'
        }
        else {
            $failureCount++
            Write-Log ("Failed to upgrade '{0}'. Exit code: {1}" -f $package.Name, $result.ExitCode) 'ERROR'
        }
    }

    if ($failureCount -eq 0) {
        Write-Log ("Upgrade all completed successfully. Upgraded {0} package(s)." -f $successCount) 'SUCCESS'
        return $true
    }

    Write-Log ("Upgrade all completed with errors. Successes: {0}. Failures: {1}." -f $successCount, $failureCount) 'ERROR'
    return $false
}

function Test-WingetSourceName {
    param([string]$Source)

    if ([string]::IsNullOrWhiteSpace($Source)) {
        return $false
    }

    switch -Regex ($Source.Trim()) {
        '^(winget|msstore|winget-font)$' { return $true }
        default { return $false }
    }
}

function Get-SafePackageSource {
    param($Package)

    $candidate = [string]$Package.Source

    if (Test-WingetSourceName -Source $candidate) {
        return $candidate.Trim()
    }

    if ($Package.Origin -eq 'Search') {
        Write-Log ("Invalid parsed source '{0}' for '{1}'. Falling back to default source '{2}'." -f $candidate, $Package.Name, $DefaultWingetSource) 'WARN'
        return $DefaultWingetSource
    }

    if ($Package.Origin -eq 'Startup') {
        return $DefaultWingetSource
    }

    return $DefaultWingetSource
}

$script:WingetExit_UpdateNotApplicable = -1978335189

function Test-WingetInstallNoOpSuccess {
    param(
        [AllowNull()]$ExitCode,
        [string]$StdOut,
        [string]$StdErr
    )

    if ($null -ne $ExitCode -and "$ExitCode" -ne '') {
        try {
            if ([int]$ExitCode -eq $script:WingetExit_UpdateNotApplicable) {
                return $true
            }
        }
        catch {}
    }

    $combinedOutput = @($StdOut, $StdErr) -join "`n"
    return ($combinedOutput -match '(?im)(No applicable update found|No applicable upgrade found|No available upgrade found|No newer package versions? are available)')
}

#endregion

#region UI
$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$AppTitle"
        Width="$WindowWidth"
        Height="$WindowHeight"
        MinWidth="1180"
        MinHeight="760"
        WindowStartupLocation="CenterScreen"
        Background="#0B1118"
        Foreground="#F5F7FA"
        FontFamily="Segoe UI"
        ResizeMode="CanResize"
        SnapsToDevicePixels="True"
        UseLayoutRounding="True">
    <Window.Resources>
        <SolidColorBrush x:Key="WindowBackgroundBrush" Color="#0B1118"/>
        <SolidColorBrush x:Key="PanelBrush" Color="#111827"/>
        <SolidColorBrush x:Key="PanelBrushSoft" Color="#162031"/>
        <SolidColorBrush x:Key="PanelBorderBrush" Color="#263244"/>
        <SolidColorBrush x:Key="AccentBrush" Color="#4F8CFF"/>
        <SolidColorBrush x:Key="AccentBrushHover" Color="#6AA0FF"/>
        <SolidColorBrush x:Key="TextMutedBrush" Color="#8FA3BF"/>
        <SolidColorBrush x:Key="InputBrush" Color="#0E1622"/>
        <SolidColorBrush x:Key="SuccessBrush" Color="#1F9D67"/>

        <Style TargetType="Border" x:Key="CardBorderStyle">
            <Setter Property="Background" Value="{StaticResource PanelBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource PanelBorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="18"/>
            <Setter Property="Padding" Value="18"/>
            <Setter Property="Margin" Value="0,0,0,16"/>
        </Style>

        <Style TargetType="Button">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Background" Value="{StaticResource PanelBrushSoft}"/>
            <Setter Property="BorderBrush" Value="{StaticResource PanelBorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="18,8"/>
            <Setter Property="Margin" Value="0,0,10,0"/>
            <Setter Property="MinHeight" Value="34"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="ButtonBorder"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="12"
                                SnapsToDevicePixels="True">
                            <ContentPresenter
                                Margin="{TemplateBinding Padding}"
                                HorizontalAlignment="Center"
                                VerticalAlignment="Center"
                                RecognizesAccessKey="True"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="ButtonBorder" Property="Background" Value="{StaticResource AccentBrush}"/>
                                <Setter TargetName="ButtonBorder" Property="BorderBrush" Value="{StaticResource AccentBrushHover}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="ButtonBorder" Property="Opacity" Value="0.92"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="ButtonBorder" Property="Opacity" Value="0.45"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="TextBox">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderBrush" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="Margin" Value="0,0,10,0"/>
            <Setter Property="MinHeight" Value="34"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="CaretBrush" Value="White"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>

        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Margin" Value="0,0,12,0"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
        </Style>

        <Style TargetType="ListView">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
        </Style>

        <Style TargetType="ListViewItem">
            <Setter Property="Padding" Value="0"/>
            <Setter Property="Margin" Value="0,0,0,8"/>
            <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListViewItem">
                        <Border x:Name="ItemBorder" Background="#0E1622" BorderBrush="#1A2333" BorderThickness="1" CornerRadius="14" Padding="12">
                            <ContentPresenter/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="ItemBorder" Property="BorderBrush" Value="#40669F"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="ItemBorder" Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
                                <Setter TargetName="ItemBorder" Property="Background" Value="#142338"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="DarkScrollBarThumbStyle" TargetType="Thumb">
            <Setter Property="OverridesDefaultStyle" Value="True"/>
            <Setter Property="IsTabStop" Value="False"/>
            <Setter Property="Focusable" Value="False"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Thumb">
                        <Border Background="#46556D" CornerRadius="6" Margin="2"/>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="DarkScrollBarPageButtonStyle" TargetType="RepeatButton">
            <Setter Property="OverridesDefaultStyle" Value="True"/>
            <Setter Property="Focusable" Value="False"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="RepeatButton">
                        <Border Background="Transparent"/>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="ScrollBar">
            <Setter Property="Background" Value="#0B1220"/>
            <Setter Property="Width" Value="12"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ScrollBar">
                        <Border Background="{TemplateBinding Background}" CornerRadius="6" SnapsToDevicePixels="True">
                            <Track x:Name="PART_Track" IsDirectionReversed="True">
                                <Track.DecreaseRepeatButton>
                                    <RepeatButton Style="{StaticResource DarkScrollBarPageButtonStyle}" Command="ScrollBar.PageUpCommand"/>
                                </Track.DecreaseRepeatButton>
                                <Track.Thumb>
                                    <Thumb Style="{StaticResource DarkScrollBarThumbStyle}"/>
                                </Track.Thumb>
                                <Track.IncreaseRepeatButton>
                                    <RepeatButton Style="{StaticResource DarkScrollBarPageButtonStyle}" Command="ScrollBar.PageDownCommand"/>
                                </Track.IncreaseRepeatButton>
                            </Track>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="ProgressBar">
            <Setter Property="Foreground" Value="{StaticResource AccentBrush}"/>
            <Setter Property="Background" Value="#0E1622"/>
            <Setter Property="Height" Value="8"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Margin" Value="0,0,0,0"/>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid x:Name="MainContent" Margin="22">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <Border Grid.Row="0" Style="{StaticResource CardBorderStyle}" Padding="22" Margin="0,0,0,18">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>

                    <StackPanel Grid.Column="0">
                        <TextBlock Text="$AppTitle" FontSize="26" FontWeight="SemiBold"/>
                        <TextBlock Margin="0,6,0,0" Text="Version 2.2 (April 2026)" Foreground="{StaticResource TextMutedBrush}" FontSize="13"/>
                        <TextBlock Margin="0,8,0,0" Foreground="{StaticResource TextMutedBrush}" FontSize="12" x:Name="TxtLogPath"/>
                    </StackPanel>

                    <Border Grid.Column="1" Background="#0F1D2E" BorderBrush="#23314A" BorderThickness="1" CornerRadius="12" Padding="14,10" HorizontalAlignment="Right">
                        <StackPanel>
                            <TextBlock Text="Current selection" Foreground="{StaticResource TextMutedBrush}" FontSize="12"/>
                            <TextBlock x:Name="TxtCommonCount" Text="0 apps" FontSize="18" FontWeight="SemiBold"/>
                        </StackPanel>
                    </Border>
                </Grid>
            </Border>

            <Grid Grid.Row="1">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="1.05*"/>
                    <ColumnDefinition Width="18"/>
                    <ColumnDefinition Width="1.35*"/>
                </Grid.ColumnDefinitions>

                <Border Grid.Column="0" Style="{StaticResource CardBorderStyle}">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>

                        <StackPanel Grid.Row="0" Margin="0,0,0,14">
                            <TextBlock Text="Current selection" FontSize="20" FontWeight="SemiBold"/>
                            <TextBlock Margin="0,4,0,0" Text="Checked apps stay here across searches. Uncheck an item to remove it from the install basket." Foreground="{StaticResource TextMutedBrush}" FontSize="12"/>
                        </StackPanel>

                        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,14">
                            <Button x:Name="BtnClearAllCommon" Content="Clear selection"/>
                        </StackPanel>

                        <ListView x:Name="LvCommon" Grid.Row="2" ScrollViewer.VerticalScrollBarVisibility="Auto" ScrollViewer.HorizontalScrollBarVisibility="Disabled">
                            <ListView.ItemTemplate>
                                <DataTemplate>
                                    <Grid>
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="Auto"/>
                                            <ColumnDefinition Width="*"/>
                                        </Grid.ColumnDefinitions>
                                        <CheckBox Grid.Column="0" IsChecked="{Binding Selected, Mode=TwoWay}" Margin="0,0,14,0" VerticalAlignment="Center"/>
                                        <StackPanel Grid.Column="1">
                                            <TextBlock Text="{Binding Name}" FontSize="14" FontWeight="SemiBold" TextTrimming="CharacterEllipsis"/>
                                            <TextBlock Margin="0,4,0,0" Text="{Binding Id}" Foreground="{StaticResource TextMutedBrush}" FontSize="12" TextTrimming="CharacterEllipsis"/>
                                        </StackPanel>
                                    </Grid>
                                </DataTemplate>
                            </ListView.ItemTemplate>
                        </ListView>
                    </Grid>
                </Border>

                <Border Grid.Column="2" Style="{StaticResource CardBorderStyle}">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>

                        <StackPanel Grid.Row="0" Margin="0,0,0,14">
                            <TextBlock Text="Search repository" FontSize="20" FontWeight="SemiBold"/>
                            <TextBlock Margin="0,4,0,0"
                                    Text="Search the winget source by package name or ID."
                                    Foreground="{StaticResource TextMutedBrush}"
                                    FontSize="12"/>
                        </StackPanel>

                        <Grid Grid.Row="1" Margin="0,0,0,14">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>

                            <Border Grid.Column="0"
                                    MinWidth="320"
                                    Height="34"
                                    Margin="0,0,10,0"
                                    Background="{StaticResource InputBrush}"
                                    BorderBrush="{StaticResource PanelBorderBrush}"
                                    BorderThickness="1"
                                    CornerRadius="12"
                                    SnapsToDevicePixels="True">
                                <TextBox x:Name="TxtSearch"
                                        Margin="0"
                                        Height="34"
                                        AcceptsReturn="False"
                                        TextWrapping="NoWrap"
                                        VerticalScrollBarVisibility="Disabled"
                                        HorizontalScrollBarVisibility="Auto"
                                        ToolTip="$RepositorySearchHint"/>
                            </Border>

                            <Button x:Name="BtnSearch"
                                    Grid.Column="1"
                                    Content="Search"
                                    Margin="12,0,0,0"/>

                            <Button x:Name="BtnClearSearch"
                                    Grid.Column="2"
                                    Content="Show common apps"
                                    Margin="12,0,0,0"/>
                        </Grid>

                        <ListView x:Name="LvSearch" Grid.Row="2" ScrollViewer.VerticalScrollBarVisibility="Auto" ScrollViewer.HorizontalScrollBarVisibility="Disabled">
                            <ListView.ItemTemplate>
                                <DataTemplate>
                                    <Grid>
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="Auto"/>
                                            <ColumnDefinition Width="2.6*"/>
                                            <ColumnDefinition Width="1.8*"/>
                                            <ColumnDefinition Width="0.9*"/>
                                            <ColumnDefinition Width="0.9*"/>
                                        </Grid.ColumnDefinitions>

                                        <CheckBox Grid.Column="0" IsChecked="{Binding Selected, Mode=TwoWay}" Margin="0,0,14,0" VerticalAlignment="Center"/>
                                        <TextBlock Grid.Column="1" Text="{Binding Name}" FontSize="13" FontWeight="SemiBold" TextTrimming="CharacterEllipsis" VerticalAlignment="Center"/>
                                        <TextBlock Grid.Column="2" Text="{Binding Id}" FontSize="12" Foreground="{StaticResource TextMutedBrush}" TextTrimming="CharacterEllipsis" VerticalAlignment="Center" Margin="12,0,0,0"/>
                                        <TextBlock Grid.Column="3" Text="{Binding Version}" FontSize="12" Foreground="{StaticResource TextMutedBrush}" VerticalAlignment="Center" Margin="12,0,0,0"/>
                                        <TextBlock Grid.Column="4" Text="{Binding Source}" FontSize="12" Foreground="{StaticResource TextMutedBrush}" VerticalAlignment="Center" Margin="12,0,0,0"/>
                                    </Grid>
                                </DataTemplate>
                            </ListView.ItemTemplate>
                        </ListView>
                    </Grid>
                </Border>
            </Grid>

            <Border Grid.Row="2" Style="{StaticResource CardBorderStyle}" Margin="0,2,0,0">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <Grid Grid.Row="0" Margin="0,0,0,14">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>

                        <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                            <CheckBox x:Name="ChkSilent" Content="Silent installation" IsChecked="$DefaultSilentInstall" Margin="0,0,20,0"/>
                            <TextBlock x:Name="LblSelectedCount" Text="0 selected" VerticalAlignment="Center" Foreground="{StaticResource TextMutedBrush}" FontSize="12"/>
                        </StackPanel>

                        <Button x:Name="BtnOpenLogs" Grid.Column="1" Content="Open log folder"/>
                        <Button x:Name="BtnInstall" Grid.Column="2" Content="Install selected"/>
                        <Button x:Name="BtnUpdateAll" Grid.Column="3" Content="Update all installed packages" Visibility="Visible"/>
                        <Button x:Name="BtnClose" Grid.Column="4" Content="Close" Margin="0,0,0,0"/>
                    </Grid>

                    <StackPanel Grid.Row="1">
                        <TextBlock x:Name="LblStatus" Text="Ready" Foreground="{StaticResource TextMutedBrush}" FontSize="12" Margin="0,0,0,8"/>
                        <ProgressBar x:Name="ProgressBar" Minimum="0" Maximum="100" Value="0"/>
                    </StackPanel>
                </Grid>
            </Border>
        </Grid>

        <Grid x:Name="BusyOverlay"
              Visibility="Collapsed"
              IsHitTestVisible="False"
              Background="#8C0B1118"
              Panel.ZIndex="999">

            <Border Width="300"
                    Padding="24"
                    HorizontalAlignment="Center"
                    VerticalAlignment="Center"
                    Background="#E6111827"
                    BorderBrush="#2B3950"
                    BorderThickness="1"
                    CornerRadius="20">
                <StackPanel HorizontalAlignment="Center">
                    <Ellipse Width="56"
                             Height="56"
                             Stroke="{StaticResource AccentBrush}"
                             StrokeThickness="5"
                             StrokeDashArray="1 2"
                             StrokeDashCap="Round"
                             HorizontalAlignment="Center"
                             RenderTransformOrigin="0.5,0.5">
                        <Ellipse.RenderTransform>
                            <RotateTransform Angle="0"/>
                        </Ellipse.RenderTransform>
                        <Ellipse.Triggers>
                            <EventTrigger RoutedEvent="FrameworkElement.Loaded">
                                <BeginStoryboard>
                                    <Storyboard RepeatBehavior="Forever">
                                        <DoubleAnimation
                                            Storyboard.TargetProperty="(UIElement.RenderTransform).(RotateTransform.Angle)"
                                            From="0"
                                            To="360"
                                            Duration="0:0:1"/>
                                    </Storyboard>
                                </BeginStoryboard>
                            </EventTrigger>
                        </Ellipse.Triggers>
                    </Ellipse>

                    <TextBlock Text="Please wait"
                               Margin="0,16,0,6"
                               FontSize="18"
                               FontWeight="SemiBold"
                               HorizontalAlignment="Center"/>

                    <TextBlock x:Name="BusyOverlayMessage"
                               Text="Working..."
                               FontSize="12"
                               Foreground="{StaticResource TextMutedBrush}"
                               TextAlignment="Center"
                               TextWrapping="Wrap"
                               HorizontalAlignment="Center"/>
                </StackPanel>
            </Border>
        </Grid>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader ([xml]$Xaml)
$script:Window = [Windows.Markup.XamlReader]::Load($reader)

$script:TxtLogPath = $script:Window.FindName('TxtLogPath')
$script:TxtCommonCount = $script:Window.FindName('TxtCommonCount')
$script:LvCommon = $script:Window.FindName('LvCommon')
$script:LvSearch = $script:Window.FindName('LvSearch')
$script:BtnClearAllCommon = $script:Window.FindName('BtnClearAllCommon')
$script:TxtSearch = $script:Window.FindName('TxtSearch')
$script:BtnSearch = $script:Window.FindName('BtnSearch')
$script:BtnClearSearch = $script:Window.FindName('BtnClearSearch')
$script:ChkSilent = $script:Window.FindName('ChkSilent')
$script:LblSelectedCount = $script:Window.FindName('LblSelectedCount')
$script:BtnOpenLogs = $script:Window.FindName('BtnOpenLogs')
$script:BtnInstall = $script:Window.FindName('BtnInstall')
$script:BtnUpdateAll = $script:Window.FindName('BtnUpdateAll')
$script:BtnClose = $script:Window.FindName('BtnClose')
$script:LblStatus = $script:Window.FindName('LblStatus')
$script:ProgressBar = $script:Window.FindName('ProgressBar')
$script:BusyOverlay = $script:Window.FindName('BusyOverlay')
$script:BusyOverlayMessage = $script:Window.FindName('BusyOverlayMessage')

if (-not $IncludeUpdateAllButton) {
    $script:BtnUpdateAll.Visibility = 'Collapsed'
}

$script:TxtLogPath.Text = "by zorrouraganu | Log: $script:LogPath"
$script:LvCommon.ItemsSource = $script:SelectedPackages
$script:LvSearch.ItemsSource = $script:ResultPackages
$script:TxtSearch.Text = ""
#endregion

#region UI actions
function Refresh-SelectionSummary {
    $count = @(Get-SelectedPackages).Count
    $script:LblSelectedCount.Text = ('{0} selected' -f $count)
}

function Refresh-CommonSummary {
    $script:TxtCommonCount.Text = ('{0} apps' -f $script:SelectedPackages.Count)
}

function Clear-AllCommon {
    foreach ($item in @($script:SelectedPackages)) {
        $item.Selected = $false
    }
    Sync-SelectedPackages
    Refresh-CommonSummary
    Refresh-SelectionSummary
}

function Clear-SearchResults {
    Set-ResultPackages -Items (Get-StartupCatalogPackages)
    $script:LblStatus.Text = 'Common apps loaded.'
    Write-Log 'Common apps loaded into results panel.' 'DEBUG'
}


function Invoke-RepositorySearchFromUi {
    if (-not (Ensure-Winget)) {
        return
    }

    $query = [string]$script:TxtSearch.Text
    $query = $query.Trim()

    if ([string]::IsNullOrWhiteSpace($query)) {
        $script:LblStatus.Text = 'Enter a package name or ID to search.'
        [System.Windows.MessageBox]::Show(
            'Enter a package name or ID to search.',
            $AppTitle,
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        ) | Out-Null
        $script:TxtSearch.Focus()
        return
    }

    Set-UiBusy -Busy $true -StatusText 'Searching repository...'
    Write-Log ("Repository search triggered. Query: {0}" -f $query)

    $args = @(
        'search',
        '--source', $DefaultWingetSource,
        '--accept-source-agreements',
        '--disable-interactivity',
        '--count', [string]$SearchResultLimit,
        '--query', $query
    )

    $convertFromWingetSearchOutputFn = (Get-Command -Name 'ConvertFrom-WingetSearchOutput' -CommandType Function -ErrorAction Stop).ScriptBlock
    $setResultPackagesFn = (Get-Command -Name 'Set-ResultPackages' -CommandType Function -ErrorAction Stop).ScriptBlock
    $setUiBusyFn = (Get-Command -Name 'Set-UiBusy' -CommandType Function -ErrorAction Stop).ScriptBlock
    $syncSelectedPackagesFn = (Get-Command -Name 'Sync-SelectedPackages' -CommandType Function -ErrorAction Stop).ScriptBlock
    $refreshCommonSummaryFn = (Get-Command -Name 'Refresh-CommonSummary' -CommandType Function -ErrorAction Stop).ScriptBlock
    $refreshSelectionSummaryFn = (Get-Command -Name 'Refresh-SelectionSummary' -CommandType Function -ErrorAction Stop).ScriptBlock
    $writeLogFn = (Get-Command -Name 'Write-Log' -CommandType Function -ErrorAction Stop).ScriptBlock
    $lblStatus = $script:LblStatus
    $logPath = $script:LogPath
    $appTitleLocal = $AppTitle
    $searchQuery = $query

    Start-WingetProcessAsync -Arguments $args -CaptureOutput -FriendlyName 'repository search' -OnCompleted ({
        param($result, $callbackError)

        try {
            if ($callbackError) {
                throw $callbackError
            }

            $allText = $result.StdOut + "`n" + $result.StdErr
            $results = @(& $convertFromWingetSearchOutputFn -Text $allText)

            if ($result.ExitCode -ne 0 -and $results.Count -eq 0 -and ($allText -notmatch '(?i)No package found')) {
                throw "winget search failed with exit code $($result.ExitCode)."
            }

            & $setResultPackagesFn -Items $results

            if ($results.Count -gt 0) {
                $lblStatus.Text = ('Search complete. {0} result(s).' -f $results.Count)
                & $writeLogFn ("Repository search complete. Results: {0}" -f $results.Count) 'SUCCESS'
            }
            else {
                $lblStatus.Text = 'No packages found.'
                & $writeLogFn ("Repository search returned no results for query: {0}" -f $searchQuery) 'INFO'
            }
        }
        catch {
            $lblStatus.Text = 'Search failed.'
            & $writeLogFn "Repository search failed: $($_.Exception.Message)" 'ERROR'
            [System.Windows.MessageBox]::Show(
                "Search failed.`n`n$($_.Exception.Message)`n`nLog: $logPath",
                $appTitleLocal,
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            ) | Out-Null
        }
        finally {
            & $setUiBusyFn -Busy $false -StatusText $lblStatus.Text
            & $syncSelectedPackagesFn
            & $refreshCommonSummaryFn
            & $refreshSelectionSummaryFn
        }
    }.GetNewClosure())
}

function New-InstallPackageInvocation {
    param(
        [Parameter(Mandatory = $true)]$Package,
        [bool]$Silent
    )

    $resolved = Resolve-PackageInstallIdentity -Package $Package
    $wingetLog = New-WingetLogPath -Prefix 'install'

    $args = @(
        'install',
        '--source', ($(if ([string]::IsNullOrWhiteSpace([string]$Package.Source)) { $DefaultWingetSource } else { [string]$Package.Source })),
        '--accept-source-agreements',
        '--accept-package-agreements',
        '--log', $wingetLog
    )

    if ($resolved.Mode -eq 'Id') {
        $args += @('--id', $resolved.Value, '--exact')
    }
    else {
        $args += @('--name', $resolved.Value, '--exact')
    }

    if ($Silent) {
        $args += '--silent'
    }
    else {
        $args += '--interactive'
    }

    Write-Log ("Installing '{0}' using {1} '{2}'. Native winget log: {3}" -f $Package.Name, $resolved.Mode, $resolved.Value, $wingetLog)

    return [pscustomobject]@{
        Arguments    = $args
        FriendlyName = ('install {0}' -f $Package.Name)
        Package      = $Package
    }
}

function Complete-InstallSelectedOperation {
    param(
        [string]$StatusText,
        [string]$StatusLevel = 'INFO'
    )

    $script:LblStatus.Text = $StatusText
    Write-Log $StatusText $StatusLevel

    $script:InstallSelectionState = $null

    Set-UiBusy -Busy $false -StatusText $script:LblStatus.Text
    Sync-SelectedPackages
    Refresh-CommonSummary
    Refresh-SelectionSummary
}

function Fail-InstallSelectedOperation {
    param([System.Exception]$Exception)

    $script:InstallSelectionState = $null
    $script:LblStatus.Text = 'Install operation failed.'
    Write-Log ("Install operation failed: {0}" -f $Exception.Message) 'ERROR'
    [System.Windows.MessageBox]::Show(
        "Install operation failed.`n`n$($Exception.Message)`n`nLog: $script:LogPath",
        $AppTitle,
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Error
    ) | Out-Null

    Set-UiBusy -Busy $false -StatusText $script:LblStatus.Text
    Sync-SelectedPackages
    Refresh-CommonSummary
    Refresh-SelectionSummary
}

function Invoke-InstallNextSelectedPackage {
    $state = $script:InstallSelectionState
    if (-not $state) {
        return
    }

    if ($state.Index -ge $state.Total) {
        $statusText = 'Install complete. Success: {0}. Failed: {1}.' -f $state.SuccessCount, $state.FailureCount
        $statusLevel = if ($state.FailureCount -gt 0) { 'WARN' } else { 'SUCCESS' }
        Complete-InstallSelectedOperation -StatusText $statusText -StatusLevel $statusLevel
        return
    }

    $package = $state.Packages[$state.Index]
    $packageNumber = $state.Index + 1
    Set-BusyMessage ('Installing {0}/{1}: {2}...' -f $packageNumber, $state.Total, $package.Name)

    try {
        $invocation = New-InstallPackageInvocation -Package $package -Silent $state.Silent
    }
    catch {
        Fail-InstallSelectedOperation -Exception $_.Exception
        return
    }

    $writeLogFn = (Get-Command -Name 'Write-Log' -CommandType Function -ErrorAction Stop).ScriptBlock
    $testWingetInstallNoOpSuccessFn = (Get-Command -Name 'Test-WingetInstallNoOpSuccess' -CommandType Function -ErrorAction Stop).ScriptBlock
    $invokeInstallNextSelectedPackageFn = (Get-Command -Name 'Invoke-InstallNextSelectedPackage' -CommandType Function -ErrorAction Stop).ScriptBlock
    $failInstallSelectedOperationFn = (Get-Command -Name 'Fail-InstallSelectedOperation' -CommandType Function -ErrorAction Stop).ScriptBlock

    Start-WingetProcessAsync -Arguments $invocation.Arguments -StreamToLog -FriendlyName $invocation.FriendlyName -OnCompleted ({
        param($result, $callbackError)

        try {
            if ($callbackError) {
                throw $callbackError
            }

            if (-not $state) {
                return
            }

            $currentPackage = $state.Packages[$state.Index]

            if ($result.ExitCode -eq 0) {
                $state.SuccessCount++
                & $writeLogFn ("Installed '{0}' successfully." -f $currentPackage.Name) 'SUCCESS'
            }
            elseif (& $testWingetInstallNoOpSuccessFn -ExitCode $result.ExitCode -StdOut $result.StdOut -StdErr $result.StdErr) {
                $state.SuccessCount++
                & $writeLogFn ("'{0}' is already installed and no applicable newer version is available. Treating as success." -f $currentPackage.Name) 'SUCCESS'
            }
            else {
                $state.FailureCount++
                & $writeLogFn ("Failed to install '{0}'. Exit code: {1}" -f $currentPackage.Name, $result.ExitCode) 'ERROR'
            }

            $state.Index++
            & $invokeInstallNextSelectedPackageFn
        }
        catch {
            & $failInstallSelectedOperationFn -Exception $_.Exception
        }
    }.GetNewClosure())
}

function Complete-UpdateAllOperation {
    param(
        [string]$StatusText,
        [string]$StatusLevel = 'INFO'
    )

    $script:LblStatus.Text = $StatusText
    Write-Log $StatusText $StatusLevel

    $script:UpdateAllState = $null
    Set-UiBusy -Busy $false -StatusText $script:LblStatus.Text
}

function Fail-UpdateAllOperation {
    param([System.Exception]$Exception)

    $script:UpdateAllState = $null
    $script:LblStatus.Text = 'Update-all operation failed.'
    Write-Log ("Update-all operation failed: {0}" -f $Exception.Message) 'ERROR'
    [System.Windows.MessageBox]::Show(
        "Update-all operation failed.`n`n$($Exception.Message)`n`nLog: $script:LogPath",
        $AppTitle,
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Error
    ) | Out-Null

    Set-UiBusy -Busy $false -StatusText $script:LblStatus.Text
}

function Invoke-UpdateAllNextPackage {
    $state = $script:UpdateAllState
    if (-not $state) {
        return
    }

    if ($state.Index -ge $state.Total) {
        $statusText = if ($state.FailureCount -eq 0) {
            'Update all completed successfully. Upgraded {0} package(s).' -f $state.SuccessCount
        }
        else {
            'Update all finished. Success: {0}. Failed: {1}.' -f $state.SuccessCount, $state.FailureCount
        }

        $statusLevel = if ($state.FailureCount -eq 0) { 'SUCCESS' } else { 'WARN' }
        Complete-UpdateAllOperation -StatusText $statusText -StatusLevel $statusLevel
        return
    }

    $package = $state.Packages[$state.Index]
    $packageNumber = $state.Index + 1
    Set-BusyMessage ('Upgrading {0}/{1}: {2}...' -f $packageNumber, $state.Total, $package.Name)

    $wingetLog = New-WingetLogPath -Prefix 'upgrade'
    $args = @(
        'upgrade',
        '--id', $package.Id,
        '--exact',
        '--accept-source-agreements',
        '--accept-package-agreements',
        '--log', $wingetLog
    )

    if (-not [string]::IsNullOrWhiteSpace([string]$package.Source)) {
        $args += @('--source', [string]$package.Source)
    }

    if ($state.Silent) {
        $args += '--silent'
    }
    else {
        $args += '--interactive'
    }

    Write-Log ("Upgrading '{0}' ({1}) from {2} to {3}. Native winget log: {4}" -f $package.Name, $package.Id, $package.Version, $package.Available, $wingetLog)

    $writeLogFn = (Get-Command -Name 'Write-Log' -CommandType Function -ErrorAction Stop).ScriptBlock
    $testWingetInstallNoOpSuccessFn = (Get-Command -Name 'Test-WingetInstallNoOpSuccess' -CommandType Function -ErrorAction Stop).ScriptBlock
    $invokeUpdateAllNextPackageFn = (Get-Command -Name 'Invoke-UpdateAllNextPackage' -CommandType Function -ErrorAction Stop).ScriptBlock
    $failUpdateAllOperationFn = (Get-Command -Name 'Fail-UpdateAllOperation' -CommandType Function -ErrorAction Stop).ScriptBlock

    Start-WingetProcessAsync -Arguments $args -StreamToLog -FriendlyName ('upgrade {0}' -f $package.Name) -OnCompleted ({
        param($result, $callbackError)

        try {
            if ($callbackError) {
                throw $callbackError
            }

            if (-not $state) {
                return
            }

            $currentPackage = $state.Packages[$state.Index]

            if ($result.ExitCode -eq 0) {
                $state.SuccessCount++
                & $writeLogFn ("Upgraded '{0}' successfully." -f $currentPackage.Name) 'SUCCESS'
            }
            elseif (& $testWingetInstallNoOpSuccessFn -ExitCode $result.ExitCode -StdOut $result.StdOut -StdErr $result.StdErr) {
                $state.SuccessCount++
                & $writeLogFn ("'{0}' no longer has an applicable upgrade. Treating as success." -f $currentPackage.Name) 'SUCCESS'
            }
            else {
                $state.FailureCount++
                & $writeLogFn ("Failed to upgrade '{0}'. Exit code: {1}" -f $currentPackage.Name, $result.ExitCode) 'ERROR'
            }

            $state.Index++
            & $invokeUpdateAllNextPackageFn
        }
        catch {
            & $failUpdateAllOperationFn -Exception $_.Exception
        }
    }.GetNewClosure())
}

function Start-WingetUpdateAllAsync {
    param([bool]$Silent)

    $args = @(
        'list',
        '--upgrade-available',
        '--accept-source-agreements',
        '--disable-interactivity'
    )

    Write-Log 'Enumerating installed packages with available upgrades.'

    $convertFromWingetUpgradeListOutputFn = (Get-Command -Name 'ConvertFrom-WingetUpgradeListOutput' -CommandType Function -ErrorAction Stop).ScriptBlock
    $writeLogFn = (Get-Command -Name 'Write-Log' -CommandType Function -ErrorAction Stop).ScriptBlock
    $completeUpdateAllOperationFn = (Get-Command -Name 'Complete-UpdateAllOperation' -CommandType Function -ErrorAction Stop).ScriptBlock
    $invokeUpdateAllNextPackageFn = (Get-Command -Name 'Invoke-UpdateAllNextPackage' -CommandType Function -ErrorAction Stop).ScriptBlock
    $failUpdateAllOperationFn = (Get-Command -Name 'Fail-UpdateAllOperation' -CommandType Function -ErrorAction Stop).ScriptBlock

    Start-WingetProcessAsync -Arguments $args -CaptureOutput -FriendlyName 'list upgrade-available' -OnCompleted ({
        param($result, $callbackError)

        try {
            if ($callbackError) {
                throw $callbackError
            }

            $allText = $result.StdOut + "`n" + $result.StdErr
            $packages = @(& $convertFromWingetUpgradeListOutputFn -Text $allText)

            if ($result.ExitCode -ne 0 -and $packages.Count -eq 0 -and ($allText -notmatch '(?i)(No package found|No installed package found|No applicable upgrade found|No installed package matching input criteria)')) {
                throw "winget list --upgrade-available failed with exit code $($result.ExitCode)."
            }

            & $writeLogFn ("Upgradeable package count: {0}" -f $packages.Count)

            if (-not $packages.Count) {
                & $completeUpdateAllOperationFn -StatusText 'No upgrades are currently available.' -StatusLevel 'INFO'
                return
            }

            $script:UpdateAllState = [pscustomobject]@{
                Packages     = @($packages)
                Index        = 0
                Total        = $packages.Count
                Silent       = $Silent
                SuccessCount = 0
                FailureCount = 0
            }

            & $invokeUpdateAllNextPackageFn
        }
        catch {
            & $failUpdateAllOperationFn -Exception $_.Exception
        }
    }.GetNewClosure())
}

function Invoke-InstallSelectedFromUi {
    if (-not (Ensure-Winget)) {
        return
    }

    Sync-SelectedPackages
    $selected = @(Get-SelectedPackages)
    if (-not $selected.Count) {
        [System.Windows.MessageBox]::Show(
            'Nothing is selected.',
            $AppTitle,
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        ) | Out-Null
        return
    }

    $silent = [bool]$script:ChkSilent.IsChecked

    Set-UiBusy -Busy $true -StatusText 'Preparing selected packages for installation...'
    Write-Log ("Install triggered. Package count: {0}. Silent: {1}" -f $selected.Count, $silent)

    $script:InstallSelectionState = [pscustomobject]@{
        Packages     = @($selected)
        Index        = 0
        Total        = $selected.Count
        Silent       = $silent
        SuccessCount = 0
        FailureCount = 0
    }

    Invoke-InstallNextSelectedPackage
}

function Invoke-UpdateAllFromUi {
    if (-not (Ensure-Winget)) {
        return
    }

    $silent = [bool]$script:ChkSilent.IsChecked
    $confirm = [System.Windows.MessageBox]::Show(
        'This will enumerate all installed packages with available upgrades and upgrade them one by one. Continue?',
        $AppTitle,
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )

    if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) {
        Write-Log 'Update-all action cancelled by user.' 'WARN'
        return
    }

    Set-UiBusy -Busy $true -StatusText 'Enumerating and upgrading installed packages... (This operation can take a long time, be patient.)'
    Write-Log ("Update-all triggered. Silent: {0}" -f $silent)

    Start-WingetUpdateAllAsync -Silent $silent
}
#endregion

#region Event wiring
$script:BtnClearAllCommon.Add_Click({ Clear-AllCommon })
$script:BtnSearch.Add_Click({ Invoke-RepositorySearchFromUi })
$script:BtnClearSearch.Add_Click({ Clear-SearchResults })
$script:BtnOpenLogs.Add_Click({ Start-Process explorer.exe $script:LogRoot })
$script:BtnInstall.Add_Click({ Invoke-InstallSelectedFromUi })
$script:BtnClose.Add_Click({ $script:Window.Close() })
if ($IncludeUpdateAllButton) {
    $script:BtnUpdateAll.Add_Click({ Invoke-UpdateAllFromUi })
}
$script:TxtSearch.Add_KeyDown({
    param($sender, $eventArgs)
    if ($eventArgs.Key -eq [System.Windows.Input.Key]::Enter -and -not $script:IsBusy -and -not ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Shift)) {
        $eventArgs.Handled = $true
        Invoke-RepositorySearchFromUi
    }
})

$selectionTimer = New-Object System.Windows.Threading.DispatcherTimer
$selectionTimer.Interval = [TimeSpan]::FromMilliseconds(350)
$selectionTimer.Add_Tick({ Sync-SelectedPackages; Refresh-CommonSummary; Refresh-SelectionSummary })
$selectionTimer.Start()

$script:Window.Add_SourceInitialized({ Set-ModernWindowChrome -Window $script:Window })
$script:Window.Add_Closing({
    param($sender, $eventArgs)

    if ($script:IsBusy) {
        $eventArgs.Cancel = $true
        Write-Log 'Close request ignored because an operation is still in progress.' 'WARN'
        return
    }

    Write-Log 'Window closing. Session ended.' 'INFO'
})
#endregion

#region Launch
Set-ResultPackages -Items (Get-StartupCatalogPackages)
Refresh-CommonSummary
Refresh-SelectionSummary
$script:LblStatus.Text = 'Ready.'

if (-not (Ensure-Winget)) {
    Write-Log 'Launching UI without a working winget backend.' 'WARN'
    $script:LblStatus.Text = 'winget is unavailable. Repair or install App Installer, then retry.'
}
else {
    Write-Log 'Backend checks completed successfully.' 'SUCCESS'
}

[void]$script:Window.ShowDialog()
#endregion
