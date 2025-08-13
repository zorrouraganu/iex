#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --- Optional fallback app list (used only if $Apps is not already defined) ---
if (-not (Get-Variable -Name Apps -Scope Script -ErrorAction SilentlyContinue)) {
    $script:Apps = @(
        [pscustomobject]@{ Display="Adobe Reader DC";           Candidates=@("Adobe.Acrobat.Reader.64-bit","Adobe.Acrobat.Reader.DC") }
        [pscustomobject]@{ Display="Bitwarden";                 Candidates=@("Bitwarden.Bitwarden") }
        [pscustomobject]@{ Display="EA Desktop";                Candidates=@("ElectronicArts.EADesktop","ElectronicArts.EAApp","ElectronicArts.EA") }
        [pscustomobject]@{ Display="GOG Galaxy";                Candidates=@("GOG.Galaxy") }
        [pscustomobject]@{ Display="Steam";                     Candidates=@("Valve.Steam") }
        [pscustomobject]@{ Display="Epic Store";                Candidates=@("EpicGames.EpicGamesLauncher") }
        [pscustomobject]@{ Display="VLC";                       Candidates=@("VideoLAN.VLC") }
        [pscustomobject]@{ Display="WinSCP";                    Candidates=@("WinSCP.WinSCP") }
        [pscustomobject]@{ Display="VSCode";                    Candidates=@("Microsoft.VisualStudioCode") }
        [pscustomobject]@{ Display="Firefox";                   Candidates=@("Mozilla.Firefox") }
        [pscustomobject]@{ Display="Chrome";                    Candidates=@("Google.Chrome") }
        [pscustomobject]@{ Display="7-zip";                     Candidates=@("7zip.7zip") }
        [pscustomobject]@{ Display="nVidia control panel";      Candidates=@("NVIDIACorp.NVIDIAControlPanel","NVIDIA.ControlPanel") }
        [pscustomobject]@{ Display="nVidia GeForce Experience"; Candidates=@("NVIDIA.GeForceExperience") }
        [pscustomobject]@{ Display="calibre";                   Candidates=@("Calibre.Calibre") }
        [pscustomobject]@{ Display="Logitech Options";          Candidates=@("Logitech.Options","Logitech.OptionsPlus") }
        [pscustomobject]@{ Display="Logitech G Hub";            Candidates=@("Logitech.GHUB") }
        [pscustomobject]@{ Display="Microsoft To Do";           Candidates=@("Microsoft.ToDo","9NBLGGH5R558") }
        [pscustomobject]@{ Display="Plex";                      Candidates=@("Plex.Plex") }
        [pscustomobject]@{ Display="Spotify";                   Candidates=@("Spotify.Spotify") }
        [pscustomobject]@{ Display="Apple Music";               Candidates=@("Apple.AppleMusic","Apple.Music") }
        [pscustomobject]@{ Display="WhatsApp";                  Candidates=@("WhatsApp.WhatsApp") }
        [pscustomobject]@{ Display="WinRAR";                    Candidates=@("RARLab.WinRAR") }
        [pscustomobject]@{ Display="PowerToys";                 Candidates=@("Microsoft.PowerToys") }
        [pscustomobject]@{ Display="WinAmp";                     Candidates=@("Winamp.Winamp") }
        [pscustomobject]@{ Display="Apple Devices";              Candidates=@("9NM4T8B9JQZ1","Apple.AppleMobileDeviceSupport") }
        [pscustomobject]@{ Display="battle.net";                 Candidates=@("Blizzard.BattleNet") }
        [pscustomobject]@{ Display="CPUID CPU-Z MSI";            Candidates=@("CPUID.CPU-Z.MSI") }
        [pscustomobject]@{ Display="discord";                    Candidates=@("Discord.Discord") }
        [pscustomobject]@{ Display="Philips Hue Sync";           Candidates=@("Philips.HueSync") }
        [pscustomobject]@{ Display="iCloud";                     Candidates=@("Apple.iCloud") }
        [pscustomobject]@{ Display="AIDA64 Extreme";             Candidates=@("FinalWire.AIDA64.Extreme") }
        [pscustomobject]@{ Display="KeePass";                    Candidates=@("DominikReichl.KeePass","KeePassXCTeam.KeePassXC") }
        [pscustomobject]@{ Display="MSI Center";                 Candidates=@("9NVMNJCR03XV") }
        [pscustomobject]@{ Display="Belgium e-ID middleware";    Candidates=@("BelgianGovernment.eIDmiddleware") }
        [pscustomobject]@{ Display="FFmpeg for yt-dlp";          Candidates=@("Gyan.FFmpeg","Gyan.FFmpeg.Essentials") }
        [pscustomobject]@{ Display="Notepad++";                  Candidates=@("Notepad++.Notepad++") }
        [pscustomobject]@{ Display="OBS Studio";                 Candidates=@("OBSProject.OBSStudio") }
        [pscustomobject]@{ Display="Plex Media Server";          Candidates=@("Plex.PlexMediaServer") }
        [pscustomobject]@{ Display="Private Internet Access";    Candidates=@("PrivateInternetAccess.PrivateInternetAccess") }
        [pscustomobject]@{ Display="NordVPN";                    Candidates=@("NordVPN.NordVPN","NordSecurity.NordVPN") }
        [pscustomobject]@{ Display="qBittorrent";                Candidates=@("qBittorrent.qBittorrent","qBittorrent.qBittorrent.Qt6") }
        [pscustomobject]@{ Display="Remote Desktop";             Candidates=@("Microsoft.RemoteDesktopClient","9WZDNCRFJ3PS") }
        [pscustomobject]@{ Display="SSHFS-Win";                  Candidates=@("SSHFS-Win.SSHFS-Win") }
        [pscustomobject]@{ Display="Tailscale";                  Candidates=@("tailscale.tailscale") }
        [pscustomobject]@{ Display="Ubisoft Connect";            Candidates=@("Ubisoft.Connect") }
        [pscustomobject]@{ Display="VirtualDJ";                  Candidates=@("AtomixProductions.VirtualDJ") }
        [pscustomobject]@{ Display="Webex";                      Candidates=@("Cisco.Webex","Cisco.CiscoWebexMeetings") }
        [pscustomobject]@{ Display="WeMod";                      Candidates=@("WeMod.WeMod") }
        [pscustomobject]@{ Display="WinFsp";                     Candidates=@("WinFsp.WinFsp") }
    )
}

# --- Winget helpers ---
function Test-Winget {
    $wg = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $wg) { throw "winget not found. Install 'App Installer' from Microsoft Store and re-run." }
}

function Write-Log {
    param([System.Windows.Controls.TextBox]$Log, [string]$Text)
    $stamp = (Get-Date).ToString("HH:mm:ss")
    $Log.AppendText("[$stamp] $Text`r`n")
    $Log.ScrollToEnd()
}

function Start-WingetProcess {
    param(
        [string[]]$Arguments,
        [System.Windows.Controls.TextBox]$Log,
        [switch]$QuietProgress = $true   # strip ASCII progress by default
    )
    $stdOut = [System.IO.Path]::GetTempFileName()
    $stdErr = [System.IO.Path]::GetTempFileName()
    try {
        Write-Log $Log ("winget " + ($Arguments -join ' '))
        $p = Start-Process -FilePath "winget" -ArgumentList $Arguments -PassThru -Wait -WindowStyle Hidden `
            -RedirectStandardOutput $stdOut -RedirectStandardError $stdErr

        $out = (Get-Content $stdOut -Raw -ErrorAction SilentlyContinue)
        $err = (Get-Content $stdErr -Raw -ErrorAction SilentlyContinue)

        if ($QuietProgress -and $out) {
            $out = ($out -split "`r?`n") |
                Where-Object {
                    $_ -and
                    ($_ -notmatch '^\s*[-\\\/\|\.]+\s*$') -and                      # spinner/bar rows
                    ($_ -notmatch '\b\d+(\.\d+)?\s*(KB|MB|GB)\s*/\s*\d+(\.\d+)?\s*(KB|MB|GB)\b') -and # 616 KB / 3.60 MB
                    ($_ -notmatch '^\s*\d{1,3}%\s*$')                                # plain percent rows
                } | ForEach-Object { $_.TrimEnd() } | Out-String
        }

        if ($out) { Write-Log $Log ($out.TrimEnd()) }
        if ($err) { Write-Log $Log ("ERROR: " + $err.TrimEnd()) }
        return $p.ExitCode
    }
    finally {
        Remove-Item $stdOut,$stdErr -ErrorAction SilentlyContinue
    }
}


function Install-App {
    param(
        [pscustomobject]$App,
        [switch]$Silent,
        [System.Windows.Controls.TextBox]$Log
    )
    $common = @('--accept-source-agreements','--accept-package-agreements')
    if ($Silent) { $common += '--silent' }

    foreach ($id in $App.Candidates) {
        try {
            $code = Start-WingetProcess -Arguments @('install','--id', $id, '-e') + $common -Log $Log
            if ($code -eq 0) { Write-Log $Log "SUCCESS: $($App.Display) via id '$id'"; return $true }
            else { Write-Log $Log "Attempt with id '$id' returned $code. Trying next candidate..." }
        } catch {
            Write-Log $Log "Attempt with id '$id' failed: $($_.Exception.Message)"
        }
    }

    try {
        $code = Start-WingetProcess -Arguments @('install','--name',$App.Display,'-e') + $common -Log $Log
        if ($code -eq 0) { Write-Log $Log "SUCCESS: $($App.Display) via exact name"; return $true }
        else { Write-Log $Log "FAILED: $($App.Display) via exact name. Exit code $code" }
    } catch {
        Write-Log $Log "FAILED: $($App.Display) via exact name. $($_.Exception.Message)"
    }
    return $false
}

function Update-All {
    param([switch]$Silent, [System.Windows.Controls.TextBox]$Log)
    $args = @('upgrade','--all','--include-unknown','--accept-source-agreements','--accept-package-agreements')
    if ($Silent) { $args += '--silent' }
    $code = Start-WingetProcess -Arguments $args -Log $Log
    if ($code -eq 0) { Write-Log $Log "All upgradable packages updated." }
    else { Write-Log $Log "winget upgrade --all exited with code $code." }
}

# --- Prep data for WPF binding ---
# Add IsSelected/Status properties so the UI can bind to them.
$AppItems = New-Object System.Collections.ObjectModel.ObservableCollection[object]
foreach ($a in $Apps) {
    $AppItems.Add([pscustomobject]@{
        Display    = $a.Display
        Candidates = $a.Candidates
        IsSelected = $false
        Status     = ''
    })
}

# --- XAML UI ---
Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase

$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Winget Installer" Height="720" Width="980" WindowStartupLocation="CenterScreen">
  <Window.Resources>
    <Style TargetType="Button">
      <Setter Property="Margin" Value="6"/>
      <Setter Property="Padding" Value="10,6"/>
      <Setter Property="Height" Value="28"/>
    </Style>
    <Style TargetType="TextBox">
      <Setter Property="Margin" Value="6"/>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="Margin" Value="6"/>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="3*"/>
      <ColumnDefinition Width="2*"/>
    </Grid.ColumnDefinitions>

    <!-- Header -->
    <Border Grid.Row="0" Grid.ColumnSpan="2" Background="#F5F5F5" Padding="10">
      <StackPanel>
        <TextBlock Text="Winget App Installer" FontSize="18" FontWeight="SemiBold"/>
        <TextBlock Text="Select apps, install them, or update everything at once." Foreground="#666"/>
      </StackPanel>
    </Border>

    <!-- LEFT -->
    <DockPanel Grid.Row="1" Grid.Column="0" LastChildFill="True">
      <StackPanel DockPanel.Dock="Top" Orientation="Horizontal">
        <TextBlock Text="Search:" VerticalAlignment="Center" Margin="6,6,0,6"/>
        <TextBox x:Name="SearchBox" Width="260"/>
        <CheckBox x:Name="SilentCheck" Content="Silent mode (where supported)" IsChecked="True" VerticalAlignment="Center"/>
      </StackPanel>

      <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal" HorizontalAlignment="Left" Margin="6">
        <Button x:Name="BtnSelectAll" Content="Select All" Width="110"/>
        <Button x:Name="BtnClear" Content="Deselect All" Width="110"/>
      </StackPanel>

      <ListView x:Name="AppList" Margin="6" SelectionMode="Extended" MinHeight="200">
        <ListView.ItemTemplate>
          <DataTemplate>
            <DockPanel LastChildFill="True">
              <CheckBox IsChecked="{Binding IsSelected, Mode=TwoWay}" VerticalAlignment="Center" Margin="0,0,8,0"/>
              <TextBlock Text="{Binding Display}" VerticalAlignment="Center"/>
            </DockPanel>
          </DataTemplate>
        </ListView.ItemTemplate>
      </ListView>
    </DockPanel>

    <GridSplitter Grid.Row="1" Grid.Column="0" HorizontalAlignment="Right" Width="5" Background="#DDD" ShowsPreview="True"/>

    <!-- RIGHT -->
    <Border Grid.Row="1" Grid.Column="1" BorderBrush="#DDD" BorderThickness="1" Margin="6">
      <TextBox x:Name="LogBox" FontFamily="Consolas" FontSize="12" IsReadOnly="True"
               VerticalScrollBarVisibility="Auto" TextWrapping="NoWrap" AcceptsReturn="True" BorderThickness="0"/>
    </Border>

    <!-- Bottom bar with real progress bar -->
    <Grid Grid.Row="2" Grid.ColumnSpan="2" Margin="0,4,0,6">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      <TextBlock x:Name="StatusText" Grid.Column="0" Margin="10,6" Foreground="#555" Text="Ready"/>
      <ProgressBar x:Name="OverallBar" Grid.Column="1" Height="16" Margin="10,6"
                   Minimum="0" Maximum="100" Value="0"/>
      <Button x:Name="BtnInstall"   Grid.Column="2" Content="Install Selected" Width="160"/>
      <Button x:Name="BtnUpdateAll" Grid.Column="3" Content="Update All" Width="160" Margin="6"/>
    </Grid>
  </Grid>
</Window>
"@



# Load XAML
$reader = (New-Object System.Xml.XmlNodeReader ([xml]$Xaml))
$window = [Windows.Markup.XamlReader]::Load($reader)

# Grab controls
$SearchBox   = $window.FindName('SearchBox')
$SilentCheck = $window.FindName('SilentCheck')
$AppList     = $window.FindName('AppList')
$LogBox      = $window.FindName('LogBox')
$BtnSelectAll= $window.FindName('BtnSelectAll')
$BtnClear    = $window.FindName('BtnClear')
$BtnInstall  = $window.FindName('BtnInstall')
$BtnUpdateAll= $window.FindName('BtnUpdateAll')
$StatusText  = $window.FindName('StatusText')

$OverallBar  = $window.FindName('OverallBar')
$OverallBar.Value = 0


# Bind items + sorting
$AppList.ItemsSource = $AppItems
$view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($AppItems)
$view.SortDescriptions.Clear()
$null = $view.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription "Display", "Ascending"))

# Filtering
$FilterHandler = {
    $view.Refresh()
}
$null = $SearchBox.Add_TextChanged($FilterHandler)

$view.Filter = {
    param($item)
    $term = $SearchBox.Text
    if ([string]::IsNullOrWhiteSpace($term)) { return $true }
    return $item.Display -like "*$term*"
}

# Busy toggle
function Set-UiBusy([bool]$Busy) {
    if ($Busy) {
        $window.Cursor = 'Wait'
        $StatusText.Text = "Working..."
    } else {
        $window.Cursor = 'Arrow'
        $StatusText.Text = "Ready"
    }
    $BtnInstall.IsEnabled   = -not $Busy
    $BtnUpdateAll.IsEnabled = -not $Busy
    $BtnSelectAll.IsEnabled = -not $Busy
    $BtnClear.IsEnabled     = -not $Busy
    $AppList.IsEnabled      = -not $Busy
    $SilentCheck.IsEnabled  = -not $Busy
}

# Wire buttons
$BtnSelectAll.Add_Click({
    foreach ($i in $AppItems) { $i.IsSelected = $true }
    $AppList.Items.Refresh()
})
$BtnClear.Add_Click({
    foreach ($i in $AppItems) { $i.IsSelected = $false }
    $AppList.Items.Refresh()
})

$BtnInstall.Add_Click({
    try {
        Test-Winget
        $selected = @($AppItems | Where-Object { $_.IsSelected })
        if (-not $selected -or $selected.Count -eq 0) {
            [System.Windows.MessageBox]::Show("Please select at least one app.","Nothing selected","OK","Information") | Out-Null
            return
        }
        Set-UiBusy $true
        $total = $selected.Count
        $OverallBar.Minimum = 0
        $OverallBar.Maximum = $total
        $OverallBar.IsIndeterminate = $false
        $OverallBar.Value = 0
        Write-Log $LogBox "===== INSTALL SELECTED ====="
        for ($i=0; $i -lt $total; $i++) {
            $app = $selected[$i]
            $StatusText.Text = "Installing $($i+1)/$($total): $($app.Display)"
            Write-Log $LogBox "Installing: $($app.Display)"
            [void](Install-App -App $app -Silent:([bool]$SilentCheck.IsChecked) -Log $LogBox)
            $OverallBar.Value = $i + 1
            Write-Log $LogBox "----------------------------------------"
        }
        $StatusText.Text = "Done"
        Write-Log $LogBox "All selected installs attempted."
    } catch {
        Write-Log $LogBox "Unexpected error: $($_.Exception.Message)"
    } finally {
        Set-UiBusy $false
        $StatusText.Text = "Ready"
    }
})

$BtnUpdateAll.Add_Click({
    try {
        Test-Winget
        Set-UiBusy $true
        $OverallBar.IsIndeterminate = $true
        Write-Log $LogBox "===== UPDATE ALL PACKAGES ====="
        Update-All -Silent:([bool]$SilentCheck.IsChecked) -Log $LogBox
    } catch {
        Write-Log $LogBox "Unexpected error: $($_.Exception.Message)"
    } finally {
        $OverallBar.IsIndeterminate = $false
        $OverallBar.Value = 0
        Set-UiBusy $false
        $StatusText.Text = "Ready"
    }
})

# Show window
[void]$window.ShowDialog()
