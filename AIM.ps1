# Load the required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "App Installation Monitor (AIM) 2.0"
$form.Size = New-Object System.Drawing.Size(700, 540)
$form.StartPosition = "CenterScreen"

# Set a larger font for the form
$formFont = New-Object System.Drawing.Font("Segoe UI", 14)
$form.Font = $formFont

# Create a menu bar
$menu = New-Object System.Windows.Forms.MenuStrip
$aboutMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Help")
$aboutItem = New-Object System.Windows.Forms.ToolStripMenuItem("About")
$aboutMenu.DropDownItems.Add($aboutItem)
$menu.Items.Add($aboutMenu)
$form.Controls.Add($menu)

# Add event handler for About menu item
$aboutItem.Add_Click({
    [System.Windows.Forms.MessageBox]::Show(
        "Application Installation Monitor (AIM)`nVersion number: 2.0`nWritten by: Todd Loenhorst`nGit page: https://github.com/burnoil", 
        "About AIM", 
        [System.Windows.Forms.MessageBoxButtons]::OK, 
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
})

# Create a label for the header
$headerLabel = New-Object System.Windows.Forms.Label
$headerLabel.Text = "Monitoring the following apps:"
$headerLabel.Size = New-Object System.Drawing.Size(680, 30)
$headerLabel.Location = New-Object System.Drawing.Point(10, 40)  # Adjusted for the menu
$headerLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$headerLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($headerLabel)

# Placeholder for system information
$systemInfoLabel = New-Object System.Windows.Forms.Label
$systemInfoLabel.Text = "Loading system information..."
$systemInfoLabel.Size = New-Object System.Drawing.Size(680, 110)
$systemInfoLabel.Location = New-Object System.Drawing.Point(10, 80)
$systemInfoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$systemInfoLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($systemInfoLabel)

# Create a listview to display the apps and their status
$listView = New-Object System.Windows.Forms.ListView
$listView.Size = New-Object System.Drawing.Size(680, 300)
$listView.Location = New-Object System.Drawing.Point(10, 190)
$listView.View = 'Details'
$listView.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$listView.Columns.Add("Application", 300) | Out-Null
$listView.Columns.Add("Version", 180) | Out-Null
$listView.Columns.Add("Status", 150) | Out-Null
$listView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($listView)

# Define bold fonts for ListView items
$boldFontGreen = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$boldFontRed = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)

# Apps to monitor
$appsToMonitor = @(
    "*Microsoft Teams*",
    "*Adobe Acrobat*",
    "*7-Zip*",
    "*Alertus Desktop*",
    "*BigFix Client*",
    "*Cisco Secure Client*",
    "*FireEye*",
    "*Microsoft 365*"
)

# Function to check installed applications
function Get-InstalledApps {
    $installedApps = @()
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    foreach ($regPath in $regPaths) {
        $keys = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
        foreach ($key in $keys) {
            if ($key.DisplayName) {
                $installedApps += [PSCustomObject]@{
                    Name = $key.DisplayName
                    Version = if ($key.DisplayVersion) { $key.DisplayVersion } else { "Unknown" }
                }
            }
        }
    }
    return $installedApps
}

# Function to update the ListView with app statuses
function Update-AppStatus {
    $listView.Items.Clear()
    $installedApps = Get-InstalledApps
    foreach ($appPattern in $appsToMonitor) {
        $matched = $false
        foreach ($app in $installedApps) {
            if ($app.Name -like $appPattern) {
                $listViewItem = $listView.Items.Add($app.Name)
                $listViewItem.SubItems.Add($app.Version) | Out-Null
                $listViewItem.SubItems.Add("Installed") | Out-Null
                $listViewItem.ForeColor = [System.Drawing.Color]::Green
                $listViewItem.Font = $boldFontGreen
                $matched = $true
                break
            }
        }
        if (-not $matched) {
            $listViewItem = $listView.Items.Add($appPattern)
            $listViewItem.SubItems.Add("N/A") | Out-Null
            $listViewItem.SubItems.Add("Not Installed") | Out-Null
            $listViewItem.ForeColor = [System.Drawing.Color]::Red
            $listViewItem.Font = $boldFontRed
        }
    }
}

# Timer to refresh the list every 10 seconds
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 10000  # 10 seconds
$timer.Add_Tick({ Update-AppStatus })
$timer.Start()

# Function to load system information incrementally
function Load-SystemInfo {
    # Load basic info first
    $systemInfoLabel.Text = "Machine Name: Loading...`nModel: Loading...`nRAM: Loading...`nFree Disk Space: Loading...`nWindows Version: Loading..."

    $computerName = (Get-WmiObject Win32_ComputerSystem).Name
    $systemInfoLabel.Text = "Machine Name: $computerName`nModel: Loading...`nRAM: Loading...`nFree Disk Space: Loading...`nWindows Version: Loading..."

    $computerModel = (Get-WmiObject Win32_ComputerSystem).Model
    $systemInfoLabel.Text = "Machine Name: $computerName`nModel: $computerModel`nRAM: Loading...`nFree Disk Space: Loading...`nWindows Version: Loading..."

    $totalRAM = [math]::round((Get-WmiObject Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2)
    $systemInfoLabel.Text = "Machine Name: $computerName`nModel: $computerModel`nRAM: $totalRAM GB`nFree Disk Space: Loading...`nWindows Version: Loading..."

    $freeDiskSpace = [math]::round((Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'").FreeSpace / 1GB, 2)
    $systemInfoLabel.Text = "Machine Name: $computerName`nModel: $computerModel`nRAM: $totalRAM GB`nFree Disk Space: $freeDiskSpace GB`nWindows Version: Loading..."

    $windowsProductName = (Get-WmiObject Win32_OperatingSystem).Caption
    $osDisplayVersion = Get-ItemPropertyValue "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name "DisplayVersion"

    # Final update with all information
    $systemInfoLabel.Text = "Machine Name: $computerName`nModel: $computerModel`nRAM: $totalRAM GB`nFree Disk Space: $freeDiskSpace GB`nWindows Version: $windowsProductName $osDisplayVersion"
}

# Use the Load event to start loading system information incrementally
$form.Add_Load({
    Load-SystemInfo
})

# Show the form
$form.ShowDialog()
