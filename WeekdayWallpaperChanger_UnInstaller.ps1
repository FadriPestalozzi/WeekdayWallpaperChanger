<#
.SYNOPSIS
    Run to install or uninstall the Daily Wallpaper Changer for Windows 11 to set desktop background image depending on weekday. 

.DESCRIPTION
    - Installer: Copies and patches the set_weekday_wallpaper.ps1 for use as a weekday-dependent desktop background switcher,
      registering a scheduled task to run it daily at midnight.
    - Uninstaller: Removes the scheduled task and deletes the installed set_weekday_wallpaper.ps1.

.NOTES
    Date: 2025-06-14
    OS: Windows 11, PowerShell 5.1 or higher
    Authors (Centaur Project): 
    - Fadri Pestalozzi (human)
    - Claude-4-sonnet (AI, aka Ada) 
    - The name "Ada" is an hommage to Ada Lovelace, the grandmother of flexibly programmable machines
#>

# CRITICAL: param block MUST be at the very top for ps2exe compatibility
param(
    [switch]$Uninstall,
    [switch]$Install
)

# Load required assemblies and hide console IMMEDIATELY - before any other operations
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Hide console window FIRST - before any other operations that might show errors
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

# Hide console window immediately (0 = hide, 5 = show)
$consolePtr = [Console.Window]::GetConsoleWindow()
[void][Console.Window]::ShowWindow($consolePtr, 0)

# Enable high DPI awareness for better display quality
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class DpiAware {
    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();
}
'@
# Properly suppress the return value (True/False) from SetProcessDPIAware
[void][DpiAware]::SetProcessDPIAware()

# Set application to use visual styles for better appearance
[System.Windows.Forms.Application]::EnableVisualStyles()

# Check if running as administrator and prompt to restart with admin rights
function Test-Administrator {
    try {
        $currentUser = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
        return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        return $false
    }
}

function Request-AdminRights {
    if (-not (Test-Administrator)) {
        # Create high-quality custom dialog for admin rights prompt
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Administrator Rights Required"
        $form.Size = New-Object System.Drawing.Size(600, 220)
        $form.StartPosition = "CenterScreen"
        $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $form.MaximizeBox = $false
        $form.MinimizeBox = $false
        $form.TopMost = $true
        
        # Use Segoe UI font for better appearance
        $font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
        $form.Font = $font
        
        # Main message
        $label = New-Object System.Windows.Forms.Label
        $label.Text = "This application requires administrator privileges to create scheduled tasks."
        $label.Location = New-Object System.Drawing.Point(20, 20)
        $label.Size = New-Object System.Drawing.Size(550, 40)
        $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $label.Font = $font
        
        # Question
        $questionLabel = New-Object System.Windows.Forms.Label
        $questionLabel.Text = "Would you like to restart with administrator rights?"
        $questionLabel.Location = New-Object System.Drawing.Point(20, 80)
        $questionLabel.Size = New-Object System.Drawing.Size(550, 30)
        $questionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $questionLabel.Font = $font
        
        # Yes button
        $yesBtn = New-Object System.Windows.Forms.Button
        $yesBtn.Text = "Yes"
        $yesBtn.Font = $font
        $yesBtn.Size = New-Object System.Drawing.Size(100, 35)
        $yesBtn.Location = New-Object System.Drawing.Point(200, 130)
        $yesBtn.DialogResult = [System.Windows.Forms.DialogResult]::Yes
        $yesBtn.BackColor = [System.Drawing.Color]::LightGreen
        $yesBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::System
        
        # No button
        $noBtn = New-Object System.Windows.Forms.Button
        $noBtn.Text = "No"
        $noBtn.Font = $font
        $noBtn.Size = New-Object System.Drawing.Size(100, 35)
        $noBtn.Location = New-Object System.Drawing.Point(320, 130)
        $noBtn.DialogResult = [System.Windows.Forms.DialogResult]::No
        $noBtn.BackColor = [System.Drawing.Color]::LightCoral
        $noBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::System
        
        # Add controls to form
        $form.Controls.AddRange(@($label, $questionLabel, $yesBtn, $noBtn))
        $form.AcceptButton = $yesBtn
        $form.CancelButton = $noBtn
        
        # Show dialog
        $result = $form.ShowDialog()
        
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                # Get the current process and determine the executable path
                $currentProcess = Get-Process -Id $PID
                $exePath = $null
                
                # Try multiple methods to get the executable path
                try {
                    $exePath = $currentProcess.MainModule.FileName
                } catch {
                    # Fallback method
                    $exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
                }
                
                # If still no path, try PowerShell method
                if (-not $exePath -or -not (Test-Path $exePath)) {
                    $exePath = $MyInvocation.MyCommand.Path
                    if (-not $exePath) {
                        $exePath = $PSCommandPath
                    }
                }
                
                # Determine if we're running as .exe
                $isExe = $exePath -and ($exePath -like "*.exe")
                
                # Prepare arguments
                $arguments = @()
                if ($Uninstall) {
                    $arguments += "-Uninstall"
                } else {
                    # If we're in Install-WallpaperTask, we know user chose Install
                    $arguments += "-Install"
                }
                
                if ($isExe -and (Test-Path $exePath)) {
                    # We're running as .exe - restart the .exe with admin rights
                    if ($arguments.Count -gt 0) {
                        Start-Process -FilePath $exePath -ArgumentList $arguments -Verb RunAs -Wait:$false
                    } else {
                        Start-Process -FilePath $exePath -Verb RunAs -Wait:$false
                    }
                } else {
                    # We're running as .ps1 - use PowerShell to restart
                    $scriptPath = $exePath
                    if (-not $scriptPath) {
                        throw "Could not determine script path"
                    }
                    
                    $psArgs = @("-ExecutionPolicy", "Bypass", "-File", "`"$scriptPath`"")
                    if ($arguments.Count -gt 0) {
                        $psArgs += $arguments
                    }
                    
                    Start-Process "powershell.exe" -ArgumentList $psArgs -Verb RunAs -Wait:$false
                }
                
                # Give a small delay to ensure the new process starts
                Start-Sleep -Milliseconds 500
                
                # Exit this instance immediately
                [System.Environment]::Exit(0)
                
            } catch {
                # Create error dialog with more specific error information
                $errorForm = New-Object System.Windows.Forms.Form
                $errorForm.Text = "Error"
                $errorForm.Size = New-Object System.Drawing.Size(600, 250)  
                $errorForm.StartPosition = "CenterScreen"
                $errorForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
                $errorForm.MaximizeBox = $false
                $errorForm.MinimizeBox = $false
                $errorForm.TopMost = $true
                
                $errorLabel = New-Object System.Windows.Forms.Label
                $errorLabel.Text = "Failed to restart with administrator rights.`n`nError: $($_.Exception.Message)`n`nPlease right-click the application and select 'Run as administrator' manually."
                $errorLabel.Location = New-Object System.Drawing.Point(20, 20)
                $errorLabel.Size = New-Object System.Drawing.Size(560, 140)  
                $errorLabel.TextAlign = [System.Drawing.ContentAlignment]::TopCenter
                $errorLabel.Font = $font
                
                $okBtn = New-Object System.Windows.Forms.Button
                $okBtn.Text = "OK"
                $okBtn.Font = $font
                $okBtn.Size = New-Object System.Drawing.Size(75, 30)
                $okBtn.Location = New-Object System.Drawing.Point(262, 180)  
                $okBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
                
                $errorForm.Controls.AddRange(@($errorLabel, $okBtn))
                $errorForm.AcceptButton = $okBtn
                $errorForm.CancelButton = $okBtn  # Enable ESC key
                $errorForm.ShowDialog()
                
                exit 1
            }
        } else {
            # Create warning dialog
            $warningForm = New-Object System.Windows.Forms.Form
            $warningForm.Text = "Cannot Continue"
            $warningForm.Size = New-Object System.Drawing.Size(600, 200)  
            $warningForm.StartPosition = "CenterScreen"
            $warningForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
            $warningForm.MaximizeBox = $false
            $warningForm.MinimizeBox = $false
            $warningForm.TopMost = $true
            
            $warningLabel = New-Object System.Windows.Forms.Label
            $warningLabel.Text = "Administrator rights are required for this application to function properly."
            $warningLabel.Location = New-Object System.Drawing.Point(20, 20)
            $warningLabel.Size = New-Object System.Drawing.Size(560, 80)  
            $warningLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $warningLabel.Font = $font
            
            $okBtn = New-Object System.Windows.Forms.Button
            $okBtn.Text = "OK"
            $okBtn.Font = $font
            $okBtn.Size = New-Object System.Drawing.Size(75, 30)
            $okBtn.Location = New-Object System.Drawing.Point(262, 120)  
            $okBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
            
            $warningForm.Controls.AddRange(@($warningLabel, $okBtn))
            $warningForm.AcceptButton = $okBtn
            $warningForm.CancelButton = $okBtn  # Enable ESC key to close dialog
            $warningForm.ShowDialog()
            
            exit 1
        }
    }
}

# Request admin rights at the beginning
# Request-AdminRights

# === PARAMETRIC UI HELPER FUNCTIONS ===
function Get-DialogConfig($type) {
    # Define different dialog configurations
    switch ($type) {
        "FolderSelect" {
            return @{
                Width = 760
                Height = 350
                Margin = 15
                Spacing = 10
                DescHeight = 120
                LabelHeight = 30
                TextboxHeight = 30
                ButtonWidth = 75
                ButtonHeight = 30
                Font = "Segoe UI"
                FontSize = 10
            }
        }
        "XmlSelect" {
            return @{
                Width = 850
                Height = 380
                Margin = 10
                Spacing = 10
                DescHeight = 40
                LabelHeight = 20
                TextboxHeight = 25
                ListHeight = 140
                ButtonWidth = 75
                ButtonHeight = 30
                Font = "Segoe UI"
                FontSize = 10
            }
        }
        "AdminRights" {
            return @{
                Width = 600
                Height = 220
                Margin = 20
                Spacing = 15
                LabelHeight = 40
                QuestionHeight = 30
                ButtonWidth = 100
                ButtonHeight = 35
                Font = "Segoe UI"
                FontSize = 10
            }
        }
        "WallpaperStyle" {
            return @{
                Width = 800
                Height = 500
                Margin = 20
                Spacing = 15
                DescHeight = 40
                GroupBoxHeight = 320
                GroupBoxTitleHeight = 25
                RadioHeight = 50
                RadioSpacing = 55
                ButtonWidth = 80
                ButtonHeight = 35
                Font = "Segoe UI"
                FontSize = 10
            }
        }
        default {
            # Default configuration
            return @{
                Width = 600
                Height = 300
                Margin = 15
                Spacing = 12
                Font = "Segoe UI"
                FontSize = 10
            }
        }
    }
}

function Choose-Folder($desc, $showNew = $true) {
    # Get parametric configuration
    $config = Get-DialogConfig "FolderSelect"
    
    # Create custom dialog with path input option
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Folder"
    $form.Size = New-Object System.Drawing.Size($config.Width, $config.Height)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true
    
    $font = New-Object System.Drawing.Font($config.Font, $config.FontSize)

    # === PARAMETRIC LAYOUT CALCULATIONS ===
    # Calculate component dimensions
    $descriptionSizeWidth = $config.Width - (2 * $config.Margin)
    $textboxWidth = 600
    $browseButtonSizeWidth = 100
    
    # Calculate Y positions step by step (avoid complex expressions in Point constructor)
    $descY = $config.Margin
    $pathLabelY = $descY + $config.DescHeight + $config.Spacing
    $textboxY = $pathLabelY + $config.LabelHeight + $config.Spacing
    $buttonRowY = $textboxY  # Align buttons with textbox
    $bottomButtonY = $textboxY + $config.TextboxHeight + (2 * $config.Spacing)  # Bottom row for OK/Cancel
    
    # Calculate X positions step by step
    $textboxX = $config.Margin
    $browseX = $textboxX + $textboxWidth + $config.Spacing
    $okX = $textboxX  # Align OK button under textbox
    $cancelX = $okX + $config.ButtonWidth + $config.Spacing
    
    # Calculate button height (avoid arithmetic in Size constructor)
    $browseButtonHeight = $config.TextboxHeight + 2
    
    # Description label with text wrapping
    $lblDesc = New-Object System.Windows.Forms.Label
    $lblDesc.Text = $desc
    $lblDesc.Location = New-Object System.Drawing.Point($config.Margin, $descY)
    $lblDesc.Size = New-Object System.Drawing.Size($descriptionSizeWidth, $config.DescHeight)  
    $lblDesc.Font = $font
    $lblDesc.AutoSize = $false
    $lblDesc.TextAlign = [System.Drawing.ContentAlignment]::TopLeft
    
    # Path input label
    $lblPath = New-Object System.Windows.Forms.Label
    $lblPath.Text = "Enter path directly (or use Browse button):"
    $lblPath.Location = New-Object System.Drawing.Point($config.Margin, $pathLabelY)
    $lblPath.Size = New-Object System.Drawing.Size(400, $config.LabelHeight)  
    $lblPath.Font = $font
    
    # Path textbox
    $txtPath = New-Object System.Windows.Forms.TextBox
    $txtPath.Location = New-Object System.Drawing.Point($textboxX, $textboxY)
    $txtPath.Size = New-Object System.Drawing.Size($textboxWidth, $config.TextboxHeight)  
    $txtPath.Font = $font
    
    # Browse button (aligned with textbox)
    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = "Browse..."
    $btnBrowse.Location = New-Object System.Drawing.Point($browseX, $buttonRowY)
    $btnBrowse.Size = New-Object System.Drawing.Size($browseButtonSizeWidth, $browseButtonHeight)
    $btnBrowse.Font = $font
    
    # OK and Cancel buttons (bottom row)
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Location = New-Object System.Drawing.Point($okX, $bottomButtonY)
    $btnOK.Size = New-Object System.Drawing.Size($config.ButtonWidth, $config.ButtonHeight)
    $btnOK.Font = $font
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point($cancelX, $bottomButtonY)
    $btnCancel.Size = New-Object System.Drawing.Size($config.ButtonWidth, $config.ButtonHeight)
    $btnCancel.Font = $font
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    
    # Browse button click event
    $btnBrowse.Add_Click({
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = $desc
        $dialog.ShowNewFolderButton = $showNew
        if ($dialog.ShowDialog() -eq 'OK') {
            $txtPath.Text = $dialog.SelectedPath
        }
    })
    
    # Validation function
    $validatePath = {
        $path = $txtPath.Text.Trim()
        if ([string]::IsNullOrEmpty($path)) {
            $btnOK.Enabled = $false
        } elseif (Test-Path $path) {
            $btnOK.Enabled = $true
            $txtPath.BackColor = [System.Drawing.Color]::White
        } else {
            $btnOK.Enabled = $true  # Allow non-existing paths for creation
            $txtPath.BackColor = [System.Drawing.Color]::LightYellow
        }
    }
    
    # Add validation on text change
    $txtPath.Add_TextChanged($validatePath)
    
    # Add controls to form
    $form.Controls.AddRange(@($lblDesc, $lblPath, $txtPath, $btnBrowse, $btnOK, $btnCancel))
    $form.AcceptButton = $btnOK
    $form.CancelButton = $btnCancel
    
    # Initial validation
    & $validatePath
    
    # Show dialog
    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedPath = $txtPath.Text.Trim()
        if ([string]::IsNullOrEmpty($selectedPath)) {
            throw "No path specified"
        }
        return $selectedPath
    } else {
        throw "Cancelled"
    }
}

function Choose-WallpaperStyle {
    # Get parametric configuration
    $config = Get-DialogConfig "WallpaperStyle"
    
    # Create custom dialog for wallpaper style selection
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Choose Wallpaper Display Mode"
    $form.Size = New-Object System.Drawing.Size($config.Width, $config.Height)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true
    
    $font = New-Object System.Drawing.Font($config.Font, $config.FontSize)
    
    # === PARAMETRIC LAYOUT CALCULATIONS ===
    # Calculate component dimensions
    $descriptionWidth = $config.Width - (2 * $config.Margin)
    $groupBoxWidth = $config.Width - (2 * $config.Margin)
    $radioWidth = $groupBoxWidth - (2 * $config.Spacing)
    
    # Calculate Y positions step by step
    $descY = $config.Margin
    $groupBoxY = $descY + $config.DescHeight + $config.Spacing
    $buttonsY = $groupBoxY + $config.GroupBoxHeight + $config.Spacing
    
    # Calculate X positions for buttons
    $okX = $config.Width - (2 * $config.ButtonWidth) - $config.Spacing - $config.Margin
    $cancelX = $okX + $config.ButtonWidth + $config.Spacing
    
    # Description label
    $lblDesc = New-Object System.Windows.Forms.Label
    $lblDesc.Text = "Choose how your wallpaper images should be displayed on your monitors:"
    $lblDesc.Location = New-Object System.Drawing.Point($config.Margin, $descY)
    $lblDesc.Size = New-Object System.Drawing.Size($descriptionWidth, $config.DescHeight)
    $lblDesc.Font = $font
    $lblDesc.TextAlign = [System.Drawing.ContentAlignment]::TopLeft
    
    # Create radio buttons for each style
    $radioButtons = @()
    $styles = @(
        @{ Value = "2"; Name = "Stretch"; Description = "Stretch image to fill screen exactly (may distort image, no black bars)" },
        @{ Value = "10"; Name = "Fill"; Description = "Fill screen while maintaining aspect ratio (may crop image, no black bars)" },
        @{ Value = "6"; Name = "Fit"; Description = "Fit entire image on screen (maintains aspect ratio, may show black bars)" },
        @{ Value = "0"; Name = "Center"; Description = "Center image without scaling (original size, may show black bars)" },
        @{ Value = "22"; Name = "Span"; Description = "Span image across all monitors as one continuous image" }
    )
    
    # Group box for radio buttons
    $groupBox = New-Object System.Windows.Forms.GroupBox
    $groupBox.Text = "Display Options"
    $groupBox.Location = New-Object System.Drawing.Point($config.Margin, $groupBoxY)
    $groupBox.Size = New-Object System.Drawing.Size($groupBoxWidth, $config.GroupBoxHeight)
    $groupBox.Font = $font
    
    # Create radio buttons inside group box (start below the GroupBox title)
    [int]$radioY = $config.GroupBoxTitleHeight + $config.Spacing
    foreach ($style in $styles) {
        $radio = New-Object System.Windows.Forms.RadioButton
        $radio.Text = "$($style.Name) - $($style.Description)"
        $radio.Location = New-Object System.Drawing.Point($config.Spacing, $radioY)
        $radio.Size = New-Object System.Drawing.Size($radioWidth, $config.RadioHeight)
        $radio.Font = $font
        $radio.Tag = $style.Value
        $radio.AutoSize = $false
        
        # Set Stretch as default (most common for multi-monitor)
        if ($style.Value -eq "2") {
            $radio.Checked = $true
        }
        
        $radioButtons += $radio
        $groupBox.Controls.Add($radio)
        $radioY = $radioY + $config.RadioSpacing
    }
    
    # OK and Cancel buttons
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Location = New-Object System.Drawing.Point($okX, $buttonsY)
    $btnOK.Size = New-Object System.Drawing.Size($config.ButtonWidth, $config.ButtonHeight)
    $btnOK.Font = $font
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point($cancelX, $buttonsY)
    $btnCancel.Size = New-Object System.Drawing.Size($config.ButtonWidth, $config.ButtonHeight)
    $btnCancel.Font = $font
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    
    # Add controls to form
    $form.Controls.AddRange(@($lblDesc, $groupBox, $btnOK, $btnCancel))
    $form.AcceptButton = $btnOK
    $form.CancelButton = $btnCancel
    
    # Show dialog
    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedStyle = $radioButtons | Where-Object { $_.Checked } | Select-Object -First 1
        if ($selectedStyle) {
            return $selectedStyle.Tag
        } else {
            return "2"  # Default to Stretch if nothing selected
        }
    } else {
        throw "Cancelled"
    }
}

function Choose-XmlFolder($desc) {
    # Create custom dialog for XML folder selection with auto-detection
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select XML Folder"
    $form.Size = New-Object System.Drawing.Size(720, 410)  
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true
    
    $font = New-Object System.Drawing.Font("Segoe UI", 10)
    
    # Description label
    $lblDesc = New-Object System.Windows.Forms.Label
    $lblDesc.Text = $desc
    $lblDesc.Location = New-Object System.Drawing.Point(10, 10)
    $lblDesc.Size = New-Object System.Drawing.Size(690, 90)  
    $lblDesc.Font = $font
    $lblDesc.AutoSize = $false
    
    # Folder path input
    $lblPath = New-Object System.Windows.Forms.Label
    $lblPath.Text = "Enter folder path (or use Browse button):"
    $lblPath.Location = New-Object System.Drawing.Point(10, 110)  
    $lblPath.Size = New-Object System.Drawing.Size(600, 30)  
    $lblPath.Font = $font
    
    $txtPath = New-Object System.Windows.Forms.TextBox
    $txtPath.Location = New-Object System.Drawing.Point(10, 150)  
    $txtPath.Size = New-Object System.Drawing.Size(600, 30)  
    $txtPath.Font = $font
    
    # Browse button
    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = "Browse..."
    $btnBrowse.Location = New-Object System.Drawing.Point(620, 150)  
    $btnBrowse.Size = New-Object System.Drawing.Size(120, 30)
    $btnBrowse.Font = $font
    
    # XML file selection - REMOVED (no longer needed)
    
    # OK and Cancel buttons
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Location = New-Object System.Drawing.Point(620, 285)  
    $btnOK.Size = New-Object System.Drawing.Size(80, 30)
    $btnOK.Font = $font
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnOK.Enabled = $false
    
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(620, 325)  
    $btnCancel.Size = New-Object System.Drawing.Size(80, 30)
    $btnCancel.Font = $font
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    
    # Function to update XML file list
    $updateXmlList = {
        $folderPath = $txtPath.Text.Trim()
        $listXml.Items.Clear()
        $btnOK.Enabled = $false
        
        if (-not [string]::IsNullOrEmpty($folderPath) -and (Test-Path $folderPath)) {
            $xmlFiles = Get-ChildItem -Path $folderPath -Filter "*.xml" -File | Sort-Object Name
            foreach ($file in $xmlFiles) {
                $listXml.Items.Add($file.Name)
            }
            if ($listXml.Items.Count -gt 0) {
                $listXml.SelectedIndex = 0  # Auto-select first XML file
                $btnOK.Enabled = $true
                $txtPath.BackColor = [System.Drawing.Color]::White
            }
        } else {
            $txtPath.BackColor = [System.Drawing.Color]::LightPink
        }
    }
    
    # Browse button click event
    $btnBrowse.Add_Click({
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = $desc
        $dialog.ShowNewFolderButton = $false
        if ($dialog.ShowDialog() -eq 'OK') {
            $txtPath.Text = $dialog.SelectedPath
            & $updateXmlList
        }
    })
    
    # Text change event
    $txtPath.Add_TextChanged($updateXmlList)
    
    # List selection change
    $listXml.Add_SelectedIndexChanged({
        $btnOK.Enabled = ($listXml.SelectedIndex -ge 0)
    })
    
    # Add controls
    $form.Controls.AddRange(@($lblDesc, $lblPath, $txtPath, $btnBrowse, $lblXml, $listXml, $btnOK, $btnCancel))
    $form.AcceptButton = $btnOK
    $form.CancelButton = $btnCancel
    
    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $folderPath = $txtPath.Text.Trim()
        $selectedXml = $listXml.SelectedItem
        if ([string]::IsNullOrEmpty($folderPath) -or [string]::IsNullOrEmpty($selectedXml)) {
            throw "No XML file selected"
        }
        return Join-Path $folderPath $selectedXml
    } else {
        throw "Cancelled"
    }
}

function Show-InstallDialog {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Weekday Wallpaper Changer - Installer/Uninstaller"
    $form.StartPosition = "CenterScreen"
    $form.Size = New-Object System.Drawing.Size(600, 230)  # Increased height for more space below buttons
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true
    
    # Use Segoe UI font for better appearance
    $font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Regular)
    $form.Font = $font

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Do you want to INSTALL or UNINSTALL the Weekday Wallpaper Changer?"
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $label.Size = New-Object System.Drawing.Size(560, 80)
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $label.Font = $font

    $installBtn = New-Object System.Windows.Forms.Button
    $installBtn.Text = "INSTALL"
    $installBtn.Font = $font
    $installBtn.AutoSize = $true
    $installBtn.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $installBtn.MinimumSize = New-Object System.Drawing.Size(100, 40)
    $installBtn.Padding = New-Object System.Windows.Forms.Padding(10, 5, 10, 5)
    $installBtn.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $installBtn.BackColor = [System.Drawing.Color]::LightGreen
    $installBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::System

    $uninstallBtn = New-Object System.Windows.Forms.Button
    $uninstallBtn.Text = "UNINSTALL"
    $uninstallBtn.Font = $font
    $uninstallBtn.AutoSize = $true
    $uninstallBtn.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $uninstallBtn.MinimumSize = New-Object System.Drawing.Size(100, 40)
    $uninstallBtn.Padding = New-Object System.Windows.Forms.Padding(10, 5, 10, 5)
    $uninstallBtn.DialogResult = [System.Windows.Forms.DialogResult]::No
    $uninstallBtn.BackColor = [System.Drawing.Color]::LightCoral
    $uninstallBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::System

    $cancelBtn = New-Object System.Windows.Forms.Button
    $cancelBtn.Text = "CANCEL"
    $cancelBtn.Font = $font
    $cancelBtn.AutoSize = $true
    $cancelBtn.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $cancelBtn.MinimumSize = New-Object System.Drawing.Size(100, 40)
    $cancelBtn.Padding = New-Object System.Windows.Forms.Padding(10, 5, 10, 5)
    $cancelBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $cancelBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::System

    # Add buttons to form first so they can auto-size
    $form.Controls.AddRange(@($label, $installBtn, $uninstallBtn, $cancelBtn))
    
    # Force the form to perform layout so buttons get their final sizes
    $form.PerformLayout()
    
    # Now position buttons dynamically based on their actual sizes
    $buttonSpacing = 15
    $totalButtonWidth = $installBtn.Width + $uninstallBtn.Width + $cancelBtn.Width + (2 * $buttonSpacing)
    $startX = ($form.ClientSize.Width - $totalButtonWidth) / 2
    
    $installBtn.Location = New-Object System.Drawing.Point($startX, 115)
    $uninstallBtn.Location = New-Object System.Drawing.Point(($startX + $installBtn.Width + $buttonSpacing), 115)
    $cancelBtn.Location = New-Object System.Drawing.Point(($startX + $installBtn.Width + $uninstallBtn.Width + (2 * $buttonSpacing)), 115)
    $form.AcceptButton = $installBtn
    $form.CancelButton = $cancelBtn

    return $form.ShowDialog()
}

function Uninstall-WallpaperTask {
    try {
        # Remove scheduled task
        schtasks.exe /Delete /TN "WeekdayWallpaperChanger" /F | Out-Null

        # Ask where the script is
        $scriptDir = Choose-Folder "Select the folder where the set_weekday_wallpaper.ps1 script was previously installed. This will remove the script file and clean up the installation." $false
        $scriptDest = Join-Path $scriptDir 'set_weekday_wallpaper.ps1'
        if (Test-Path $scriptDest) {
            Remove-Item $scriptDest -Force
            Write-Host "Deleted $scriptDest"
        } else {
            Write-Host "Script not found at $scriptDest (maybe already removed)."
        }

        Write-Host "`nUninstallation complete!"
        Write-Host "Your images remain untouched. You may delete them manually if desired."
    }
    catch {
        # Smart error handling - only show meaningful errors
        $errorMsg = $_.Exception.Message
        
        # Only show error popup for non-cancellation errors, remove console output to avoid duplicate messages
        if ($errorMsg -ne "Cancelled" -and $errorMsg -ne "No path specified") {
            [System.Windows.Forms.MessageBox]::Show("Uninstallation failed:`n$errorMsg", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
        # For cancellation, just silently return without popup
    }
}

function Install-WallpaperTask {
    # Check for admin rights before starting installation process
    if (-not (Test-Administrator)) {
        Request-AdminRights
        return  # Exit this function as the app will restart with admin rights
    }
    
    try {
        # 1. Prompt for images folder
        $imgFolder = Choose-Folder @"
Select the folder containing your 7 weekday wallpapers.
These should be named like: 1-Sun, 2-Mon, 3-Tue, 4-Wed, 5-Thu, 6-Fri, 7-Sat
You can use any supported image format: JPG, JPEG, PNG, BMP, GIF, TIFF, WEBP
Examples: 1-Sun.jpg, 2-Mon.png, 3-Tue.webp, etc.
"@ $true

        # 2. Prompt for script installation location with explanation
        $scriptDir = Choose-Folder "Select where to install the set_weekday_wallpaper.ps1 script. This script will automatically change your wallpaper based on the day of the week. Choose a permanent location like C:\Scripts or D:\Programs where the script can remain installed." $true

        # 3. Prompt for wallpaper display style
        $wallpaperStyle = Choose-WallpaperStyle

# --------------------- script start ---------------------

        # 4. Compose PowerShell script content 
        $wallpaperScript = @'
        
# Enhanced wallpaper setter that handles Windows 11 slideshow conflicts and multiple image formats
param(
    [string]$ImagePath = ""
)

# If no path provided, use today's image
if ([string]::IsNullOrEmpty($ImagePath)) {
    $dayIndex = [int](Get-Date).DayOfWeek
    $dayMap = @("1-Sun", "2-Mon", "3-Tue", "4-Wed", "5-Thu", "6-Fri", "7-Sat")
    $baseFilename = $dayMap[$dayIndex]
    $imgFolder = "PLACEHOLDER_IMG_FOLDER"
    
    Write-Host "Looking for wallpaper for day: $($dayMap[$dayIndex]) in folder: $imgFolder"
    
    # Search for image file with any supported extension (case-insensitive)
    $supportedExtensions = @("jpg", "jpeg", "png", "bmp", "gif", "tiff", "tif", "webp", "JPG", "JPEG", "PNG", "BMP", "GIF", "TIFF", "TIF", "WEBP")
    $ImagePath = $null
    
    foreach ($ext in $supportedExtensions) {
        $testPath = Join-Path $imgFolder "$baseFilename.$ext"
        Write-Host "Checking: $testPath"
        if (Test-Path $testPath) {
            $ImagePath = $testPath
            Write-Host "Found image: $ImagePath"
            break
        }
    }
    
    # If no exact match found, try fuzzy search (in case of naming variations)
    if (-not $ImagePath) {
        Write-Host "No exact match found, trying fuzzy search..."
        $searchPattern = "$baseFilename*"
        $foundFiles = Get-ChildItem -Path $imgFolder -Filter $searchPattern -File 2>$null
        if ($foundFiles) {
            # Filter to only image extensions
            $imageFiles = $foundFiles | Where-Object { 
                $_.Extension.ToLower() -in @(".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif", ".webp") 
            }
            if ($imageFiles) {
                $ImagePath = $imageFiles[0].FullName
                Write-Host "Found via fuzzy search: $ImagePath"
            }
        }
    }
    
    # If still no file found, default to .jpg for error message
    if (-not $ImagePath) {
        $ImagePath = Join-Path $imgFolder "$baseFilename.jpg"
        Write-Host "No image found, defaulting to: $ImagePath"
    }
}

Write-Host "Setting wallpaper to: $ImagePath"

# Verify image exists
if (-not (Test-Path $ImagePath)) {
    Write-Host "ERROR: Image not found: $ImagePath"
    Write-Host "Available files in directory:"
    try {
        Get-ChildItem -Path (Split-Path $ImagePath -Parent) -File | ForEach-Object { Write-Host "  $($_.Name)" }
    } catch {
        Write-Host "  Could not list directory contents"
    }
    exit 1
}

# Get absolute path to avoid any path resolution issues
$ImagePath = (Get-Item $ImagePath).FullName
Write-Host "Using absolute path: $ImagePath"

# Method 1: Enhanced Registry approach with better error handling
Write-Host "Setting registry values..."
try {
    # Set wallpaper style first
    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "WallpaperStyle" -Value "PLACEHOLDER_WALLPAPER_STYLE" -Force
    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "TileWallpaper" -Value "0" -Force
    
    # Clear any existing wallpaper first
    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "Wallpaper" -Value "" -Force
    Start-Sleep -Milliseconds 200
    
    # Set new wallpaper
    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "Wallpaper" -Value $ImagePath -Force
    Write-Host "Registry values set successfully"
} catch {
    Write-Host "ERROR setting registry values: $($_.Exception.Message)"
}

# Method 2: Disable slideshow and set picture mode more thoroughly
Write-Host "Configuring Windows 11 personalization settings..."
try {
    # Disable slideshow in multiple locations
    $personalPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
    if (-not (Test-Path $personalPath)) {
        New-Item -Path $personalPath -Force | Out-Null
    }
    
    # Windows 11 specific settings
    $desktopPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Wallpapers"
    if (-not (Test-Path $desktopPath)) {
        New-Item -Path $desktopPath -Force | Out-Null
    }
    Set-ItemProperty -Path $desktopPath -Name "BackgroundType" -Value 0 -Force  # 0 = Picture, 1 = Slideshow
    
    # Additional slideshow disable
    $lockPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Lock Screen\Creative"
    if (Test-Path $lockPath) {
        Set-ItemProperty -Path $lockPath -Name "CreativeId" -Value "" -Force -ErrorAction SilentlyContinue
    }
    
    Write-Host "Personalization settings configured"
} catch {
    Write-Host "Warning: Could not fully configure personalization settings: $($_.Exception.Message)"
}

# Method 3: Enhanced SystemParametersInfo with better constants
Write-Host "Applying via SystemParametersInfo..."
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class WallpaperAPI {
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
    
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool InvalidateRect(IntPtr hWnd, IntPtr lpRect, bool bErase);
    
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool UpdateWindow(IntPtr hWnd);
    
    public const int SPI_SETDESKWALLPAPER = 0x0014;
    public const int SPIF_UPDATEINIFILE = 0x01;
    public const int SPIF_SENDCHANGE = 0x02;
}
"@

try {
    # Use proper constants and force immediate update
    $result = [WallpaperAPI]::SystemParametersInfo([WallpaperAPI]::SPI_SETDESKWALLPAPER, 0, $ImagePath, [WallpaperAPI]::SPIF_UPDATEINIFILE -bor [WallpaperAPI]::SPIF_SENDCHANGE)
    Write-Host "SystemParametersInfo result: $result"
    
    if ($result -eq 0) {
        Write-Host "Warning: SystemParametersInfo returned 0 (may indicate failure)"
    }
} catch {
    Write-Host "ERROR with SystemParametersInfo: $($_.Exception.Message)"
}

# Method 4: Force desktop refresh
Write-Host "Forcing desktop refresh..."
try {
    # Refresh desktop
    [WallpaperAPI]::InvalidateRect([IntPtr]::Zero, [IntPtr]::Zero, $true)
    [WallpaperAPI]::UpdateWindow([IntPtr]::Zero)
    
    # Also try refreshing explorer
    $shell = New-Object -ComObject Shell.Application
    $shell.Windows() | ForEach-Object { 
        try { $_.Refresh() } catch { }
    }
} catch {
    Write-Host "Warning: Could not fully refresh desktop: $($_.Exception.Message)"
}

# Method 5: Additional notification methods
Write-Host "Broadcasting settings changes..."
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class SettingsNotify {
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern bool SendNotifyMessage(IntPtr hWnd, uint Msg, UIntPtr wParam, string lParam);
    
    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    
    public const uint WM_SETTINGCHANGE = 0x001A;
    public const uint WM_WININICHANGE = 0x001A;
    public const IntPtr HWND_BROADCAST = (IntPtr)0xFFFF;
}
"@

try {
    # Broadcast multiple setting change messages
    [SettingsNotify]::SendNotifyMessage([SettingsNotify]::HWND_BROADCAST, [SettingsNotify]::WM_SETTINGCHANGE, [UIntPtr]::Zero, "Environment")
    [SettingsNotify]::SendMessage([SettingsNotify]::HWND_BROADCAST, [SettingsNotify]::WM_SETTINGCHANGE, [IntPtr]::Zero, [IntPtr]::Zero)
    
    Write-Host "Settings change notifications sent"
} catch {
    Write-Host "Warning: Could not send all notifications: $($_.Exception.Message)"
}

# Give system time to process changes
Start-Sleep -Milliseconds 500

Write-Host "Wallpaper setting process complete!"
Write-Host ""
Write-Host "Troubleshooting info:"
Write-Host "- Image file: $ImagePath"
Write-Host "- File exists: $(Test-Path $ImagePath)"
Write-Host "- File size: $((Get-Item $ImagePath -ErrorAction SilentlyContinue).Length) bytes"
Write-Host ""
Write-Host "If wallpaper didn't change:"
Write-Host "1. Check Windows Settings > Personalization > Background"
Write-Host "2. Ensure it's set to 'Picture' (not 'Slideshow' or 'Solid color')"
Write-Host "3. Try manually selecting the image: $ImagePath"
Write-Host "4. Restart Windows Explorer: taskkill /f /im explorer.exe && explorer.exe"

'@

# --------------------- script end ---------------------

        # 5. Write the script to install location, patch placeholders
        $scriptDest = Join-Path $scriptDir 'set_weekday_wallpaper.ps1'
        $wallpaperScript -replace 'PLACEHOLDER_IMG_FOLDER', [regex]::Escape($imgFolder) -replace 'PLACEHOLDER_WALLPAPER_STYLE', $wallpaperStyle | Set-Content $scriptDest -Force

        # 6. Create advanced scheduled task using XML for better control
        $taskXml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>Daily wallpaper changer based on weekday</Description>
    <Author>$env:USERNAME</Author>
  </RegistrationInfo>
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>$(Get-Date -Format 'yyyy-MM-dd')T00:01:00</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByDay>
        <DaysInterval>1</DaysInterval>
      </ScheduleByDay>
    </CalendarTrigger>
    <LogonTrigger>
      <Enabled>true</Enabled>
      <UserId>$env:USERNAME</UserId>
    </LogonTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>$env:USERDOMAIN\$env:USERNAME</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT10M</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>powershell.exe</Command>
      <Arguments>-ExecutionPolicy Bypass -File "$scriptDest"</Arguments>
    </Exec>
  </Actions>
</Task>
"@

        # 7. Save XML to temp file and import it
        $tempXmlFile = Join-Path $env:TEMP "WeekdayWallpaperChanger.xml"
        $taskXml | Set-Content $tempXmlFile -Force -Encoding UTF8
        
        Write-Host "Creating advanced task with XML configuration..."
        $scheduleResult = Start-Process schtasks.exe -ArgumentList "/Create", "/TN", "WeekdayWallpaperChanger", "/XML", "`"$tempXmlFile`"", "/F" -PassThru -Wait -Verb RunAs
        
        # Clean up temp file
        if (Test-Path $tempXmlFile) {
            Remove-Item $tempXmlFile -Force
        }
        
        # Check if Task Scheduler creation was successful
        if ($scheduleResult.ExitCode -eq 0) {
            # Verify the task was created
            $taskExists = (schtasks.exe /Query /TN "WeekdayWallpaperChanger" 2>$null) -ne $null
            if ($taskExists) {
                Write-Host "✅ Task Scheduler entry created successfully!"
                Write-Host "Task Name: WeekdayWallpaperChanger"
                Write-Host ""
                Write-Host "🔧 Manual Testing Options:"
                Write-Host "1. Open Task Scheduler (taskschd.msc) and look for 'WeekdayWallpaperChanger'"
                Write-Host "2. Right-click the task and select 'Run' to test immediately"
                Write-Host "3. Or run this command to trigger manually:"
                Write-Host "   schtasks.exe /Run /TN `"WeekdayWallpaperChanger`""
                Write-Host ""
                Write-Host "📄 Script Location: $scriptDest"
                Write-Host "🖼️ Images Folder: $imgFolder"
                
                # Display selected wallpaper style
                $styleNames = @{
                    "0" = "Center"
                    "2" = "Stretch" 
                    "6" = "Fit"
                    "10" = "Fill"
                    "22" = "Span"
                }
                $styleName = $styleNames[$wallpaperStyle]
                Write-Host "🎨 Wallpaper Style: $styleName (value: $wallpaperStyle)"
                
                Write-Host "✅ Installation complete! Your wallpaper will change automatically daily."
            } else {
                throw "Task was created but cannot be found in Task Scheduler"
            }
        } else {
            throw "Failed to create scheduled task. Please check if you have administrator privileges (Exit code: $($scheduleResult.ExitCode))"
        }
    }
    catch {
        # Smart error handling - only show meaningful errors, remove console output to avoid duplicate messages
        $errorMsg = $_.Exception.Message
        
        # Only show error popup for non-cancellation errors
        if ($errorMsg -ne "Cancelled" -and $errorMsg -ne "No path specified" -and $errorMsg -ne "Invalid file path") {
            [System.Windows.Forms.MessageBox]::Show("Installation failed:`n$errorMsg", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
        # For cancellation or path errors, just silently return without popup
    }
}

# Main execution logic
if ($Install) {
    Install-WallpaperTask
} elseif ($Uninstall) {
    Uninstall-WallpaperTask
} else {
    # Interactive prompt when no parameters provided
    $result = Show-InstallDialog
    switch ($result) {
        'Yes' { Install-WallpaperTask }
        'No' { Uninstall-WallpaperTask }
        default { 
            Write-Host "Cancelled by user."
            exit 0
        }
    }
}
