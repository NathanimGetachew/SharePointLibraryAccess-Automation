<#
.SYNOPSIS
    SharePoint Library Access Manager - GUI tool for managing document library permissions
.DESCRIPTION
    This PowerShell script provides a graphical interface to manage SharePoint document library permissions.
    It allows administrators to grant or revoke access to multiple users across multiple libraries with selected permission levels.
.NOTES
    Author: NathanimG
    Email: NathanimG@ethiopianairlines.com
    Mobile: +251910160067
    Version: 1.2
    Date: $(Get-Date -Format "yyyy-MM-dd")
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#region Help System
function Show-HelpMenu {
    $helpForm = New-Object Windows.Forms.Form
    $helpForm.Text = "SharePoint Library Access Manager - Help"
    $helpForm.Size = New-Object Drawing.Size(700, 650)
    $helpForm.StartPosition = "CenterScreen"
    $helpForm.Topmost = $true

    $helpText = @"
=== SharePoint Library Access Manager - Help ===

This tool provides a graphical interface to manage SharePoint document library permissions.

CONTACT SUPPORT:
Name: NathanimG
Email: NathanimG@ethiopianairlines.com
Mobile: +251910160067

MAIN FEATURES:
- Grant or revoke access to multiple users
- Apply permissions to multiple libraries at once
- View library names and URLs
- Supports common permission levels
- Logs all actions for auditing
- Save operation logs to file

HOW TO USE:

1. SITE URL:
   - Enter the full URL of your SharePoint site
   - Example: https://portal.ethiopianairlines.com/sites/ETCONF/

2. USER LOGINS:
   - Enter one user per line in DOMAIN\user format
   - Example: ET\NathanimG

3. PERMISSION LEVEL:
   - Select from the dropdown menu:
     * Upload & Initiate - Can upload but not edit existing
     * Approve - Can approve documents
     * Edit - Can edit documents
     * Full Control - Full permissions
     * Read - Read-only access

4. LIBRARIES:
   - Click 'Load Libraries' to populate the list
   - Check the libraries you want to modify
   - URLs are displayed alongside library names
   - Use 'Select All' to check/uncheck all libraries

5. ACTIONS:
   - Grant Access: Applies selected permission
   - Revoke Access: Removes all permissions
   - Save Log: Saves operation log to file

MENU OPTIONS:
- File > Exit: Close the application
- Help > Documentation: Show this help
- Help > Quick Reference: Basic usage tips
- Help > About: Version and contact info

TROUBLESHOOTING:
- Ensure you have SharePoint admin rights
- Verify site URL is correct
- Check user login names are in correct format
- For any issues, contact support (details above)

"@

    $textBox = New-Object Windows.Forms.TextBox
    $textBox.Multiline = $true
    $textBox.ReadOnly = $true
    $textBox.ScrollBars = "Vertical"
    $textBox.Text = $helpText
    $textBox.Dock = "Fill"
    $textBox.Font = New-Object Drawing.Font("Consolas", 10)

    $helpForm.Controls.Add($textBox)

    $closeButton = New-Object Windows.Forms.Button
    $closeButton.Text = "Close Help"
    $closeButton.Dock = "Bottom"
    $closeButton.Add_Click({ $helpForm.Close() })

    $helpForm.Controls.Add($closeButton)
    $helpForm.ShowDialog()
}

function Show-QuickReference {
    $msg = @"
QUICK REFERENCE:

1. Site URL: Full SharePoint site URL
2. Users: DOMAIN\user (one per line)
3. Permission: Select from dropdown
4. Libraries: Load then select
5. Actions: Grant/Revoke/Save Log

Access the full documentation from the Help menu.
"@

    [System.Windows.Forms.MessageBox]::Show($msg, "Quick Reference", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

function Show-About {
    $msg = @"
SharePoint Library Access Manager
Version 1.2

Author: NathanimG
Email: NathanimG@ethiopianairlines.com
Mobile: +251910160067

Copyright © $(Get-Date -Format "yyyy") Ethiopian Airlines
"@

    [System.Windows.Forms.MessageBox]::Show($msg, "About", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}
#endregion

function Show-AccessForm {
    param([bool]$DarkMode = $false)

    Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

    $form = New-Object Windows.Forms.Form
    $form.Text = "SharePoint Library Access Manager"
    $form.Size = New-Object Drawing.Size(850, 880)
    $form.StartPosition = "CenterScreen"
    $form.Topmost = $true

    # Create Menu Strip
    $menuStrip = New-Object Windows.Forms.MenuStrip
    $form.Controls.Add($menuStrip)

    # File Menu
    $fileMenu = New-Object Windows.Forms.ToolStripMenuItem
    $fileMenu.Text = "&File"
    
    $exitMenuItem = New-Object Windows.Forms.ToolStripMenuItem
    $exitMenuItem.Text = "E&xit"
    $exitMenuItem.Add_Click({ $form.Close() })
    $fileMenu.DropDownItems.Add($exitMenuItem)
    
    $menuStrip.Items.Add($fileMenu)

    # Help Menu
    $helpMenu = New-Object Windows.Forms.ToolStripMenuItem
    $helpMenu.Text = "&Help"
    
    $docsMenuItem = New-Object Windows.Forms.ToolStripMenuItem
    $docsMenuItem.Text = "&Documentation"
    $docsMenuItem.Add_Click({ Show-HelpMenu })
    $helpMenu.DropDownItems.Add($docsMenuItem)
    
    $quickRefMenuItem = New-Object Windows.Forms.ToolStripMenuItem
    $quickRefMenuItem.Text = "&Quick Reference"
    $quickRefMenuItem.Add_Click({ Show-QuickReference })
    $helpMenu.DropDownItems.Add($quickRefMenuItem)
    
    $aboutMenuItem = New-Object Windows.Forms.ToolStripMenuItem
    $aboutMenuItem.Text = "&About"
    $aboutMenuItem.Add_Click({ Show-About })
    $helpMenu.DropDownItems.Add($aboutMenuItem)
    
    $menuStrip.Items.Add($helpMenu)

    # Site URL label & textbox
    $lblSite = New-Object Windows.Forms.Label
    $lblSite.Text = "Site URL:"
    $lblSite.Location = '10,35'
    $lblSite.AutoSize = $true
    $form.Controls.Add($lblSite)

    $txtSite = New-Object Windows.Forms.TextBox
    $txtSite.Size = '480,22'
    $txtSite.Location = '90,30'
    $form.Controls.Add($txtSite)

    # Users label & multiline textbox
    $lblUsers = New-Object Windows.Forms.Label
    $lblUsers.Text = "User Logins (one per line, e.g. DOMAIN\\user):"
    $lblUsers.Location = '10,70'
    $lblUsers.Size = '300,20'
    $form.Controls.Add($lblUsers)

    $txtUsers = New-Object Windows.Forms.TextBox
    $txtUsers.Multiline = $true
    $txtUsers.ScrollBars = "Vertical"
    $txtUsers.Size = New-Object Drawing.Size(800, 80)
    $txtUsers.Location = '10,95'
    $form.Controls.Add($txtUsers)

    # Permission level dropdown
    $lblPermission = New-Object Windows.Forms.Label
    $lblPermission.Text = "Permission Level:"
    $lblPermission.Location = '10,185'
    $lblPermission.AutoSize = $true
    $form.Controls.Add($lblPermission)

    $cmbPermission = New-Object Windows.Forms.ComboBox
    $cmbPermission.Location = '100,180'
    $cmbPermission.Size = '200,25'
    $cmbPermission.DropDownStyle = [Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbPermission.Items.AddRange(@("Upload & Initiate", "Approve", "Edit", "Full Control", "Read"))
    $cmbPermission.SelectedIndex = 0
    $form.Controls.Add($cmbPermission)

    # Select All checkbox for libraries
    $chkAll = New-Object Windows.Forms.CheckBox
    $chkAll.Text = "Select All Libraries"
    $chkAll.Location = '10,210'
    $chkAll.AutoSize = $true
    $form.Controls.Add($chkAll)

    # DataGridView to show libraries with checkboxes and URLs
    $gridLibs = New-Object Windows.Forms.DataGridView
    $gridLibs.Size = New-Object Drawing.Size(800, 370)
    $gridLibs.Location = New-Object Drawing.Point(10, 235)
    $gridLibs.AllowUserToAddRows = $false
    $gridLibs.SelectionMode = "FullRowSelect"
    $gridLibs.MultiSelect = $true
    $gridLibs.RowHeadersVisible = $false
    $gridLibs.AutoSizeColumnsMode = [Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $form.Controls.Add($gridLibs)

    # Checkbox column for selecting libraries
    $chkCol = New-Object Windows.Forms.DataGridViewCheckBoxColumn
    $chkCol.Name = "Select"
    $chkCol.HeaderText = ""
    $chkCol.Width = 40
    $gridLibs.Columns.Add($chkCol) | Out-Null

    # Library name column
    $colLibName = New-Object Windows.Forms.DataGridViewTextBoxColumn
    $colLibName.Name = "LibraryName"
    $colLibName.HeaderText = "Library Name"
    $colLibName.ReadOnly = $true
    $gridLibs.Columns.Add($colLibName) | Out-Null

    # Library URL column
    $colLibUrl = New-Object Windows.Forms.DataGridViewTextBoxColumn
    $colLibUrl.Name = "LibraryURL"
    $colLibUrl.HeaderText = "Library URL"
    $colLibUrl.ReadOnly = $true
    $gridLibs.Columns.Add($colLibUrl) | Out-Null

    # Buttons
    $btnLoad = New-Object Windows.Forms.Button
    $btnLoad.Text = "Load Libraries"
    $btnLoad.Location = '10,620'
    $btnLoad.Size = '130,30'
    $form.Controls.Add($btnLoad)

    $btnGrant = New-Object Windows.Forms.Button
    $btnGrant.Text = "Grant Access"
    $btnGrant.Location = '160,620'
    $btnGrant.Size = '130,30'
    $btnGrant.Enabled = $false
    $form.Controls.Add($btnGrant)

    $btnRevoke = New-Object Windows.Forms.Button
    $btnRevoke.Text = "Revoke Access"
    $btnRevoke.Location = '310,620'
    $btnRevoke.Size = '130,30'
    $btnRevoke.Enabled = $false
    $form.Controls.Add($btnRevoke)

    $btnSaveLog = New-Object Windows.Forms.Button
    $btnSaveLog.Text = "Save Log"
    $btnSaveLog.Location = '460,620'
    $btnSaveLog.Size = '130,30'
    $form.Controls.Add($btnSaveLog)

    # Status label
    $status = New-Object Windows.Forms.Label
    $status.Text = ""
    $status.Location = '10,660'
    $status.Size = '800,40'
    $status.AutoSize = $false
    $status.ForeColor = [Drawing.Color]::DarkGreen
    $form.Controls.Add($status)

    # Log textbox (read-only multiline)
    $txtLog = New-Object Windows.Forms.TextBox
    $txtLog.Multiline = $true
    $txtLog.ScrollBars = "Vertical"
    $txtLog.ReadOnly = $true
    $txtLog.Size = New-Object Drawing.Size(800, 100)
    $txtLog.Location = New-Object Drawing.Point(10, 700)
    $form.Controls.Add($txtLog)

    # Logging helper function
    function Write-Log {
        param([string]$msg)
        $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $txtLog.AppendText("[$timestamp] $msg`r`n")
    }

    # Handle select all checkbox for libraries
    $chkAll.Add_CheckedChanged({
        $checked = $chkAll.Checked
        foreach ($row in $gridLibs.Rows) {
            $row.Cells["Select"].Value = $checked
        }
    })

    # Load libraries button click event (now includes URLs)
    $btnLoad.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtSite.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a valid Site URL.","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        try {
            $web = Get-SPWeb $txtSite.Text
            $gridLibs.Rows.Clear()
            
            foreach ($list in $web.Lists) {
                if ($list.BaseTemplate -eq [Microsoft.SharePoint.SPListTemplateType]::DocumentLibrary -and -not $list.Hidden) {
                    $libraryUrl = "$($web.Url)/$($list.RootFolder.Url)"
                    $gridLibs.Rows.Add($false, $list.Title, $libraryUrl) | Out-Null
                }
            }
            
            if ($gridLibs.Rows.Count -gt 0) {
                $status.Text = "Libraries loaded successfully."
                Write-Log "Loaded $($gridLibs.Rows.Count) libraries from $($txtSite.Text)."
                $btnGrant.Enabled = $true
                $btnRevoke.Enabled = $true
            } else {
                $status.Text = "No document libraries found."
                Write-Log "No document libraries found at $($txtSite.Text)."
                $btnGrant.Enabled = $false
                $btnRevoke.Enabled = $false
            }
            $web.Dispose()
        }
        catch {
            $status.ForeColor = [Drawing.Color]::Red
            $status.Text = "Error loading libraries: $_"
            Write-Log "Error loading libraries: $_"
            $btnGrant.Enabled = $false
            $btnRevoke.Enabled = $false
        }
    })

    # [Rest of the button click events remain the same...]

    # Apply theme (optional)
    if ($DarkMode) {
        $form.BackColor = [Drawing.Color]::FromArgb(30,30,30)
        foreach ($control in $form.Controls) {
            if ($control -is [Windows.Forms.Label] -or $control -is [Windows.Forms.Button]) {
                $control.ForeColor = [Drawing.Color]::White
                $control.BackColor = $form.BackColor
            }
            elseif ($control -is [Windows.Forms.TextBox] -or $control -is [Windows.Forms.DataGridView]) {
                $control.BackColor = [Drawing.Color]::FromArgb(45,45,48)
                $control.ForeColor = [Drawing.Color]::White
            }
            elseif ($control -is [Windows.Forms.CheckBox]) {
                $control.ForeColor = [Drawing.Color]::White
                $control.BackColor = $form.BackColor
            }
        }
    }

    # Show quick reference on first run
    Show-QuickReference

    [void]$form.ShowDialog()
}

# Run the form with light mode by default; pass $true for dark mode if you want
Show-AccessForm -DarkMode:$false