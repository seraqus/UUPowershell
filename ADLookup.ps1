Add-Type -AssemblyName System.Windows.Forms

## Variables ====================================================================================

$global:userValue = $null
$global:domainControllers = $null

## Functions ====================================================================================

# Function to check and install Active Directory module
function Test-ActiveDirectoryModule {
    if (Get-Module -ListAvailable -Name ActiveDirectory) {
        Import-Module ActiveDirectory -ErrorAction Stop
    } else {
        [System.Windows.Forms.MessageBox]::Show(
            "The Active Directory module is not available on this system. Please verify RSAT is installed before continuing.",
            "Module Missing",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        #exit
    }
}

# Function to reset all Text Fields
function Reset-TextFields {
    $userTextboxInfo.Clear()
    $userTextboxGroups.Clear()
    $currentGroupsTextbox.Clear()
    $compareGroupsTextbox.Clear()
    $missingGroupsTextbox.Clear()
    Write-Host "All text fields reset."
}

# Function to reset all Group Fields
function Reset-GroupFields {
    $currentGroupsTextbox.Clear()
    $missingGroupsTextbox.Clear()
    Write-Host "All group fields reset."
}

# Function to enable Ctrl+A
function Enable-CtrlA {
    param($textbox)
    $textbox.Add_KeyDown({
        param($src, $e)
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
            $sender.SelectAll()
            $e.SuppressKeyPress = $true
        }
    })
}

# Function to reset local user group cache
function Reset-LocalGroupCache {
    Clear-Variable -Name userGroups -Scope Global -ErrorAction SilentlyContinue
    Write-Host "User group cache cleared."    
}

#Functions to query all domain controllers
function Get-DomainControllers {
    if ($null -eq $global:domainControllers) {
        $global:domainControllers = Get-ADDomainController -Filter *
        foreach ($dc in $global:domainControllers) {
            Write-Host "Domain Controller found: $($dc.Name)"
        }
        return $global:domainControllers
    } else {
        Write-Host "Domain controllers are already loaded."
    }
}

# Function to gather user text field info and trim into a global variable
function Get-UserSearch {
    param (
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.TextBox]$textbox
    )
    if ($null -eq $textbox.Text) {
        [System.Windows.Forms.MessageBox]::Show("No user value entered.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        Write-Host "No user info provided."
        return $null
    } else {
        $global:userValue = $textbox.Text.Trim()
        Write-Host "Search value entered: $($textbox.Text)"
        return $true
    }
}

# Function to use $global:userValue to search for user by username, first/last name, first & last name, or email
function Search-User {
    if ($null -eq $global:userValue) {
        [System.Windows.Forms.MessageBox]::Show("No user value entered.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        Write-Host "No user info provided."
        return $null
    }

    try {
        $user = Get-ADUser -Filter {
            (SamAccountName -eq $global:userValue) -or
            (GivenName -eq $global:userValue) -or
            (Surname -eq $global:userValue) -or
            (Name -eq $global:userValue) -or
            (EmailAddress -eq $global:userValue)
        }

        if ($null -eq $user) {
            [System.Windows.Forms.MessageBox]::Show("User not found.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            Write-Host "User not found."
            return $null
        } else {
            Write-Host "User found: $($user.SamAccountName)"
            return $user
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error searching for user: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        Write-Host "Error searching for user: $_"
        return $null
    }
}

## Main Program construction ====================================================================================

# Start-up Functions
Test-ActiveDirectoryModule
Get-DomainControllers

# Main Window
$form = New-Object System.Windows.Forms.Form
$form.Text = "ADLookup"
$form.Size = New-Object System.Drawing.Size(375, 515)
$form.MinimumSize = New-Object System.Drawing.Size(375, 515)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable

# Main Tab Control
$mainTabControl = New-Object System.Windows.Forms.TabControl
$mainTabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
$form.Controls.Add($mainTabControl)

    # User Tab
    $userTab = New-Object System.Windows.Forms.TabPage
    $userTab.Text = "User"
    $mainTabControl.TabPages.Add($userTab)

    # User Sub Tab Control
    $userSubTabControl = New-Object System.Windows.Forms.TabControl
    $userSubTabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
    $userTab.Controls.Add($userSubTabControl)

        # Info Sub Tab
        $userInfoSubTab = New-Object System.Windows.Forms.TabPage
        $userInfoSubTab.Text = "Info"
        $userSubTabControl.TabPages.Add($userInfoSubTab)

            # Username/Name label
            $userLabelInfo = New-Object System.Windows.Forms.Label
            $userLabelInfo.Text = "Enter Username/First or Last Name:"
            $userLabelInfo.AutoSize = $true
            $userLabelInfo.Location = New-Object System.Drawing.Point(5, 5)
            $userInfoSubTab.Controls.Add($userLabelInfo)

            # Username/Name textbox
            $userTextboxInfo = New-Object System.Windows.Forms.TextBox
            $userTextboxInfo.Size = New-Object System.Drawing.Size(85, 20)
            $userTextboxInfo.Location = New-Object System.Drawing.Point(5, 25)
            $userTextboxInfo.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
            $userInfoSubTab.Controls.Add($userTextboxInfo)
            Enable-CtrlA -textbox $userTextboxInfo

            # Search button
            $searchButtonInfo = New-Object System.Windows.Forms.Button
            $searchButtonInfo.Text = "Search"
            $searchButtonInfo.Size = New-Object System.Drawing.Size(50, 20)
            $searchButtonInfo.Location = New-Object System.Drawing.Point(95, 25)
            $searchButtonInfo.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
            $userInfoSubTab.Controls.Add($searchButtonInfo)

            # Reset button
            $resetButtonInfo = New-Object System.Windows.Forms.Button
            $resetButtonInfo.Text = "Reset"
            $resetButtonInfo.Size = New-Object System.Drawing.Size(45, 20)
            $resetButtonInfo.Location = New-Object System.Drawing.Point(150, 25)
            $resetButtonInfo.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
            $userInfoSubTab.Controls.Add($resetButtonInfo)
            $resetButtonInfo.Add_Click({
                Reset-TextFields
            })

        # AD Groups Sub Tab
        $adGroupsSubTab = New-Object System.Windows.Forms.TabPage
        $adGroupsSubTab.Text = "AD Groups"
        $userSubTabControl.TabPages.Add($adGroupsSubTab)

            # Username/Name label
            $userLabelGroups = New-Object System.Windows.Forms.Label
            $userLabelGroups.Text = "Enter Username/First or Last Name:"
            $userLabelGroups.AutoSize = $true
            $userLabelGroups.Location = New-Object System.Drawing.Point(5, 5)
            $adGroupsSubTab.Controls.Add($userLabelGroups)

            # Username/Name textbox
            $userTextboxGroups = New-Object System.Windows.Forms.TextBox
            $userTextboxGroups.Size = New-Object System.Drawing.Size(85, 20)
            $userTextboxGroups.Location = New-Object System.Drawing.Point(5, 25)
            $userTextboxGroups.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
            $adGroupsSubTab.Controls.Add($userTextboxGroups)
            Enable-CtrlA -textbox $userTextboxGroups

            # Search button
            $searchButtonGroups = New-Object System.Windows.Forms.Button
            $searchButtonGroups.Text = "Search"
            $searchButtonGroups.Size = New-Object System.Drawing.Size(50, 20)
            $searchButtonGroups.Location = New-Object System.Drawing.Point(95, 25)
            $searchButtonGroups.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
            $adGroupsSubTab.Controls.Add($searchButtonGroups)

            # Reset button
            $resetButtonGroups = New-Object System.Windows.Forms.Button
            $resetButtonGroups.Text = "Reset"
            $resetButtonGroups.Size = New-Object System.Drawing.Size(45, 20)
            $resetButtonGroups.Location = New-Object System.Drawing.Point(150, 25)
            $resetButtonGroups.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
            $adGroupsSubTab.Controls.Add($resetButtonGroups)
            $resetButtonGroups.Add_Click({
                Reset-TextFields
            })

            # Current Groups label
            $currentGroupsLabel = New-Object System.Windows.Forms.Label
            $currentGroupsLabel.Text = "Current Groups:"
            $currentGroupsLabel.AutoSize = $true
            $currentGroupsLabel.Location = New-Object System.Drawing.Point(5, 50)
            $adGroupsSubTab.Controls.Add($currentGroupsLabel)

            # Current Groups textbox
            $currentGroupsTextbox = New-Object System.Windows.Forms.TextBox
            $currentGroupsTextbox.Multiline = $true
            $currentGroupsTextbox.Size = New-Object System.Drawing.Size(190, 100)
            $currentGroupsTextbox.ScrollBars = "Vertical"
            $currentGroupsTextbox.Location = New-Object System.Drawing.Point(5, 70)
            $currentGroupsTextbox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
            $adGroupsSubTab.Controls.Add($currentGroupsTextbox)
            Enable-CtrlA -textbox $currentGroupsTextbox

            # Copy Current Groups button
            $currentGroupsButton = New-Object System.Windows.Forms.Button
            $currentGroupsButton.Text = "Copy Groups"
            $currentGroupsButton.Size = New-Object System.Drawing.Size(80, 20)
            $currentGroupsButton.Location = New-Object System.Drawing.Point(115, 47.5)
            $currentGroupsButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
            $adGroupsSubTab.Controls.Add($currentGroupsButton)
            $currentGroupsButton.Add_Click({
                $copiedText = $currentGroupsTextbox.Text
                Set-Clipboard -Value $copiedText
                Write-Host "Text copied: $copiedText"
            })

            # Comparison Groups label
            $label = New-Object System.Windows.Forms.Label
            $label.Text = "Comparison Groups/User:"
            $label.AutoSize = $true
            $label.Location = New-Object System.Drawing.Point(5, 175)
            $adGroupsSubTab.Controls.Add($label)

            # Comparison Groups textbox
            $compareGroupsTextbox = New-Object System.Windows.Forms.TextBox
            $compareGroupsTextbox.Multiline = $true
            $compareGroupsTextbox.Size = New-Object System.Drawing.Size(190, 100)
            $compareGroupsTextbox.ScrollBars = "Vertical"
            $compareGroupsTextbox.Location = New-Object System.Drawing.Point(5, 195)
            $compareGroupsTextbox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
            $adGroupsSubTab.Controls.Add($compareGroupsTextbox)
            Enable-CtrlA -textbox $compareGroupsTextbox

            # Compare button
            $compareButton = New-Object System.Windows.Forms.Button
            $compareButton.Text = "Compare"
            $compareButton.Size = New-Object System.Drawing.Size(60, 20)
            $compareButton.Location = New-Object System.Drawing.Point(135, 172.5)
            $compareButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
            $adGroupsSubTab.Controls.Add($compareButton)
            $compareButton.Add_Click({
                Reset-GroupFields
                Reset-LocalGroupCache
                Get-UserSearch -textbox $userTextboxGroups
            })

            # Missing Groups label
            $missingGroupsLabel = New-Object System.Windows.Forms.Label
            $missingGroupsLabel.Text = "Missing Groups:"
            $missingGroupsLabel.AutoSize = $true
            $missingGroupsLabel.Location = New-Object System.Drawing.Point(5, 300)
            $adGroupsSubTab.Controls.Add($missingGroupsLabel)

            # Missing Groups textbox
            $missingGroupsTextbox = New-Object System.Windows.Forms.TextBox
            $missingGroupsTextbox.Multiline = $true
            $missingGroupsTextbox.Size = New-Object System.Drawing.Size(190, 100)
            $missingGroupsTextbox.ScrollBars = "Vertical"
            $missingGroupsTextbox.Location = New-Object System.Drawing.Point(5, 320)
            $missingGroupsTextbox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
            $adGroupsSubTab.Controls.Add($missingGroupsTextbox)
            Enable-CtrlA -textbox $missingGroupsTextbox

            # Copy Missing Groups button
            $missingGroupsButton = New-Object System.Windows.Forms.Button
            $missingGroupsButton.Text = "Copy Groups"
            $missingGroupsButton.Size = New-Object System.Drawing.Size(80, 20)
            $missingGroupsButton.Location = New-Object System.Drawing.Point(115, 297.5)
            $missingGroupsButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
            $adGroupsSubTab.Controls.Add($missingGroupsButton)
            $missingGroupsButton.Add_Click({
                $copiedText = $missingGroupsTextbox.Text
                Set-Clipboard -Value $copiedText
                Write-Host "Text copied: $copiedText"
            })

    # Computer Tab
    $computerTab = New-Object System.Windows.Forms.TabPage
    $computerTab.Text = "Computer"
    $mainTabControl.TabPages.Add($computerTab)

## Button Interactions ====================================================================================

# Handle Enter key for username textbox only
$userTextboxInfo.Add_KeyDown({
    param($src, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $compareButton.PerformClick()
        $e.SuppressKeyPress = $true
    }
})

# Run the form
[void] $form.ShowDialog()