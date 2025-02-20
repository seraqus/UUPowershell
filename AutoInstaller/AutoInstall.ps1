Add-Type -AssemblyName System.Windows.Forms

## Variables ====================================================================================

#$global:domainControllers = $null

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
        exit
    }
}

#Functions to Get all domain controllers, this will be used a later point possible to query a username against in order to obtain the most updated info as permission to force replication is not available
function Get-DomainControllers {
    if ($null -eq $global:domainControllers) {
        $global:domainControllers = Get-ADDomainController -Filter *
        foreach ($dc in $global:domainControllers) {
            Write-Host "Domain Controller found: $($dc.Name)"
        }
    } else {
        Write-Host "Domain controllers are already loaded."
    }
}

# Function to enable Ctrl+A on the specified -textbox
function Enable-CtrlA {
    param($textbox)
    $textbox.Add_KeyDown({
        param($src, $e)
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
            $src.SelectAll()
            $e.SuppressKeyPress = $true
        }
    })
}

# Function to reset all Text Fields
function Reset-TextFields {
    #$userTextboxInfo.Clear()
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

# Function to search for AD User and update the Search box with username
function Get-ADUserFromTextbox {
    param (
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.TextBox]$textbox
    )

    $searchValue = $textbox.Text.Trim()

    if ($searchValue -eq '') {
        [System.Windows.Forms.MessageBox]::Show('Please enter a user to search for.')
        Write-Host "No user has been entered."
        return
    }

    Write-Host "Searching for user with value: $searchValue"

    $searchValue = [System.Text.RegularExpressions.Regex]::Escape($searchValue)
    $matchedUsers = @()

    $nameParts = $searchValue.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)
    $firstName = $nameParts[0]
    $lastName = if ($nameParts.Length -gt 1) { $nameParts[1] } else { "" }

    $firstName = [System.Text.RegularExpressions.Regex]::Escape($firstName)
    $lastName = [System.Text.RegularExpressions.Regex]::Escape($lastName)

    $user = Get-ADUser -LDAPFilter "(sAMAccountName=$searchValue)" -Properties SamAccountName,  DisplayName, Mail
    if ($user) {
        Write-Host "Found user by Username: $($user.SamAccountName)"
        $matchedUsers += $user
    }
    if (-not $user) {
        $users = Get-ADUser -Filter "(displayName -like '*$searchValue*')" -Properties DisplayName,  SamAccountName, Mail
        if ($users) {
            Write-Host "Found user by Display Name: $($users | ForEach-Object { $_.DisplayName } | ForEach-Object -Begin { $result = @() } -Process { $result += $_ } -End { $result -join ', ' })"
            $matchedUsers += $users
        }
    }
    if (-not $user) {
        $users = Get-ADUser -Filter "(mail -like '*$searchValue*')" -Properties Mail, SamAccountName, DisplayName
        if ($users) {
            Write-Host "Found user by Email: $($users | ForEach-Object { $_.mail } | ForEach-Object -Begin { $result = @() } -Process { $result += $_ } -End { $result -join ', ' })"
            $matchedUsers += $users
        }
    }

    $matchedUsers = $matchedUsers | Sort-Object -Property SamAccountName -Unique
    if ($matchedUsers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No matching user found in Active Directory.')
        Write-Host "No matching user found."
        return
    }
    if ($matchedUsers.Count -eq 1) {
        $textbox.Text = $matchedUsers[0].SamAccountName
    }
    elseif ($matchedUsers.Count -gt 1) {
        $selectionForm = New-Object System.Windows.Forms.Form
        $selectionForm.Text = 'Select a User'
        $selectionForm.Size = New-Object System.Drawing.Size(257.5, 180)

        $listBox = New-Object System.Windows.Forms.ListBox
        $listBox.Size = New-Object System.Drawing.Size(230, 110)
        $listBox.Location = New-Object System.Drawing.Point(5, 5)
        $listBox.SelectionMode = [System.Windows.Forms.SelectionMode]::One
        $matchedUsers | ForEach-Object {
            $displayName = $_.DisplayName
            $listBox.Items.Add($displayName)
        }

        $selectButton = New-Object System.Windows.Forms.Button
        $selectButton.Text = "Select"
        $selectButton.Size = New-Object System.Drawing.Size(45, 20)
        $selectButton.Location = New-Object System.Drawing.Point(190, 115)
        $selectButton.Add_Click({
            $selectedUser = $matchedUsers[$listBox.SelectedIndex]
            $textbox.Text = $selectedUser.SamAccountName
            $selectionForm.Close()
        })

        $selectionForm.Controls.Add($listBox)
        $selectionForm.Controls.Add($selectButton)
        $selectionForm.ShowDialog()
    }
}

#Function to gather User AD Groups
function Get-ADUserGroups {
    param (
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.TextBox]$inputTextbox,
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.TextBox]$outputTextbox
    )
    Clear-Variable -Name userGroups -Scope Global -ErrorAction SilentlyContinue
    Write-Host "User group cache cleared." 
    $username = $inputTextbox.Text
    if ($username -eq '') {
        Write-Host "No user to populate groups from."
        return
    }
    Write-Host "User groups found."
    $userGroups = (Get-ADUser -Identity $username -Properties MemberOf).MemberOf
    $outputTextbox.Text = ($userGroups | ForEach-Object { $_ -replace '^CN=([^,]+).+$', '$1' }) -join "`r`n"
}

#Function to compare 2 textboxes and output to a 3rd
function Compare-Textboxes {
    param (
        [System.Windows.Forms.TextBox]$inputTextbox,
        [System.Windows.Forms.TextBox]$desiredTextbox,
        [System.Windows.Forms.TextBox]$outputTextbox
    )
    $desiredGroups = $desiredTextbox.Text -split "`r`n"
    $userGroups = $inputTextbox.Text.Split("`r`n") | ForEach-Object { $_ -replace '^CN=([^,]+).+$','$1' }
    $missingGroups = $desiredGroups | Where-Object { $_ -notin $userGroups }
    $outputTextbox.Text = $missingGroups -join "`r`n"
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

        <## Info Sub Tab
        $userInfoSubTab = New-Object System.Windows.Forms.TabPage
        $userInfoSubTab.Text = "Info"
        $userSubTabControl.TabPages.Add($userInfoSubTab)

            # Username/Name label
            $userLabelInfo = New-Object System.Windows.Forms.Label
            $userLabelInfo.Text = "Enter Username, First or Last Name:"
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
            $searchButtonInfo.Add_Click({
                Get-ADUserFromTextBox -textbox $userTextboxInfo
            })
            
            # Reset button
            $resetButtonInfo = New-Object System.Windows.Forms.Button
            $resetButtonInfo.Text = "Reset"
            $resetButtonInfo.Size = New-Object System.Drawing.Size(45, 20)
            $resetButtonInfo.Location = New-Object System.Drawing.Point(150, 25)
            $resetButtonInfo.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
            $userInfoSubTab.Controls.Add($resetButtonInfo)
            $resetButtonInfo.Add_Click({
                Reset-TextFields
            })#>

        # AD Groups Sub Tab
        $adGroupsSubTab = New-Object System.Windows.Forms.TabPage
        $adGroupsSubTab.Text = "AD Groups"
        $userSubTabControl.TabPages.Add($adGroupsSubTab)

            # Username/Name label
            $userLabelGroups = New-Object System.Windows.Forms.Label
            $userLabelGroups.Text = "Username, First or Last Name:"
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
            $searchButtonGroups.Add_Click({
                Reset-GroupFields
                Get-ADUserFromTextbox -textbox $userTextboxGroups
                Get-ADUserGroups -inputTextbox $userTextboxGroups -outputTextbox $currentGroupsTextbox

            })

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
            $label.Text = "Comparison Groups:"
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
                Get-ADUserFromTextbox -textbox $userTextboxGroups
                Get-ADUserGroups -inputTextbox $userTextboxGroups -outputTextbox $currentGroupsTextbox
                Compare-Textboxes -inputTextbox $currentGroupsTextbox -desiredTextbox $compareGroupsTextbox -outputTextbox $missingGroupsTextbox
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
    #$computerTab = New-Object System.Windows.Forms.TabPage
    #$computerTab.Text = "Computer"
    #$mainTabControl.TabPages.Add($computerTab)

[void] $form.ShowDialog()
