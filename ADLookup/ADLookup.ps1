Add-Type -AssemblyName System.Windows.Forms

# Function to check and install Active Directory module
function Ensure-ActiveDirectoryModule {
    # Function to check if the script is running as Administrator
    function Is-Administrator {
        return ([bool](net session 2>$null))
    }

    # Self-elevation logic
    if (-not (Is-Administrator)) {
        $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        Start-Process -FilePath "powershell.exe" -ArgumentList $arguments -Verb RunAs
        exit
    }

    # Check if the Active Directory module is already available
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "The Active Directory module for PowerShell is not installed. Would you like to install it now?",
            "Module Missing",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                # Determine the environment and install the module accordingly
                if (Get-WindowsFeature -Name RSAT-AD-PowerShell -ErrorAction SilentlyContinue) {
                    # Install RSAT-AD-PowerShell on server environments
                    Install-WindowsFeature -Name RSAT-AD-PowerShell -IncludeManagementTools -ErrorAction Stop
                } elseif ((Get-WindowsCapability -Name RSAT.ActiveDirectory.DS-LDS.Tools* -Online -ErrorAction SilentlyContinue).State -eq "NotPresent") {
                    # Install RSAT tools on modern Windows clients
                    Add-WindowsCapability -Online -Name RSAT.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 -ErrorAction Stop
                } else {
                    # Fallback to installing module via PowerShell Gallery
                    Install-Module -Name ActiveDirectory -Force -ErrorAction Stop
                }
                Import-Module ActiveDirectory -ErrorAction Stop
                [System.Windows.Forms.MessageBox]::Show(
                    "The Active Directory module has been successfully installed.",
                    "Success",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            } catch {
                [System.Windows.Forms.MessageBox]::Show(
                    "Failed to install the Active Directory module. Please ensure you have administrative privileges and internet access.",
                    "Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
                return
            }
        } else {
            # User chose not to install the module
            [System.Windows.Forms.MessageBox]::Show(
                "The script cannot proceed without the Active Directory module.",
                "Module Missing",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return
        }
    } else {
        # Module is already installed, just import it
        Import-Module ActiveDirectory -ErrorAction Stop
        [System.Windows.Forms.MessageBox]::Show(
            "The Active Directory module is already available.",
            "Module Available",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
}

# Ensure the Active Directory module is available
Ensure-ActiveDirectoryModule

# Main Window
$form = New-Object System.Windows.Forms.Form
$form.Text = "AD Group Comparison"
$form.Size = New-Object System.Drawing.Size(400, 600)

# Username/Name label
$userLabel = New-Object System.Windows.Forms.Label
$userLabel.Text = "Enter Username/First and Last Name:"
$userLabel.AutoSize = $true
$userLabel.Location = New-Object System.Drawing.Point(10, 10)
$form.Controls.Add($userLabel)

# Username/Name textbox
$userTextbox = New-Object System.Windows.Forms.TextBox
$userTextbox.Size = New-Object System.Drawing.Size(360, 20)
$userTextbox.Location = New-Object System.Drawing.Point(10, 30)
$form.Controls.Add($userTextbox)

# Check button
$button = New-Object System.Windows.Forms.Button
$button.Text = "Check Groups"
$button.Size = New-Object System.Drawing.Size(150, 30)
$button.Location = New-Object System.Drawing.Point(10, 60)
$form.Controls.Add($button)

# Current Groups label
$currentGroupsLabel = New-Object System.Windows.Forms.Label
$currentGroupsLabel.Text = "Current Groups:"
$currentGroupsLabel.AutoSize = $true
$currentGroupsLabel.Location = New-Object System.Drawing.Point(10, 100)
$form.Controls.Add($currentGroupsLabel)

# Current Groups textbox
$currentGroupsTextbox = New-Object System.Windows.Forms.TextBox
$currentGroupsTextbox.Multiline = $true
$currentGroupsTextbox.Size = New-Object System.Drawing.Size(360, 100)
$currentGroupsTextbox.ScrollBars = "Vertical"
$currentGroupsTextbox.Location = New-Object System.Drawing.Point(10, 120)
$form.Controls.Add($currentGroupsTextbox)

# Comparison Groups label
$label = New-Object System.Windows.Forms.Label
$label.Text = "Comparison Groups:"
$label.AutoSize = $true
$label.Location = New-Object System.Drawing.Point(10, 230)
$form.Controls.Add($label)

# Comparison Groups textbox
$textbox = New-Object System.Windows.Forms.TextBox
$textbox.Multiline = $true
$textbox.Size = New-Object System.Drawing.Size(360, 100)
$textbox.ScrollBars = "Vertical"
$textbox.Location = New-Object System.Drawing.Point(10, 250)
$form.Controls.Add($textbox)

# Missing Groups label
$missingGroupsLabel = New-Object System.Windows.Forms.Label
$missingGroupsLabel.Text = "Missing Groups:"
$missingGroupsLabel.AutoSize = $true
$missingGroupsLabel.Location = New-Object System.Drawing.Point(10, 360)
$form.Controls.Add($missingGroupsLabel)

# Missing Groups textbox
$missingGroupsTextbox = New-Object System.Windows.Forms.TextBox
$missingGroupsTextbox.Multiline = $true
$missingGroupsTextbox.Size = New-Object System.Drawing.Size(360, 100)
$missingGroupsTextbox.ScrollBars = "Vertical"
$missingGroupsTextbox.Location = New-Object System.Drawing.Point(10, 380)
$form.Controls.Add($missingGroupsTextbox)

# Check groups function
$button.Add_Click({
    $inputValue = $userTextbox.Text.Trim() # Retrieve and trim the input text

    try {
        # Attempt to retrieve domain controllers
        $domainControllers = Get-ADDomainController -Filter * -ErrorAction Stop

        if (-not $domainControllers) {
            [System.Windows.Forms.MessageBox]::Show("No domain controllers found.", "Error")
            return
        }

        $latestUserData = $null
        $latestTimestamp = [datetime]::MinValue
        $mostRecentDC = $null # Variable to track the domain controller with the most up-to-date info

        foreach ($dc in $domainControllers) {
            try {
                # Log the domain controller being queried
                Write-Host "Querying Domain Controller: $($dc.HostName)"

                # Query user details from each domain controller
                $currentUserData = if ($inputValue -notmatch '\s') {
                    Get-ADUser -Filter "SamAccountName -eq '$inputValue'" -Server $dc.HostName -Properties MemberOf, WhenChanged
                } else {
                    $parts = $inputValue -split '\s+', 2
                    if ($parts.Count -eq 2) {
                        $firstName = $parts[0]
                        $lastName = $parts[1]
                        Get-ADUser -Filter "GivenName -eq '$firstName' -and Surname -eq '$lastName'" -Server $dc.HostName -Properties MemberOf, WhenChanged
                    } else {
                        throw "Please enter a valid username or both first and last names."
                    }
                }

                # Update the most recent data if necessary
                if ($currentUserData -and $currentUserData.WhenChanged -gt $latestTimestamp) {
                    $latestTimestamp = $currentUserData.WhenChanged
                    $latestUserData = $currentUserData
                    $mostRecentDC = $dc.HostName
                }
            } catch {
                Write-Host "Error querying domain controller $($dc.HostName): $_"
                continue
            }
        }

        # Log the domain controller with the most up-to-date info
        if ($mostRecentDC) {
            Write-Host "Domain Controller with most up-to-date info: $mostRecentDC at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        } else {
            Write-Host "No domain controller provided updated user information."
        }

        if (-not $latestUserData) {
            [System.Windows.Forms.MessageBox]::Show("User not found.", "Error")
            return
        }

        # Explicitly refresh group data from the most recent DC
        $userGroups = if ($inputValue -notmatch '\s') {
            (Get-ADUser -Filter "SamAccountName -eq '$inputValue'" -Server $mostRecentDC -Properties MemberOf).MemberOf
        } else {
            $parts = $inputValue -split '\s+', 2
            if ($parts.Count -eq 2) {
                $firstName = $parts[0]
                $lastName = $parts[1]
                (Get-ADUser -Filter "GivenName -eq '$firstName' -and Surname -eq '$lastName'" -Server $mostRecentDC -Properties MemberOf).MemberOf
            } else {
                throw "Please enter a valid username or both first and last names."
            }
        }

        # Ensure $userGroups is not null or empty
        if ($userGroups -and $userGroups.Count -gt 0) {
            $userGroups = $userGroups | Where-Object { $_ -ne $null } | ForEach-Object {
                (Get-ADGroup -Identity $_ -Server $mostRecentDC).Name
            } | Sort-Object
        } else {
            throw "No groups found for the user."
        }

        # Display current group memberships
        $currentGroupsTextbox.Text = $userGroups -join "`r`n"

        # Compare with desired groups
        $desiredGroups = $textbox.Text -split "`r`n"
        $missingGroups = $desiredGroups | Where-Object { $_ -notin $userGroups } | Sort-Object
        $missingGroupsTextbox.Text = $missingGroups -join "`r`n"
    } catch {
        $errorMessage = $_.Exception.Message
        Write-Host "Error querying Active Directory: $errorMessage"
        [System.Windows.Forms.MessageBox]::Show("Error querying Active Directory: $errorMessage", "Error")
    }
})

# Handle Enter key for username textbox only
$userTextbox.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $button.PerformClick()
        $e.SuppressKeyPress = $true
    }
})

# Function to enable Ctrl+A
function Enable-CtrlA {
    param($textbox)
    $textbox.Add_KeyDown({
        param($sender, $e)
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
            $sender.SelectAll()
            $e.SuppressKeyPress = $true
        }
    })
}

# Apply Ctrl+A to all textboxes
Enable-CtrlA -textbox $userTextbox
Enable-CtrlA -textbox $currentGroupsTextbox
Enable-CtrlA -textbox $textbox
Enable-CtrlA -textbox $missingGroupsTextbox

# Reset button
$resetButton = New-Object System.Windows.Forms.Button
$resetButton.Text = "Reset"
$resetButton.Size = New-Object System.Drawing.Size(80, 30)
$resetButton.Location = New-Object System.Drawing.Point(300, 500) # Adjust as needed for proper placement
$form.Controls.Add($resetButton)

# Reset button click event
$resetButton.Add_Click({
    $userTextbox.Clear()
    $currentGroupsTextbox.Clear()
    $textbox.Clear()
    $missingGroupsTextbox.Clear()
})

# Run the form
[void] $form.ShowDialog()
