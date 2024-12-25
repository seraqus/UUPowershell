Add-Type -AssemblyName System.Windows.Forms

# Function to check and install Active Directory module
function Ensure-ActiveDirectoryModule {
    # Check if the Active Directory module is already available
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        # Show error pop-up and terminate the script
        [System.Windows.Forms.MessageBox]::Show(
            "The Active Directory module is not available on this system. Please install RSAT using the DeploymentShare before continuing.",
            "Module Missing",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        exit
    } else {
        # Module is available, silently import it
        Import-Module ActiveDirectory -ErrorAction Stop
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

$button.Add_Click({
    # Clear the existing fields before continuing
    $currentGroupsTextbox.Text = ""
    $textbox.Text = ""
    $missingGroupsTextbox.Text = ""
        
    # Refresh the user's group membership data
    Clear-Variable -Name userGroups -Scope Global -ErrorAction SilentlyContinue
    Write-Host "Clearing user group cache..."

    # Retrieve and trim the input text
    $inputValue = $userTextbox.Text.Trim()

    try {
        # Attempt to retrieve ALL domain controllers
        $domainControllers = Get-ADDomainController -Filter * -ErrorAction Stop

        #Error handling when unable to locate Domain Controllers
        if (-not $domainControllers) {
            [System.Windows.Forms.MessageBox]::Show("No domain controllers found.", "Error")
            return
        }

        $latestUserData = $null
        $latestTimestamp = [datetime]::MinValue
        $mostRecentDC = $null 

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

        # Display current group memberships
        $currentGroupsTextbox.Text = $userGroups -join "`r`n"

        # Compare with desired groups
        $desiredGroups = $textbox.Text -split "`r`n"
        $missingGroups = $desiredGroups | Where-Object { $_ -notin $userGroups } | Sort-Object
        $missingGroupsTextbox.Text = $missingGroups -join "`r`n"

    #Error Handling for when unable to query AD    
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
