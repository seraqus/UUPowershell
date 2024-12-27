Add-Type -AssemblyName System.Windows.Forms

## Variables ====================================================================================

$global:userValue = $null

## Functions ====================================================================================

# Function to check and install Active Directory module
function Ensure-ActiveDirectoryModule {
    # Check if the Active Directory module is already available and import it
    if (Get-Module -ListAvailable -Name ActiveDirectory) {
        Import-Module ActiveDirectory -ErrorAction Stop
    } else {
        # Show error pop-up and terminate the script if not available
        [System.Windows.Forms.MessageBox]::Show(
            "The Active Directory module is not available on this system. Please install RSAT from the DeploymentShare before continuing.",
            "Module Missing",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        exit
    }
}

# Function to reset all Text Fields
function Reset-TextFields {
    $userTextbox.Clear()
    $currentGroupsTextbox.Clear()
    $textbox.Clear()
    $missingGroupsTextbox.Clear()
    Write-Host "All text fields reset."
}

# Function to reset all Group Fields
function Reset-GroupFields {
    $currentGroupsTextbox.Clear()
    $textbox.Clear()
    $missingGroupsTextbox.Clear()
    Write-Host "All group fields reset."
}

#Function to gather user text field info and trim into a global variable
function Retrieve-User {
    $global:userValue = $userTextbox.Text.Trim()
    Write-Host "Selected user: $global:userValue"
}

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

# Function to reset local user group cache
function Reset-LocalGroupCache {
    Clear-Variable -Name userGroups -Scope Global -ErrorAction SilentlyContinue
    Write-Host "User group cache cleared."    
}

## Main Program construction ====================================================================================

# Start-up Functions
Ensure-ActiveDirectoryModule

# Main Window
$form = New-Object System.Windows.Forms.Form
$form.Text = "ADLookup"
$form.Size = New-Object System.Drawing.Size(400, 600)

# Main Tab Control
$mainTabControl = New-Object System.Windows.Forms.TabControl
$mainTabControl.Size = New-Object System.Drawing.Size(400, 600)
$form.Controls.Add($mainTabControl)

    # User Tab
    $userTab = New-Object System.Windows.Forms.TabPage
    $userTab.Text = "User"
    $mainTabControl.TabPages.Add($userTab)

    # User Sub Tab Control
    $userSubTabControl = New-Object System.Windows.Forms.TabControl
    $userSubTabControl.Size = New-Object System.Drawing.Size(400, 600)
    $userTab.Controls.Add($userSubTabControl)

        # Info Sub Tab
        $userInfoSubTab = New-Object System.Windows.Forms.TabPage
        $userInfoSubTab.Text = "Info"
        $userSubTabControl.TabPages.Add($userInfoSubTab)

        # Comparison Sub Tab
        $comparisonSubTab = New-Object System.Windows.Forms.TabPage
        $comparisonSubTab.Text = "Comparison"
        $userSubTabControl.TabPages.Add($comparisonSubTab)

            # Username/Name label
            $userLabel = New-Object System.Windows.Forms.Label
            $userLabel.Text = "Enter Username/First and Last Name:"
            $userLabel.AutoSize = $true
            $userLabel.Location = New-Object System.Drawing.Point(10, 10)
            $comparisonSubTab.Controls.Add($userLabel)

            # Username/Name textbox
            $userTextbox = New-Object System.Windows.Forms.TextBox
            $userTextbox.Size = New-Object System.Drawing.Size(360, 20)
            $userTextbox.Location = New-Object System.Drawing.Point(10, 30)
            $comparisonSubTab.Controls.Add($userTextbox)
            Enable-CtrlA -textbox $userTextbox

            # Compare button
            $compareButton = New-Object System.Windows.Forms.Button
            $compareButton.Text = "Compare"
            $compareButton.Size = New-Object System.Drawing.Size(150, 30)
            $compareButton.Location = New-Object System.Drawing.Point(10, 60)
            $comparisonSubTab.Controls.Add($compareButton)

            # Current Groups label
            $currentGroupsLabel = New-Object System.Windows.Forms.Label
            $currentGroupsLabel.Text = "Current Groups:"
            $currentGroupsLabel.AutoSize = $true
            $currentGroupsLabel.Location = New-Object System.Drawing.Point(10, 100)
            $comparisonSubTab.Controls.Add($currentGroupsLabel)

            # Current Groups textbox
            $currentGroupsTextbox = New-Object System.Windows.Forms.TextBox
            $currentGroupsTextbox.Multiline = $true
            $currentGroupsTextbox.Size = New-Object System.Drawing.Size(360, 100)
            $currentGroupsTextbox.ScrollBars = "Vertical"
            $currentGroupsTextbox.Location = New-Object System.Drawing.Point(10, 120)
            $comparisonSubTab.Controls.Add($currentGroupsTextbox)
            Enable-CtrlA -textbox $currentGroupsTextbox

            # Comparison Groups label
            $label = New-Object System.Windows.Forms.Label
            $label.Text = "Comparison Groups:"
            $label.AutoSize = $true
            $label.Location = New-Object System.Drawing.Point(10, 230)
            $comparisonSubTab.Controls.Add($label)

            # Comparison Groups textbox
            $textbox = New-Object System.Windows.Forms.TextBox
            $textbox.Multiline = $true
            $textbox.Size = New-Object System.Drawing.Size(360, 100)
            $textbox.ScrollBars = "Vertical"
            $textbox.Location = New-Object System.Drawing.Point(10, 250)
            $comparisonSubTab.Controls.Add($textbox)
            Enable-CtrlA -textbox $textbox

            # Missing Groups label
            $missingGroupsLabel = New-Object System.Windows.Forms.Label
            $missingGroupsLabel.Text = "Missing Groups:"
            $missingGroupsLabel.AutoSize = $true
            $missingGroupsLabel.Location = New-Object System.Drawing.Point(10, 360)
            $comparisonSubTab.Controls.Add($missingGroupsLabel)

            # Missing Groups textbox
            $missingGroupsTextbox = New-Object System.Windows.Forms.TextBox
            $missingGroupsTextbox.Multiline = $true
            $missingGroupsTextbox.Size = New-Object System.Drawing.Size(360, 100)
            $missingGroupsTextbox.ScrollBars = "Vertical"
            $missingGroupsTextbox.Location = New-Object System.Drawing.Point(10, 380)
            $comparisonSubTab.Controls.Add($missingGroupsTextbox)
            Enable-CtrlA -textbox $missingGroupsTextbox

            # Reset button
            $resetButton = New-Object System.Windows.Forms.Button
            $resetButton.Text = "Reset"
            $resetButton.Size = New-Object System.Drawing.Size(80, 30)
            $resetButton.Location = New-Object System.Drawing.Point(300, 500)
            $comparisonSubTab.Controls.Add($resetButton)

    # Computer Tab
    $computerTab = New-Object System.Windows.Forms.TabPage
    $computerTab.Text = "Computer"
    $mainTabControl.TabPages.Add($computerTab)

## Button Interactions ====================================================================================

# Comparison Button click
$compareButton.Add_Click({
    Reset-GroupFields
    Reset-LocalGroupCache
    Retrieve-User

    try {
        # Attempt to retrieve domain controllers
        $domainControllers = Get-ADDomainController -Filter * -ErrorAction Stop
        Write-Host "Domain Controllers retrieved: $($domainControllers.Count)"

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
                $currentUserData = if ($global:userValue -notmatch '\s') {
                    Get-ADUser -Filter "SamAccountName -eq '$global:userValue'" -Server $dc.HostName -Properties MemberOf, WhenChanged
                } else {
                    $parts = $global:userValue -split '\s+', 2
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
        $userGroups = if ($global:userValue -notmatch '\s') {
            (Get-ADUser -Filter "SamAccountName -eq '$global:userValue'" -Server $mostRecentDC -Properties MemberOf).MemberOf
        } else {
            $parts = $global:userValue -split '\s+', 2
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
        $compareButton.PerformClick()
        $e.SuppressKeyPress = $true
    }
})

# Reset button click event
$resetButton.Add_Click({
    Reset-TextFields
})

# Run the form
[void] $form.ShowDialog()
