Add-Type -AssemblyName System.Windows.Forms

# Function to get the latest event log from specified domain controllers
function Get-EventFromDomainControllers {
    param (
        [string]$eventID,
        [string]$username,
        [string[]]$domainControllers
    )
    
    $events = @()

    foreach ($dc in $domainControllers) {
        try {
            $dcEvents = Get-WinEvent -ComputerName $dc -FilterHashtable @{LogName='Security'; ID=$eventID; Data=$username} | Sort-Object TimeCreated -Descending | Select-Object -First 1
            if ($dcEvents) {
                $events += $dcEvents
            }
        } catch {
            Write-Host "Failed to get events from ${dc}: $_"
        }
    }

    return $events | Sort-Object TimeCreated -Descending | Select-Object -First 1
}

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "AD Group Comparison and User Activity"
$form.Size = New-Object System.Drawing.Size(420,750)

# Create a tab control
$mainTabControl = New-Object System.Windows.Forms.TabControl
$mainTabControl.Size = New-Object System.Drawing.Size(400,700)
$mainTabControl.Location = New-Object System.Drawing.Point(10,10)

# Create 'User' tab
$userTab = New-Object System.Windows.Forms.TabPage
$userTab.Text = "User"

# Create 'Computer' tab
$computerTab = New-Object System.Windows.Forms.TabPage
$computerTab.Text = "Computer"

# Add main tabs to tab control
$mainTabControl.TabPages.Add($userTab)
$mainTabControl.TabPages.Add($computerTab)
$form.Controls.Add($mainTabControl)

# ------------------------------
# Create sub-tabs for 'User' tab
# ------------------------------
$userTabControl = New-Object System.Windows.Forms.TabControl
$userTabControl.Size = New-Object System.Drawing.Size(380,660)
$userTabControl.Location = New-Object System.Drawing.Point(10,10)
$userTab.Controls.Add($userTabControl)

# Create 'AD Groups' sub-tab
$adGroupsTab = New-Object System.Windows.Forms.TabPage
$adGroupsTab.Text = "AD Groups"
$userTabControl.TabPages.Add($adGroupsTab)

# ------------------------------
# AD Groups Sub-Tab Content
# ------------------------------
# Create a label for username (AD Groups tab)
$adGroupsUserLabel = New-Object System.Windows.Forms.Label
$adGroupsUserLabel.Text = "Enter Username:"
$adGroupsUserLabel.AutoSize = $true
$adGroupsUserLabel.Location = New-Object System.Drawing.Point(10,10)
$adGroupsTab.Controls.Add($adGroupsUserLabel)

# Create a textbox for username (AD Groups tab)
$adGroupsUserTextbox = New-Object System.Windows.Forms.TextBox
$adGroupsUserTextbox.Size = New-Object System.Drawing.Size(200,20)
$adGroupsUserTextbox.Location = New-Object System.Drawing.Point(10,30)
$adGroupsTab.Controls.Add($adGroupsUserTextbox)

# Create a button to run the comparison (AD Groups tab)
$adGroupsButton = New-Object System.Windows.Forms.Button
$adGroupsButton.Text = "Compare Groups"
$adGroupsButton.Size = New-Object System.Drawing.Size(150,30)
$adGroupsButton.Location = New-Object System.Drawing.Point(220,30)
$adGroupsTab.Controls.Add($adGroupsButton)

# Create a label for current user AD groups (AD Groups tab)
$adGroupsCurrentGroupsLabel = New-Object System.Windows.Forms.Label
$adGroupsCurrentGroupsLabel.Text = "Current User AD Groups:"
$adGroupsCurrentGroupsLabel.AutoSize = $true
$adGroupsCurrentGroupsLabel.Location = New-Object System.Drawing.Point(10,60)
$adGroupsTab.Controls.Add($adGroupsCurrentGroupsLabel)

# Create a textbox for current user AD groups with scroll bars (AD Groups tab)
$adGroupsCurrentGroupsTextbox = New-Object System.Windows.Forms.TextBox
$adGroupsCurrentGroupsTextbox.Multiline = $true
$adGroupsCurrentGroupsTextbox.Size = New-Object System.Drawing.Size(360,100)
$adGroupsCurrentGroupsTextbox.ScrollBars = "Vertical"
$adGroupsCurrentGroupsTextbox.Location = New-Object System.Drawing.Point(10,80)
$adGroupsTab.Controls.Add($adGroupsCurrentGroupsTextbox)

# Create a label for AD groups (AD Groups tab)
$adGroupsLabel = New-Object System.Windows.Forms.Label
$adGroupsLabel.Text = "Comparison Groups:"
$adGroupsLabel.AutoSize = $true
$adGroupsLabel.Location = New-Object System.Drawing.Point(10,190)
$adGroupsTab.Controls.Add($adGroupsLabel)

# Create a textbox for AD groups with scroll bars (AD Groups tab)
$adGroupsTextbox = New-Object System.Windows.Forms.TextBox
$adGroupsTextbox.Multiline = $true
$adGroupsTextbox.Size = New-Object System.Drawing.Size(360,200)
$adGroupsTextbox.ScrollBars = "Vertical"
$adGroupsTextbox.Location = New-Object System.Drawing.Point(10,210)
$adGroupsTab.Controls.Add($adGroupsTextbox)

# Create a label for missing groups (AD Groups tab)
$adGroupsMissingGroupsLabel = New-Object System.Windows.Forms.Label
$adGroupsMissingGroupsLabel.Text = "Missing Groups:"
$adGroupsMissingGroupsLabel.AutoSize = $true
$adGroupsMissingGroupsLabel.Location = New-Object System.Drawing.Point(10,420)
$adGroupsTab.Controls.Add($adGroupsMissingGroupsLabel)

# Create a textbox for missing groups with scroll bars (AD Groups tab)
$adGroupsMissingGroupsTextbox = New-Object System.Windows.Forms.TextBox
$adGroupsMissingGroupsTextbox.Multiline = $true
$adGroupsMissingGroupsTextbox.Size = New-Object System.Drawing.Size(360,100)
$adGroupsMissingGroupsTextbox.ScrollBars = "Vertical"
$adGroupsMissingGroupsTextbox.Location = New-Object System.Drawing.Point(10,440)
$adGroupsTab.Controls.Add($adGroupsMissingGroupsTextbox)

# Add button click event (AD Groups tab)
$adGroupsButton.Add_Click({
    # Split input into an array of group names
    $desiredGroups = $adGroupsTextbox.Text -split "`r`n"
    
    # Get the username from the text box
    $username = $adGroupsUserTextbox.Text

    # Refresh the user's group membership data
    Clear-Variable -Name userGroups -Scope Global -ErrorAction SilentlyContinue

    # Get the groups the user belongs to
    $userGroups = (Get-ADUser -Identity $username -Properties MemberOf).MemberOf

    # Extract just the group names from the full CN string
    $userGroups = $userGroups | ForEach-Object { $_ -replace '^CN=([^,]+).+$','$1' }

    # Output the current user groups in the text box
    $adGroupsCurrentGroupsTextbox.Text = $userGroups -join "`r`n"

    # Compare the groups and find missing ones
    $missingGroups = $desiredGroups | Where-Object { $_ -notin $userGroups }

    # Output the missing groups in the text box
    $adGroupsMissingGroupsTextbox.Text = $missingGroups -join "`r`n"
})

# ------------------------------
# Computer Tab Content
# ------------------------------

# Create a label for username (Computer tab)
$compUserLabel = New-Object System.Windows.Forms.Label
$compUserLabel.Text = "Enter Username:"
$compUserLabel.AutoSize = $true
$compUserLabel.Location = New-Object System.Drawing.Point(10,10)
$computerTab.Controls.Add($compUserLabel)

# Create a textbox for username (Computer tab)
$compUserTextbox = New-Object System.Windows.Forms.TextBox
$compUserTextbox.Size = New-Object System.Drawing.Size(200,20)
$compUserTextbox.Location = New-Object System.Drawing.Point(10,30)
$computerTab.Controls.Add($compUserTextbox)

# Create a button to run the comparison (Computer tab)
$compButton = New-Object System.Windows.Forms.Button
$compButton.Text = "Compare Groups"
$compButton.Size = New-Object System.Drawing.Size(150,30)
$compButton.Location = New-Object System.Drawing.Point(220,30)
$computerTab.Controls.Add($compButton)

# Create a label for current user AD groups (Computer tab)
$compCurrentGroupsLabel = New-Object System.Windows.Forms.Label
$compCurrentGroupsLabel.Text = "Current User AD Groups:"
$compCurrentGroupsLabel.AutoSize = $true
$compCurrentGroupsLabel.Location = New-Object System.Drawing.Point(10,60)
$computerTab.Controls.Add($compCurrentGroupsLabel)

# Create a textbox for current user AD groups with scroll bars (Computer tab)
$compCurrentGroupsTextbox = New-Object System.Windows.Forms.TextBox
$compCurrentGroupsTextbox.Multiline = $true
$compCurrentGroupsTextbox.Size = New-Object System.Drawing.Size(360,100)
$compCurrentGroupsTextbox.ScrollBars = "Vertical"
$compCurrentGroupsTextbox.Location = New-Object System.Drawing.Point(10,80)
$computerTab.Controls.Add($compCurrentGroupsTextbox)

# Create a label for AD groups (Computer tab)
$compLabel = New-Object System.Windows.Forms.Label
$compLabel.Text = "Comparison Groups:"
$compLabel.AutoSize = $true
$compLabel.Location = New-Object System.Drawing.Point(10,190)
$computerTab.Controls.Add($compLabel)

# Create a textbox for AD groups with scroll bars (Computer tab)
$compTextbox = New-Object System.Windows.Forms.TextBox
$compTextbox.Multiline = $true
$compTextbox.Size = New-Object System.Drawing.Size(360,200)
$compTextbox.ScrollBars = "Vertical"
$compTextbox.Location = New-Object System.Drawing.Point(10,210)
$computerTab.Controls.Add($compTextbox)

# Create a label for missing groups (Computer tab)
$compMissingGroupsLabel = New-Object System.Windows.Forms.Label
$compMissingGroupsLabel.Text = "Missing Groups:"
$compMissingGroupsLabel.AutoSize = $true
$compMissingGroupsLabel.Location = New-Object System.Drawing.Point(10,420)
$computerTab.Controls.Add($compMissingGroupsLabel)

# Create a textbox for missing groups with scroll bars (Computer tab)
$compMissingGroupsTextbox = New-Object System.Windows.Forms.TextBox
$compMissingGroupsTextbox.Multiline = $true
$compMissingGroupsTextbox.Size = New-Object System.Drawing.Size(360,100)
$compMissingGroupsTextbox.ScrollBars = "Vertical"
$compMissingGroupsTextbox.Location = New-Object System.Drawing.Point(10,440)
$computerTab.Controls.Add($compMissingGroupsTextbox)

# Add button click event (Computer tab)
$compButton.Add_Click({
    # Split input into an array of group names
    $desiredGroups = $compTextbox.Text -split "`r`n"
    
    # Get the username from the text box
    $username = $compUserTextbox.Text

    # Refresh the user's group membership data
    Clear-Variable -Name userGroups -Scope Global -ErrorAction SilentlyContinue

    # Get the groups the user belongs to
    $userGroups = (Get-ADUser -Identity $username -Properties MemberOf).MemberOf

    # Extract just the group names from the full CN string
    $userGroups = $userGroups | ForEach-Object { $_ -replace '^CN=([^,]+).+$','$1' }

    # Output the current user groups in the text box
    $compCurrentGroupsTextbox.Text = $userGroups -join "`r`n"

    # Compare the groups and find missing ones
    $missingGroups = $desiredGroups | Where-Object { $_ -notin $userGroups }

    # Output the missing groups in the text box
    $compMissingGroupsTextbox.Text = $missingGroups -join "`r`n"
})

# Run the form
[void] $form.ShowDialog()
