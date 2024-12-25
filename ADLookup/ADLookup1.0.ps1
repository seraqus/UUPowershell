Add-Type -AssemblyName System.Windows.Forms

# Main Window
$form = New-Object System.Windows.Forms.Form
$form.Text = "AD Group Comparison"
$form.Size = New-Object System.Drawing.Size(400,500)

# Username label
$userLabel = New-Object System.Windows.Forms.Label
$userLabel.Text = "Enter Username:"
$userLabel.AutoSize = $true
$userLabel.Location = New-Object System.Drawing.Point(10,10)
$form.Controls.Add($userLabel)

# Username textbox
$userTextbox = New-Object System.Windows.Forms.TextBox
$userTextbox.Size = New-Object System.Drawing.Size(200,20)
$userTextbox.Location = New-Object System.Drawing.Point(10,30)
$form.Controls.Add($userTextbox)

# Check button
$button = New-Object System.Windows.Forms.Button
$button.Text = "Check Groups"
$button.Size = New-Object System.Drawing.Size(150,30)  # Adjusted size
$button.Location = New-Object System.Drawing.Point(220,25)
$form.Controls.Add($button)

# Current Groups label
$currentGroupsLabel = New-Object System.Windows.Forms.Label
$currentGroupsLabel.Text = "Current Groups:"
$currentGroupsLabel.AutoSize = $true
$currentGroupsLabel.Location = New-Object System.Drawing.Point(10,60)
$form.Controls.Add($currentGroupsLabel)

# Current Groups textbox
$currentGroupsTextbox = New-Object System.Windows.Forms.TextBox
$currentGroupsTextbox.Multiline = $true
$currentGroupsTextbox.Size = New-Object System.Drawing.Size(360,100)
$currentGroupsTextbox.ScrollBars = "Vertical"
$currentGroupsTextbox.Location = New-Object System.Drawing.Point(10,80)
$form.Controls.Add($currentGroupsTextbox)

# Comparison Groups label
$label = New-Object System.Windows.Forms.Label
$label.Text = "Comparison Groups:"
$label.AutoSize = $true
$label.Location = New-Object System.Drawing.Point(10,190)
$form.Controls.Add($label)

# Comparison Groups textbox
$textbox = New-Object System.Windows.Forms.TextBox
$textbox.Multiline = $true
$textbox.Size = New-Object System.Drawing.Size(360,100)
$textbox.ScrollBars = "Vertical"
$textbox.Location = New-Object System.Drawing.Point(10,210)
$form.Controls.Add($textbox)

# Missing Groups label
$missingGroupsLabel = New-Object System.Windows.Forms.Label
$missingGroupsLabel.Text = "Missing Groups:"
$missingGroupsLabel.AutoSize = $true
$missingGroupsLabel.Location = New-Object System.Drawing.Point(10,320)
$form.Controls.Add($missingGroupsLabel)

# Missing Groups textbox
$missingGroupsTextbox = New-Object System.Windows.Forms.TextBox
$missingGroupsTextbox.Multiline = $true
$missingGroupsTextbox.Size = New-Object System.Drawing.Size(360,100)
$missingGroupsTextbox.ScrollBars = "Vertical"
$missingGroupsTextbox.Location = New-Object System.Drawing.Point(10,340)
$form.Controls.Add($missingGroupsTextbox)

# Check groups function
$button.Add_Click({
    # Split input into an array of group names
    $desiredGroups = $textbox.Text -split "`r`n"
    
    # Get the username from the text box
    $username = $userTextbox.Text

    # Refresh the user's group membership data
    Clear-Variable -Name userGroups -Scope Global -ErrorAction SilentlyContinue

    # Get the groups the user belongs to
    $userGroups = (Get-ADUser -Identity $username -Properties MemberOf).MemberOf

    # Extract just the group names from the full CN string
    $userGroups = $userGroups | ForEach-Object { $_ -replace '^CN=([^,]+).+$','$1' }

    # Output the current user groups in the text box
    $currentGroupsTextbox.Text = $userGroups -join "`r`n"

    # Compare the groups and find missing ones
    $missingGroups = $desiredGroups | Where-Object { $_ -notin $userGroups }

    # Output the missing groups in the text box
    $missingGroupsTextbox.Text = $missingGroups -join "`r`n"
})

# Run the form
[void] $form.ShowDialog()
