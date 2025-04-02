# Ensure the required module is imported
Import-Module ActiveDirectory

# Define the path to the configuration file
$configPath = "$env:LOCALAPPDATA\ADLookup\Config.txt"

# Check if the configuration file exists
if (-not (Test-Path $configPath)) {
    Write-Host "Configuration file not found at $configPath. Creating a default configuration file..." -ForegroundColor Yellow

    # Create the directory if it doesn't exist
    $configDir = Split-Path $configPath
    if (-not (Test-Path $configDir)) {
        New-Item -ItemType Directory -Path $configDir -Force | Out-Null
    }

    # Define default configuration values
    $defaultConfig = @"
ServiceNowInstance: https://your-instance.service-now.com
Disabled Accounts location: OU=DisabledAccounts,DC=yourdomain,DC=com
"@

    # Write the default configuration to the file
    $defaultConfig | Set-Content -Path $configPath
    Write-Host "Default configuration file created at $configPath. Please update it with the necessary details." -ForegroundColor Green
    exit
}

# Load configuration details
$config = Get-Content $configPath | ForEach-Object {
    $key, $value = $_ -split ":", 2
    @{ Key = $key.Trim(); Value = $value.Trim() }
} | Group-Object -AsHashTable -Property Key

# Validate required configuration keys
if (-not $config["ServiceNowInstance"] -or -not $config["Disabled Accounts location"]) {
    Write-Host "Configuration file is missing required keys. Please ensure it includes ServiceNowInstance and Disabled Accounts location." -ForegroundColor Red
    exit
}

# Create the form for user input
Add-Type -AssemblyName System.Windows.Forms
$form = New-Object System.Windows.Forms.Form
$form.Text = "Terminate AD Account"
$form.Size = New-Object System.Drawing.Size(600, 600)
$form.StartPosition = "CenterScreen"

# Add input fields
$labels = @(
    "First Name", "Last Name", "Username", "Legal Hold (Yes/No)", "Ticket Number", 
    "Asset Number(s)", "Department", "Date of Termination (MM/DD/YYYY)", 
    "Email to Forward To", "Cell Phone"
)
$inputs = @{ }
$yPos = 20

foreach ($label in $labels) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $label
    $lbl.Location = New-Object System.Drawing.Point(10, $yPos)
    $lbl.Size = New-Object System.Drawing.Size(200, 20)
    $form.Controls.Add($lbl)

    if ($label -eq "Legal Hold (Yes/No)") {
        $dropdown = New-Object System.Windows.Forms.ComboBox
        $dropdown.Items.AddRange(@("Yes", "No"))
        $dropdown.Location = New-Object System.Drawing.Point(220, $yPos)
        $dropdown.Size = New-Object System.Drawing.Size(200, 20)
        $form.Controls.Add($dropdown)
        $inputs[$label] = $dropdown
    } else {
        $txt = New-Object System.Windows.Forms.TextBox
        $txt.Location = New-Object System.Drawing.Point(220, $yPos)
        $txt.Size = New-Object System.Drawing.Size(200, 20)
        $form.Controls.Add($txt)
        $inputs[$label] = $txt
    }

    $yPos += 30
}

# Add a search button for auto-populating fields
$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Text = "Search"
$searchButton.Location = New-Object System.Drawing.Point(220, $yPos)
$form.Controls.Add($searchButton)

# Add a dropdown for multiple results
$resultsDropdown = New-Object System.Windows.Forms.ComboBox
$resultsDropdown.Location = New-Object System.Drawing.Point(220, $yPos + 40)
$resultsDropdown.Size = New-Object System.Drawing.Size(200, 20)
$resultsDropdown.Visible = $false
$form.Controls.Add($resultsDropdown)

# Add a submit button
$submitButton = New-Object System.Windows.Forms.Button
$submitButton.Text = "Submit"
$submitButton.Location = New-Object System.Drawing.Point(150, $yPos + 80)
$form.Controls.Add($submitButton)

# Define the search button click event
$searchButton.Add_Click({
    $username = $inputs["Username"].Text
    if (-not $username) {
        Write-Host "Please enter a username to search." -ForegroundColor Red
        return
    }

    # Search for the user in AD
    $users = Get-ADUser -Filter "SamAccountName -like '*$username*'" -Properties DisplayName, EmailAddress, GivenName, Surname, Department, MobilePhone

    if ($users.Count -eq 0) {
        Write-Host "No users found matching the username." -ForegroundColor Red
        return
    } elseif ($users.Count -eq 1) {
        # Auto-populate fields if only one user is found
        $user = $users[0]
        $inputs["First Name"].Text = $user.GivenName
        $inputs["Last Name"].Text = $user.Surname
        $inputs["Email to Forward To"].Text = $user.EmailAddress
        $inputs["Department"].Text = $user.Department
        $inputs["Cell Phone"].Text = $user.MobilePhone
    } else {
        # Populate the dropdown with multiple results
        $resultsDropdown.Items.Clear()
        foreach ($user in $users) {
            $resultsDropdown.Items.Add("$($user.SamAccountName) - $($user.DisplayName)")
        }
        $resultsDropdown.Visible = $true
        $resultsDropdown.Add_SelectedIndexChanged({
            $selectedUser = $users[$resultsDropdown.SelectedIndex]
            $inputs["First Name"].Text = $selectedUser.GivenName
            $inputs["Last Name"].Text = $selectedUser.Surname
            $inputs["Email to Forward To"].Text = $selectedUser.EmailAddress
            $inputs["Department"].Text = $selectedUser.Department
            $inputs["Cell Phone"].Text = $selectedUser.MobilePhone
        })
    }
})

# Define the submit button click event
$submitButton.Add_Click({
    $form.Close()
})

# Show the form
$form.ShowDialog()

# Collect input values
$userDetails = @{}
foreach ($key in $inputs.Keys) {
    if ($inputs[$key] -is [System.Windows.Forms.ComboBox]) {
        $userDetails[$key] = $inputs[$key].SelectedItem
    } else {
        $userDetails[$key] = $inputs[$key].Text
    }
}

# Validate inputs
if ($userDetails.Values -contains "") {
    Write-Host "All fields are required. Please fill out the form completely." -ForegroundColor Red
    exit
}

# ServiceNow Integration
# Ensure the ServiceNow module is imported
if (-not (Get-Module -Name ServiceNow -ListAvailable)) {
    Write-Host "ServiceNow module is not installed. Please install it to proceed." -ForegroundColor Red
    exit
}

# Connect to ServiceNow using SSO
$ServiceNowInstance = $config["ServiceNowInstance"]

try {
    $ServiceNowSession = Connect-ServiceNow -Instance $ServiceNowInstance -UseSSO
    Write-Host "Connected to ServiceNow successfully using SSO." -ForegroundColor Green
} catch {
    Write-Host "Failed to connect to ServiceNow using SSO. Please ensure your SSO configuration is correct." -ForegroundColor Red
    exit
}

# Create a ServiceNow task for Disable Account
$ticketNumber = $userDetails["Ticket Number"]
$terminationDate = $userDetails["Date of Termination (MM/DD/YYYY)"]
$disabledOU = $config["Disabled Accounts location"]

$disableAccountTaskDetails = @{
    short_description = "SERVICE CENTER: $ticketNumber--TERM, Please disable AD account for $($userDetails["First Name"]) $($userDetails["Last Name"]) ($($userDetails["Username"])) and move it to $disabledOU, $terminationDate"
    description = "SERVICE CENTER: $ticketNumber--TERM, Please disable AD account for $($userDetails["First Name"]) $($userDetails["Last Name"]) ($($userDetails["Username"])) and move it to $disabledOU, $terminationDate"
    assignment_group = "SN-ITD-Service Center"
    assigned_to = $env:USERNAME
    state = "Open"
}

try {
    $disableAccountTask = New-ServiceNowRecord -Table "task" -Values $disableAccountTaskDetails
    Write-Host "ServiceNow task created successfully for Disable Account. Task ID: $($disableAccountTask.sys_id)" -ForegroundColor Green
} catch {
    Write-Host "Failed to create ServiceNow task for Disable Account. Error: $_" -ForegroundColor Red
    exit
}

# Create a ServiceNow task for Alert Sense
$alertSenseTaskDetails = @{
    short_description = "SERVICE CENTER: $ticketNumber--TERM, Please remove $($userDetails["First Name"]) $($userDetails["Last Name"]) ($($userDetails["Username"])) from Alert Sense, $terminationDate"
    description = "SERVICE CENTER: $ticketNumber--TERM, Please remove $($userDetails["First Name"]) $($userDetails["Last Name"]) ($($userDetails["Username"])) from Alert Sense, $terminationDate"
    assignment_group = "SN-ITD-Service Center"
    assigned_to = "Arthur Mora"
    state = "Open"
}

try {
    $alertSenseTask = New-ServiceNowRecord -Table "task" -Values $alertSenseTaskDetails
    Write-Host "ServiceNow task created successfully for Alert Sense. Task ID: $($alertSenseTask.sys_id)" -ForegroundColor Green
} catch {
    Write-Host "Failed to create ServiceNow task for Alert Sense. Error: $_" -ForegroundColor Red
    exit
}

# Create a ServiceNow task for Disable Local Admin
$disableLocalAdminTaskDetails = @{
    short_description = "SERVICE CENTER: $ticketNumber--TERM, Please remove $($userDetails["First Name"]) $($userDetails["Last Name"]) ($($userDetails["Username"])) as a local administrator, $terminationDate"
    description = "SERVICE CENTER: $ticketNumber--TERM, Please remove $($userDetails["First Name"]) $($userDetails["Last Name"]) ($($userDetails["Username"])) as a local administrator, $terminationDate"
    assignment_group = "SN-ITD-Service Center"
    assigned_to = $env:USERNAME
    state = "Open"
}

try {
    $disableLocalAdminTask = New-ServiceNowRecord -Table "task" -Values $disableLocalAdminTaskDetails
    Write-Host "ServiceNow task created successfully for Disable Local Admin. Task ID: $($disableLocalAdminTask.sys_id)" -ForegroundColor Green
} catch {
    Write-Host "Failed to create ServiceNow task for Disable Local Admin. Error: $_" -ForegroundColor Red
    exit
}