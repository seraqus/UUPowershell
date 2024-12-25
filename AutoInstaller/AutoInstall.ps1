Add-Type -AssemblyName "System.Windows.Forms"
Add-Type -AssemblyName "System.Drawing"

# Check if the script is running as Administrator
function Is-Admin {
    $currentIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($currentIdentity)
    return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

# If not running as administrator, re-launch the script as administrator
if (-not (Is-Admin)) {
    $args = [System.String]::Join(' ', $MyInvocation.Line)
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command $args" -Verb RunAs
    Exit
}

# Function to create and show the UI
function Show-UI {
    # Create Form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Folders for Installation"
    $form.Size = New-Object System.Drawing.Size(400, 400)

    # Query the file share for folder names
    $folderPath = "\\itdhq1dsp02\DeploymentShareProd5$\Applications"
    $folders = Get-ChildItem -Path $folderPath -Directory | Select-Object -ExpandProperty Name

    # Create checkbox list
    $checkboxList = New-Object System.Windows.Forms.CheckedListBox
    $checkboxList.Location = New-Object System.Drawing.Point(20, 20)
    $checkboxList.Size = New-Object System.Drawing.Size(340, 200)
    $checkboxList.Items.AddRange($folders)

    # Add a button to execute the batch files
    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = "Run Installations"
    $btnRun.Location = New-Object System.Drawing.Point(20, 250)
    $btnRun.Size = New-Object System.Drawing.Size(100, 30)

    # Add event for button click
    $btnRun.Add_Click({
        $selectedFolders = $checkboxList.CheckedItems
        if ($selectedFolders.Count -gt 0) {
            # Prompt for username and password
            $credentialForm = New-Object System.Windows.Forms.Form
            $credentialForm.Text = "Enter Service Account Credentials"
            $credentialForm.Size = New-Object System.Drawing.Size(300, 150)

            $usernameLabel = New-Object System.Windows.Forms.Label
            $usernameLabel.Text = "Username:"
            $usernameLabel.Location = New-Object System.Drawing.Point(10, 20)
            $usernameLabel.Size = New-Object System.Drawing.Size(80, 20)

            $usernameBox = New-Object System.Windows.Forms.TextBox
            $usernameBox.Location = New-Object System.Drawing.Point(100, 20)
            $usernameBox.Size = New-Object System.Drawing.Size(150, 20)

            $passwordLabel = New-Object System.Windows.Forms.Label
            $passwordLabel.Text = "Password:"
            $passwordLabel.Location = New-Object System.Drawing.Point(10, 60)
            $passwordLabel.Size = New-Object System.Drawing.Size(80, 20)

            $passwordBox = New-Object System.Windows.Forms.TextBox
            $passwordBox.Location = New-Object System.Drawing.Point(100, 60)
            $passwordBox.Size = New-Object System.Drawing.Size(150, 20)
            $passwordBox.UseSystemPasswordChar = $true

            $btnSubmit = New-Object System.Windows.Forms.Button
            $btnSubmit.Text = "Submit"
            $btnSubmit.Location = New-Object System.Drawing.Point(100, 90)
            $btnSubmit.Size = New-Object System.Drawing.Size(75, 30)
            $btnSubmit.Add_Click({
                # Get the entered credentials
                $username = $usernameBox.Text
                $password = $passwordBox.Text

                # Close the credential form
                $credentialForm.Close()

                # Proceed to run batch files with the entered credentials
                Run-BatchFiles -folders $selectedFolders -username $username -password $password
            })

            # Add controls to form
            $credentialForm.Controls.Add($usernameLabel)
            $credentialForm.Controls.Add($usernameBox)
            $credentialForm.Controls.Add($passwordLabel)
            $credentialForm.Controls.Add($passwordBox)
            $credentialForm.Controls.Add($btnSubmit)

            # Show the credential form
            $credentialForm.ShowDialog()
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one folder.")
        }
    })

    # Add controls to form
    $form.Controls.Add($checkboxList)
    $form.Controls.Add($btnRun)

    # Show Form
    $form.ShowDialog()
}

# Function to run the batch files and handle reboots
function Run-BatchFiles {
    param(
        [Parameter(Mandatory=$true)]
        [System.Collections.ArrayList]$folders,
        
        [string]$username,
        [string]$password
    )

    # Count batch files to determine how many reboots are needed
    $totalBatchFiles = 0
    foreach ($folder in $folders) {
        $batchFiles = Get-ChildItem -Path "\\server\location\etc\$folder" -Filter "*.bat"
        $totalBatchFiles += $batchFiles.Count
    }

    # Set AutoLogin to run the user in next boot, based on the total number of batch files
    Set-AutoLogin -username $username -password $password -reboots $totalBatchFiles

    # Execute the batch files
    foreach ($folder in $folders) {
        $batchFiles = Get-ChildItem -Path "\\server\location\etc\$folder" -Filter "*.bat"
        foreach ($batchFile in $batchFiles) {
            # Run the batch file as Administrator
            Write-Host "Running $batchFile as Administrator"
            Start-Process -FilePath $batchFile.FullName -Verb RunAs -Wait

            # Restart PC after each install
            [System.Windows.Forms.MessageBox]::Show("Installation complete for $folder. Restarting PC.")
            Restart-Computer -Force
        }
    }

    # Clean up the auto-login settings after all installations are complete
    Clear-AutoLogin

    # Completion message after everything is done
    [System.Windows.Forms.MessageBox]::Show("All installations completed and system is ready for use.")
}

# Function to set auto login on the system, with domain user support
function Set-AutoLogin {
    param (
        [string]$username,   # Expecting the format "domain\username"
        [string]$password,
        [int]$reboots
    )

    $regKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"

    # Ensure the username is in the "domain\username" format
    if ($username -notmatch '\\') {
        Write-Host "The username must be in the format 'domain\username'."
        return
    }

    # Set AutoLogin for domain user
    Set-ItemProperty -Path $regKey -Name "AutoAdminLogon" -Value "1"
    Set-ItemProperty -Path $regKey -Name "DefaultUserName" -Value $username
    Set-ItemProperty -Path $regKey -Name "DefaultPassword" -Value $password
    Set-ItemProperty -Path $regKey -Name "AutoLogonCount" -Value $reboots

    Write-Host "Auto-login configured for domain user $username."
}


# Function to clear auto login settings
function Clear-AutoLogin {
    $regKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
    Remove-ItemProperty -Path $regKey -Name "AutoAdminLogon" -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path $regKey -Name "DefaultUserName" -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path $regKey -Name "DefaultPassword" -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path $regKey -Name "AutoLogonCount" -ErrorAction SilentlyContinue
}

# Run the UI to start the process
Show-UI




## Does not execute batch files in a row at the moment ##