# Função para conectar ao Microsoft Graph
function Connect-MgGraphWithScopes {
    Connect-MgGraph -Scopes "User.Read.All", "Group.ReadWrite.All", "Directory.ReadWrite.All", "DeviceManagementManagedDevices.ReadWrite.All", "DeviceManagementConfiguration.ReadWrite.All", "DeviceManagementServiceConfig.ReadWrite.All", "DeviceManagementApps.ReadWrite.All"
}

# Função para obter perfis de implantação do Windows Autopilot
function Get-AutopilotProfiles {
    $profiles = Get-MgBetaDeviceManagementWindowsAutopilotDeploymentProfile
    return $profiles
}

# Função para obter informações do Windows Autopilot com GroupTag
function Get-WindowsAutopilotInfoWithGroupTag($groupTag) {
    Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Force
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Confirm:$false -Force:$true
    Install-Script get-windowsautopilotinfo -Confirm:$false -Force:$true
    get-windowsautopilotinfo -Online -GroupTag $groupTag
}

# Função para obter informações do Windows Autopilot e salvar em CSV
function Get-WindowsAutopilotInfoCSV {
    Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Force
    Install-Script -Name Get-WindowsAutoPilotInfo -Force
    $serialNumber = (Get-WmiObject -Class Win32_BIOS).SerialNumber
    $outputFile = "C:\$serialNumber.csv"
    Get-WindowsAutoPilotInfo.ps1 -OutputFile $outputFile
    return $outputFile
}

# Criar a interface gráfica
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "Autopilot - Marcelo Goncalves v1.0 - Intune Lovers!"
$form.Size = New-Object System.Drawing.Size(400, 400)

$buttonInstallGraph = New-Object System.Windows.Forms.Button
$buttonInstallGraph.Text = "Install Microsoft Graph"
$buttonInstallGraph.Location = New-Object System.Drawing.Point(100, 20)
$buttonInstallGraph.Size = New-Object System.Drawing.Size(200, 30)
$buttonInstallGraph.BackColor = [System.Drawing.Color]::Yellow
$form.Controls.Add($buttonInstallGraph)

$buttonLogin = New-Object System.Windows.Forms.Button
$buttonLogin.Text = "Login"
$buttonLogin.Location = New-Object System.Drawing.Point(150, 70)
$buttonLogin.Size = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($buttonLogin)

$buttonAutopilotGroupTag = New-Object System.Windows.Forms.Button
$buttonAutopilotGroupTag.Text = "Autopilot Online With Group Tag"
$buttonAutopilotGroupTag.Location = New-Object System.Drawing.Point(100, 120)
$buttonAutopilotGroupTag.Size = New-Object System.Drawing.Size(200, 30)
$buttonAutopilotGroupTag.Enabled = $false
$form.Controls.Add($buttonAutopilotGroupTag)

$buttonAutopilotCSV = New-Object System.Windows.Forms.Button
$buttonAutopilotCSV.Text = "Windows Autopilot CSV"
$buttonAutopilotCSV.Location = New-Object System.Drawing.Point(100, 170)
$buttonAutopilotCSV.Size = New-Object System.Drawing.Size(200, 30)
$buttonAutopilotCSV.Enabled = $false
$form.Controls.Add($buttonAutopilotCSV)

$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Location = New-Object System.Drawing.Point(150, 50)
$labelStatus.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($labelStatus)

# Evento de clique do botão Install Microsoft Graph
$buttonInstallGraph.Add_Click({
    $installForm = New-Object System.Windows.Forms.Form
    $installForm.Text = "Installing Microsoft Graph"
    $installForm.Size = New-Object System.Drawing.Size(300, 100)

    $labelInstall = New-Object System.Windows.Forms.Label
    $labelInstall.Text = "Installing Microsoft Graph, please wait..."
    $labelInstall.AutoSize = $true
    $labelInstall.Location = New-Object System.Drawing.Point(10, 10)
    $installForm.Controls.Add($labelInstall)

    $installForm.Show()

    Start-Process powershell -ArgumentList "Install-Module Microsoft.Graph -AllowClobber -Force" -Wait

    $installForm.Close()
    [System.Windows.Forms.MessageBox]::Show("Microsoft Graph installed successfully!", "Installation Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

# Evento de clique do botão de login
$buttonLogin.Add_Click({
    try {
        Connect-MgGraphWithScopes
        $buttonAutopilotGroupTag.Enabled = $true
        $buttonAutopilotCSV.Enabled = $true
        $successForm = New-Object System.Windows.Forms.Form
        $successForm.Text = "Login Status"
        $successForm.Size = New-Object System.Drawing.Size(300, 100)
        $successForm.BackColor = [System.Drawing.Color]::Green

        $labelSuccess = New-Object System.Windows.Forms.Label
        $labelSuccess.Text = "SUCCESS"
        $labelSuccess.AutoSize = $true
        $labelSuccess.Location = New-Object System.Drawing.Point(10, 10)
        $successForm.Controls.Add($labelSuccess)

        $successForm.ShowDialog()
    } catch {
        $failForm = New-Object System.Windows.Forms.Form
        $failForm.Text = "Login Status"
        $failForm.Size = New-Object System.Drawing.Size(300, 100)
        $failForm.BackColor = [System.Drawing.Color]::Red

        $labelFail = New-Object System.Windows.Forms.Label
        $labelFail.Text = "FAIL"
        $labelFail.AutoSize = $true
        $labelFail.Location = New-Object System.Drawing.Point(10, 10)
        $failForm.Controls.Add($labelFail)

        $failForm.ShowDialog()
    }
})

# Evento de clique do botão Autopilot Online With Group Tag
$buttonAutopilotGroupTag.Add_Click({
    $inputForm = New-Object System.Windows.Forms.Form
    $inputForm.Text = "Enter Group Tag"
    $inputForm.Size = New-Object System.Drawing.Size(300, 150)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Group Tag:"
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $inputForm.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(80, 20)
    $textBox.Size = New-Object System.Drawing.Size(200, 20)
    $textBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $inputForm.Controls.Add($textBox)

    $buttonOK = New-Object System.Windows.Forms.Button
    $buttonOK.Text = "OK"
    $buttonOK.Location = New-Object System.Drawing.Point(100, 60)
    $buttonOK.Size = New-Object System.Drawing.Size(75, 30)
    $inputForm.Controls.Add($buttonOK)

    $buttonOK.Add_Click({
        $groupTag = $textBox.Text
        $inputForm.Close()
        $result = Get-WindowsAutopilotInfoWithGroupTag $groupTag
        $textBoxProfiles.Text = $result | Out-String
    })

    $inputForm.ShowDialog()
})

# Evento de clique do botão Windows Autopilot CSV
$buttonAutopilotCSV.Add_Click({
    $outputFile = Get-WindowsAutopilotInfoCSV
    [System.Windows.Forms.MessageBox]::Show("CSV file created at $outputFile")
})

# Exibir o formulário
[void]$form.ShowDialog()
