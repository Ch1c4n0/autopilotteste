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

# Criar a interface gráfica
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "Microsoft Graph Login"
$form.Size = New-Object System.Drawing.Size(400, 300)

$buttonLogin = New-Object System.Windows.Forms.Button
$buttonLogin.Text = "Login"
$buttonLogin.Location = New-Object System.Drawing.Point(150, 50)
$buttonLogin.Size = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($buttonLogin)

$buttonAutopilotGroupTag = New-Object System.Windows.Forms.Button
$buttonAutopilotGroupTag.Text = "Autopilot GroupTag"
$buttonAutopilotGroupTag.Location = New-Object System.Drawing.Point(150, 100)
$buttonAutopilotGroupTag.Size = New-Object System.Drawing.Size(150, 30)
$buttonAutopilotGroupTag.Enabled = $false
$form.Controls.Add($buttonAutopilotGroupTag)

$textBoxProfiles = New-Object System.Windows.Forms.TextBox
$textBoxProfiles.Multiline = $true
$textBoxProfiles.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$textBoxProfiles.Location = New-Object System.Drawing.Point(50, 150)
$textBoxProfiles.Size = New-Object System.Drawing.Size(300, 100)
$form.Controls.Add($textBoxProfiles)

# Evento de clique do botão de login
$buttonLogin.Add_Click({
    Connect-MgGraphWithScopes
    $buttonAutopilotGroupTag.Enabled = $true
    [System.Windows.Forms.MessageBox]::Show("Login successful!")
})

# Evento de clique do botão Autopilot GroupTag
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

# Exibir o formulário
[void]$form.ShowDialog()
