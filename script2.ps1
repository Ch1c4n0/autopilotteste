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

# Função para enviar arquivo por e-mail
function Send-EmailWithAttachment($filePath, $smtpServer, $smtpPort, $smtpUser, $smtpPassword, $to, $subject, $body) {
    $securePassword = ConvertTo-SecureString $smtpPassword -AsPlainText -Force
    $credentials = New-Object System.Management.Automation.PSCredential ($smtpUser, $securePassword)
    $message = New-Object System.Net.Mail.MailMessage
    $message.From = $smtpUser
    $message.To.Add($to)
    $message.Subject = $subject
    $message.Body = $body
    $attachment = New-Object System.Net.Mail.Attachment($filePath)
    $message.Attachments.Add($attachment)
    $smtp = New-Object System.Net.Mail.SmtpClient($smtpServer, $smtpPort)
    $smtp.EnableSsl = $true
    $smtp.Credentials = $credentials
    $smtp.Send($message)
}

# Criar a interface gráfica
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "Autopilot"
$form.Size = New-Object System.Drawing.Size(400, 500)

$buttonInstallGraph = New-Object System.Windows.Forms.Button
$buttonInstallGraph.Text = "Install Microsoft Graph"
$buttonInstallGraph.Location = New-Object System.Drawing.Point(100, 20)
$buttonInstallGraph.Size = New-Object System.Drawing.Size(200, 30)
$buttonInstallGraph.BackColor = [System.Drawing.Color]::Yellow
$form.Controls.Add($buttonInstallGraph)

$buttonInstallPnP = New-Object System.Windows.Forms.Button
$buttonInstallPnP.Text = "PNP Sharepoint"
$buttonInstallPnP.Location = New-Object System.Drawing.Point(100, 70)
$buttonInstallPnP.Size = New-Object System.Drawing.Size(200, 30)
$buttonInstallPnP.BackColor = [System.Drawing.Color]::Purple
$form.Controls.Add($buttonInstallPnP)

$buttonLogin = New-Object System.Windows.Forms.Button
$buttonLogin.Text = "Login"
$buttonLogin.Location = New-Object System.Drawing.Point(150, 120)
$buttonLogin.Size = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($buttonLogin)

$buttonAutopilotGroupTag = New-Object System.Windows.Forms.Button
$buttonAutopilotGroupTag.Text = "Autopilot Online With Group Tag"
$buttonAutopilotGroupTag.Location = New-Object System.Drawing.Point(100, 170)
$buttonAutopilotGroupTag.Size = New-Object System.Drawing.Size(200, 30)
$buttonAutopilotGroupTag.Enabled = $false
$form.Controls.Add($buttonAutopilotGroupTag)

$buttonAutopilotCSV = New-Object System.Windows.Forms.Button
$buttonAutopilotCSV.Text = "Windows Autopilot CSV"
$buttonAutopilotCSV.Location = New-Object System.Drawing.Point(100, 220)
$buttonAutopilotCSV.Size = New-Object System.Drawing.Size(200, 30)
$buttonAutopilotCSV.Enabled = $false
$form.Controls.Add($buttonAutopilotCSV)

$buttonAutopilotEmail = New-Object System.Windows.Forms.Button
$buttonAutopilotEmail.Text = "Windows Autopilot E-mail"
$buttonAutopilotEmail.Location = New-Object System.Drawing.Point(100, 270)
$buttonAutopilotEmail.Size = New-Object System.Drawing.Size(200, 30)
$buttonAutopilotEmail.Enabled = $false
$form.Controls.Add($buttonAutopilotEmail)

$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Location = New-Object System.Drawing.Point(150, 150)
$labelStatus.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($labelStatus)

# Evento de clique do botão Install Microsoft Graph
$buttonInstallGraph.Add_Click({
    Start-Process powershell -ArgumentList "Install-Module Microsoft.Graph -AllowClobber -Force" -NoNewWindow
})

# Evento de clique do botão PNP Sharepoint
$buttonInstallPnP.Add_Click({
    Start-Process powershell -ArgumentList "Install-Module -Name PnP.PowerShell -Force" -NoNewWindow
})

# Evento de clique do botão de login
$buttonLogin.Add_Click({
    try {
        Connect-MgGraphWithScopes
        $buttonAutopilotGroupTag.Enabled = $true
        $buttonAutopilotCSV.Enabled = $true
        $buttonAutopilotEmail.Enabled = $true
        $labelStatus.Text = "SUCCESS"
        $labelStatus.ForeColor = [System.Drawing.Color]::Green
    } catch {
        $labelStatus.Text = "FAIL"
        $labelStatus.ForeColor = [System.Drawing.Color]::Red
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

# Evento de clique do botão Windows Autopilot E-mail
$buttonAutopilotEmail.Add_Click({
    $outputFile = Get-WindowsAutopilotInfoCSV

    $inputForm = New-Object System.Windows.Forms.Form
    $inputForm.Text = "Enter E-mail Details"
    $inputForm.Size = New-Object System.Drawing.Size(300, 350)

    $labelSmtp = New-Object System.Windows.Forms.Label
    $labelSmtp.Text = "SMTP Server:"
    $labelSmtp.Location = New-Object System.Drawing.Point(10, 20)
    $inputForm.Controls.Add($labelSmtp)

    $textBoxSmtp = New-Object System.Windows.Forms.TextBox
    $textBoxSmtp.Location = New-Object System.Drawing.Point(100, 20)
    $textBoxSmtp.Size = New-Object System.Drawing.Size(180, 20)
    $textBoxSmtp.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $inputForm.Controls.Add($textBoxSmtp)

    $labelPort = New-Object System.Windows.Forms.Label
    $labelPort.Text = "SMTP Port:"
    $labelPort.Location = New-Object System.Drawing.Point(10, 60)
    $inputForm.Controls.Add($labelPort)

    $textBoxPort = New-Object System.Windows.Forms.TextBox
    $textBoxPort.Location = New-Object System.Drawing.Point(100, 60)
    $textBoxPort.Size = New-Object System.Drawing.Size(180, 20)
    $textBoxPort.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $inputForm.Controls.Add($textBoxPort)

    $labelUsername = New-Object System.Windows.Forms.Label
    $labelUsername.Text = "Username:"
    $labelUsername.Location = New-Object System.Drawing.Point(10, 100)
    $inputForm.Controls.Add($labelUsername)

    $textBoxUsername = New-Object System.Windows.Forms.TextBox
    $textBoxUsername.Location = New-Object System.Drawing.Point(100, 100)
    $textBoxUsername.Size = New-Object System.Drawing.Size(180, 20)
    $textBoxUsername.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $inputForm.Controls.Add($textBoxUsername)

    $labelPassword = New-Object System.Windows.Forms.Label
    $labelPassword.Text = "Password:"
    $labelPassword.Location = New-Object System.Drawing.Point(10, 140)
    $inputForm.Controls.Add($labelPassword)

    $textBoxPassword = New-Object System.Windows.Forms.TextBox
    $textBoxPassword.Location = New-Object System.Drawing.Point(100, 140)
    $textBoxPassword.Size = New-Object System.Drawing.Size(180, 20)
    $textBoxPassword.PasswordChar = '*'
    $textBoxPassword.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $inputForm.Controls.Add($textBoxPassword)

    $labelTo = New-Object System.Windows.Forms.Label
    $labelTo.Text = "To:"
    $labelTo.Location = New-Object System.Drawing.Point(10, 180)
    $inputForm.Controls.Add($labelTo)

    $textBoxTo = New-Object System.Windows.Forms.TextBox
    $textBoxTo.Location = New-Object System.Drawing.Point(100, 180)
    $textBoxTo.Size = New-Object System.Drawing.Size(180, 20)
    $textBoxTo.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $inputForm.Controls.Add($textBoxTo)

    $labelSubject = New-Object System.Windows.Forms.Label
    $labelSubject.Text = "Subject:"
    $labelSubject.Location = New-Object System.Drawing.Point(10, 220)
    $inputForm.Controls.Add($labelSubject)

    $textBoxSubject = New-Object System.Windows.Forms.TextBox
    $textBoxSubject.Location = New-Object System.Drawing.Point(100, 220)
    $textBoxSubject.Size = New-Object System.Drawing.Size(180, 20)
    $textBoxSubject.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $inputForm.Controls.Add($textBoxSubject)

    $labelBody = New-Object System.Windows.Forms.Label
    $labelBody.Text = "Body:"
    $labelBody.Location = New-Object System.Windows.Forms.Label
    $labelBody.Location = New-Object System.Drawing.Point(10, 260)
    $inputForm.Controls.Add($labelBody)

    $textBoxBody = New-Object System.Windows.Forms.TextBox
    $textBoxBody.Location = New-Object System.Drawing.Point(100, 260)
    $textBoxBody.Size = New-Object System.Drawing.Size(180, 20)
    $textBoxBody.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $inputForm.Controls.Add($textBoxBody)

    $buttonOK = New-Object System.Windows.Forms.Button
    $buttonOK.Text = "OK"
    $buttonOK.Location = New-Object System.Drawing.Point(100, 300)
    $buttonOK.Size = New-Object System.Drawing.Size(75, 30)
    $inputForm.Controls.Add($buttonOK)

    $buttonOK.Add_Click({
        $smtpServer = $textBoxSmtp.Text
        $smtpPort = $textBoxPort.Text
        $username = $textBoxUsername.Text
        $password = $textBoxPassword.Text
        $to = $textBoxTo.Text
        $subject = $textBoxSubject.Text
        $body = $textBoxBody.Text
        $inputForm.Close()
        Send-EmailWithAttachment $outputFile $smtpServer $smtpPort $username $password $to $subject $body
        [System.Windows.Forms.MessageBox]::Show("CSV file sent to $to")
    })

    $inputForm.ShowDialog()
})

# Exibir o formulário
[void]$form.ShowDialog()
