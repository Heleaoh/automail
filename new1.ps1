# Sicherstellen, dass PowerShell UTF-8 verwendet
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Unicode-Zeichen explizit definieren
$oe = [char]0x00F6  # ö
$ue = [char]0x00FC  # ü

# Wort "für" definieren
$fuer = "für"

# Funktion zum Anzeigen einer MessageBox für Benutzereingabe
function Show-InputBoxDialog {
    param (
        [string]$Prompt = "Bitte geben Sie den Namen des Users ein:",
        [string]$Title = "Benutzername eingeben"
    )
    
    Add-Type -AssemblyName System.Windows.Forms
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(300,200)
    $form.StartPosition = "CenterScreen"
    
    $labelName = New-Object System.Windows.Forms.Label
    $labelName.Text = $Prompt
    $labelName.AutoSize = $true
    $labelName.Location = New-Object System.Drawing.Point(50,20)
    $form.Controls.Add($labelName)
    
    $textBoxName = New-Object System.Windows.Forms.TextBox
    $textBoxName.Location = New-Object System.Drawing.Point(50,50)
    $textBoxName.Size = New-Object System.Drawing.Size(200,20)
    $form.Controls.Add($textBoxName)
    
    $labelEmail = New-Object System.Windows.Forms.Label
    $labelEmail.Text = "Bitte geben Sie die E-Mail-Adresse des Users ein:"
    $labelEmail.AutoSize = $true
    $labelEmail.Location = New-Object System.Drawing.Point(50,80)
    $form.Controls.Add($labelEmail)
    
    $textBoxEmail = New-Object System.Windows.Forms.TextBox
    $textBoxEmail.Location = New-Object System.Drawing.Point(50,110)
    $textBoxEmail.Size = New-Object System.Drawing.Size(200,20)
    $form.Controls.Add($textBoxEmail)
    
    $button = New-Object System.Windows.Forms.Button
    $button.Location = New-Object System.Drawing.Point(120,150)
    $button.Size = New-Object System.Drawing.Size(75,23)
    $button.Text = "OK"
    $button.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($button)
    
    $form.AcceptButton = $button
    
    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $nameEmail = New-Object PSObject -Property @{
            Name = $textBoxName.Text
            Email = $textBoxEmail.Text
        }
        $nameEmail
    }
    
    $form.Dispose()
}

# Benutzereingabe abfragen
$userInfo = Show-InputBoxDialog

# Abbrechen, falls kein Name eingegeben wurde
if ([string]::IsNullOrEmpty($userInfo.Name) -or [string]::IsNullOrEmpty($userInfo.Email)) {
    Write-Host "Kein vollständiger Benutzername oder E-Mail eingegeben. Skript wird abgebrochen."
    exit
}

# Outlook COM-Objekt erstellen
$Outlook = New-Object -ComObject Outlook.Application

# Neue Mail-Nachricht erstellen
$Mail = $Outlook.CreateItem(0)  # 0 steht für olMailItem

# E-Mail-Details festlegen
$Mail.Subject = "Adobe Lizenz"
$Mail.Body = "Hallo Herr Mustermann,

ich benötige eine Lizenz für folgenden User:

User: $($userInfo.Name)
E-Mail: $($userInfo.Email)"
$Mail.To = "test@hotmail.de"

# E-Mail senden
$Mail.Send()

# Outlook COM-Objekt freigeben
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
