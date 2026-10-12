$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms

try {
    $sourceDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
    $installDirectory = Join-Path $env:LOCALAPPDATA 'MLC Claimant Drop'
    New-Item -ItemType Directory -Path $installDirectory -Force | Out-Null

    Copy-Item (Join-Path $sourceDirectory 'MLCClaimantDrop.ps1') $installDirectory -Force
    Copy-Item (Join-Path $sourceDirectory 'MLC Claimant Drop.cmd') $installDirectory -Force

    $desktop = [Environment]::GetFolderPath('Desktop')
    $sendTo = [Environment]::GetFolderPath('SendTo')
    if ([string]::IsNullOrWhiteSpace($sendTo)) {
        throw 'Windows could not locate your Send To folder.'
    }
    New-Item -ItemType Directory -Path $sendTo -Force | Out-Null

    $shell = New-Object -ComObject WScript.Shell
    foreach ($shortcutDirectory in @($sendTo, $desktop)) {
        $shortcutPath = Join-Path $shortcutDirectory 'MLC Claimant Drop.lnk'
        $shortcut = $shell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = Join-Path $installDirectory 'MLC Claimant Drop.cmd'
        $shortcut.WorkingDirectory = $installDirectory
        $shortcut.Description = 'Review transfer scope and copy administrator/publisher records.'
        $shortcut.IconLocation = "$env:SystemRoot\System32\shell32.dll,71"
        $shortcut.Save()
        if (-not (Test-Path -LiteralPath $shortcutPath)) {
            throw "Windows could not create the shortcut: $shortcutPath"
        }
    }

    [void][System.Windows.Forms.MessageBox]::Show(
        "Installed!`r`n`r`nRight-click one .xlsx report, choose Send to, then MLC Claimant Drop.`r`n`r`nIf Send to is hidden, choose Show more options first.`r`n`r`nYou can also drag a report onto the desktop icon.",
        'MLC Claimant Drop',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
}
catch {
    [void][System.Windows.Forms.MessageBox]::Show(
        "Installation failed:`r`n`r`n$($_.Exception.Message)",
        'MLC Claimant Drop',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit 1
}
