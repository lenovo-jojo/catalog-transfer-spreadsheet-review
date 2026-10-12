Add-Type -AssemblyName System.Windows.Forms

try {
    $sourceDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
    $installDirectory = Join-Path $env:LOCALAPPDATA 'MLC Claimant Drop'
    New-Item -ItemType Directory -Path $installDirectory -Force | Out-Null

    Copy-Item (Join-Path $sourceDirectory 'MLCClaimantDrop.ps1') $installDirectory -Force
    Copy-Item (Join-Path $sourceDirectory 'MLC Claimant Drop.cmd') $installDirectory -Force

    $desktop = [Environment]::GetFolderPath('Desktop')
    $shortcutPath = Join-Path $desktop 'MLC Claimant Drop.lnk'
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = Join-Path $installDirectory 'MLC Claimant Drop.cmd'
    $shortcut.WorkingDirectory = $installDirectory
    $shortcut.Description = 'Drop an MLC Excel report here to copy unique YES claimants.'
    $shortcut.IconLocation = "$env:SystemRoot\System32\shell32.dll,71"
    $shortcut.Save()

    [void][System.Windows.Forms.MessageBox]::Show(
        "Installed. Look for 'MLC Claimant Drop' on your desktop.`r`n`r`nDrag an .xlsx report onto that icon.",
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
