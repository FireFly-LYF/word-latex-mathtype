param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,
    [string]$TargetPath,
    [string]$LogPath,
    [string]$WordExe,
    [ValidateSet('toggle', 'word')]
    [string]$MathtypeMode = 'toggle',
    [switch]$PreferWord,
    [switch]$UseLegacy
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Resolve-DefaultOutputPath {
    param(
        [string]$InputPath,
        [string]$Suffix
    )

    $directory = Split-Path -Parent $InputPath
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($InputPath)
    $extension = [System.IO.Path]::GetExtension($InputPath)
    $suffixExtension = [System.IO.Path]::GetExtension($Suffix)

    if ([string]::IsNullOrWhiteSpace($suffixExtension)) {
        return (Join-Path $directory ($baseName + $Suffix + $extension))
    }

    return (Join-Path $directory ($baseName + $Suffix))
}

function Resolve-PythonExecutable {
    foreach ($name in @('python', 'python3', 'py')) {
        $command = Get-Command $name -ErrorAction SilentlyContinue
        if ($null -ne $command -and -not [string]::IsNullOrWhiteSpace($command.Source)) {
            if ($name -eq 'py') {
                return @($command.Source, '-3')
            }
            return @($command.Source)
        }
    }

    return $null
}

function Get-WordInstallPath {
    $appPathKeys = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE'
    )

    foreach ($key in $appPathKeys) {
        $item = Get-ItemProperty -Path $key -ErrorAction SilentlyContinue
        if ($null -ne $item -and (Test-Path -LiteralPath $item.'(default)')) {
            return $item.'(default)'
        }
    }

    $candidates = @(
        'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE',
        'C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE'
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    return $null
}

$SourcePath = (Resolve-Path -LiteralPath $SourcePath).Path
if ([string]::IsNullOrWhiteSpace($TargetPath)) {
    $TargetPath = Resolve-DefaultOutputPath -InputPath $SourcePath -Suffix '_formula.docx'
}
if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $LogPath = Resolve-DefaultOutputPath -InputPath $SourcePath -Suffix '_formula_log.json'
}

$TargetPath = [System.IO.Path]::GetFullPath($TargetPath)
$LogPath = [System.IO.Path]::GetFullPath($LogPath)

if ($SourcePath -eq $TargetPath) {
    throw 'TargetPath must be different from SourcePath.'
}

if ($UseLegacy) {
    $legacyScript = Join-Path $PSScriptRoot 'convert_docx_latex_to_equations_legacy.ps1'
    if (-not (Test-Path -LiteralPath $legacyScript)) {
        throw "Legacy converter not found: $legacyScript"
    }

    & $legacyScript -SourcePath $SourcePath -TargetPath $TargetPath -LogPath $LogPath
    exit $LASTEXITCODE
}

$pythonScript = Join-Path $PSScriptRoot 'convert_docx_latex_to_equations.py'
if (-not (Test-Path -LiteralPath $pythonScript)) {
    throw "Python converter not found: $pythonScript"
}

$pythonCommand = Resolve-PythonExecutable
if ($null -eq $pythonCommand) {
    throw 'Python was not found. Run scripts/setup_environment.ps1 first.'
}

& @pythonCommand -c "import win32com.client" 1>$null 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host 'Installing pywin32...'
    & @pythonCommand -m pip install pywin32 | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw 'Failed to install pywin32. Run scripts/setup_environment.ps1 first.'
    }
}

if ([string]::IsNullOrWhiteSpace($WordExe)) {
    $WordExe = Get-WordInstallPath
}

$pythonArgs = @(
    $pythonScript,
    '--source', $SourcePath,
    '--target', $TargetPath,
    '--log', $LogPath
)

if (-not [string]::IsNullOrWhiteSpace($WordExe)) {
    $pythonArgs += @('--word-exe', $WordExe)
}

if ($PreferWord) {
    $pythonArgs += @('--mathtype-mode', 'word')
} else {
    $pythonArgs += @('--mathtype-mode', $MathtypeMode)
}

Write-Host "Using headless MathType TeX-toggle converter: $pythonScript"
& @pythonCommand @pythonArgs
exit $LASTEXITCODE
