#Requires -RunAsAdministrator
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SoftwareName,

    [Parameter(Mandatory = $false)]
    [string]$Version,

    [Parameter(Mandatory = $false)]
    [string]$ForceVersion
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ─────────────────────────────────────────────
# Parseo múltiple (NUEVO)
# ─────────────────────────────────────────────
$SoftwareList = $SoftwareName.Split(',') | ForEach-Object { $_.Trim() }

$VersionList = if ($Version) {
    $Version.Split(',') | ForEach-Object { $_.Trim() }
} else { @() }

$ForceList = @()

if ($PSBoundParameters.ContainsKey('ForceVersion')) {
    if ([string]::IsNullOrWhiteSpace($ForceVersion)) {
        $ForceList = @($true) * $SoftwareList.Count
    }
    else {
        $ForceList = $ForceVersion.Split(',') | ForEach-Object {
            $_.Trim() -match '^(1|true)$'
        }
    }
}

# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────
function Write-Step  { param([string]$Msg) Write-Host "`n[*] $Msg" -ForegroundColor Cyan }
function Write-Ok    { param([string]$Msg) Write-Host "[+] $Msg"  -ForegroundColor Green }
function Write-Warn  { param([string]$Msg) Write-Host "[!] $Msg"  -ForegroundColor Yellow }
function Write-Fail  { param([string]$Msg) Write-Host "[-] $Msg"  -ForegroundColor Red }

function Compare-Version {
    param($v1, $v2)
    try {
        return ([version]$v1).CompareTo([version]$v2)
    }
    catch {
        Write-Warn "No se pudo comparar versiones ('$v1' vs '$v2'), se omite validación estricta"
        return $null
    }
}

function Test-ConsoleAvailable {
    try {
        $null = [Console]::KeyAvailable
        return $true
    }
    catch {
        return $false
    }
}

$CanUseConsole = Test-ConsoleAvailable

# ─────────────────────────────────────────────
# WinGet module
# ─────────────────────────────────────────────
Write-Step "Verificando WinGet..."

$wingetModule = 'Microsoft.WinGet.Client'

if (-not (Get-Module -ListAvailable -Name $wingetModule)) {
    Write-Warn "Instalando módulo WinGet..."

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Install-PackageProvider -Name NuGet -Force -Scope AllUsers | Out-Null

        Set-PSRepository PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue

        Install-Module -Name $wingetModule -Force -Scope AllUsers | Out-Null
        Import-Module $wingetModule -Force
        Repair-WinGetPackageManager -Force -Latest

        Write-Ok "WinGet instalado."
    }
    catch {
        Write-Fail "Error instalando WinGet: $_"
        exit 1
    }
}
else {
    Import-Module $wingetModule -Force
    Write-Ok "WinGet OK."
}

# ─────────────────────────────────────────────
# LOOP PRINCIPAL (NUEVO)
# ─────────────────────────────────────────────
for ($i = 0; $i -lt $SoftwareList.Count; $i++) {

    $SoftwareName = $SoftwareList[$i]

    $Version = if ($i -lt $VersionList.Count -and $VersionList[$i]) {
        $VersionList[$i]
    } else { $null }

    $ForceVersion = if ($i -lt $ForceList.Count) {
        $ForceList[$i]
    } else { $false }

    Write-Step "Buscando '$SoftwareName'..."

    $results = @(Find-WinGetPackage -Query $SoftwareName)

    if (-not $results) {
        Write-Fail "Sin resultados."
        continue
    }

    function Test-IsMsStore {
        param($p)
        return ($p.Source -eq 'msstore' -or $p.Id -match '^\d{9,}$')
    }

    function Get-BestNonStoreMatch {
        param($Packages, $Query)

        $nonStore = $Packages | Where-Object { -not (Test-IsMsStore $_) }
        if (-not $nonStore) { return $null }

        $exact = $nonStore | Where-Object { $_.Name -eq $Query -or $_.Id -eq $Query }
        if ($exact) { return $exact[0] }

        return $nonStore[0]
    }

    # ─────────────────────────────────────────────
    # SELECCIÓN ORIGINAL (SIN TOCAR)
    # ─────────────────────────────────────────────
    $chosen = $null

    if ($results.Count -eq 1) {
        $chosen = $results[0]
        Write-Ok "Único paquete: $($chosen.Name)"
    }
    else {
        Write-Warn "$($results.Count) paquetes encontrados:"

        for ($j = 0; $j -lt $results.Count; $j++) {
            $pkg = $results[$j]
            $store = if (Test-IsMsStore $pkg) { " [STORE]" } else { "" }
            Write-Host ("[{0}] {1} ({2}){3}" -f $j, $pkg.Name, $pkg.Id, $store)
        }

        $userInput = $null

        if ($CanUseConsole) {
            Write-Host ""
            Write-Warn "10 segundos para elegir (Enter = auto)"
            Write-Host -NoNewline "Opción: "

            $timeout = 10
            $start = Get-Date
            $buffer = ""

            while ((Get-Date) -lt $start.AddSeconds($timeout)) {
                try {
                    if ([Console]::KeyAvailable) {
                        $key = [Console]::ReadKey($true)

                        if ($key.Key -eq 'Enter') { break }

                        if ($key.Key -eq 'Backspace') {
                            if ($buffer.Length -gt 0) {
                                $buffer = $buffer.Substring(0, $buffer.Length - 1)
                                Write-Host "`b `b" -NoNewline
                            }
                            continue
                        }

                        if ($key.KeyChar -match '\d') {
                            $buffer += $key.KeyChar
                            Write-Host $key.KeyChar -NoNewline
                        }
                    }
                }
                catch {
                    break
                }

                Start-Sleep -Milliseconds 100
            }

            Write-Host ""

            if ($buffer -match '^\d+$') {
                $userInput = [int]$buffer
            }
        }
        else {
            Write-Warn "Sin consola interactiva → auto-selección"
        }

        if ($null -ne $userInput -and $userInput -lt $results.Count) {
            $chosen = $results[$userInput]
            Write-Ok "Elegido: $($chosen.Name)"
        }
        else {
            $chosen = Get-BestNonStoreMatch $results $SoftwareName
            if (-not $chosen) { $chosen = $results[0] }
            Write-Ok "Auto-selección: $($chosen.Name)"
        }
    }

    # ─────────────────────────────────────────────
    # INSTALACIÓN (SIN CAMBIOS)
    # ─────────────────────────────────────────────
    Write-Step "Verificando instalación..."

    $installed = Get-WinGetPackage -Id $chosen.Id -ErrorAction SilentlyContinue

    if ($installed) {
        Write-Ok "Instalado: $($installed.InstalledVersion)"

        if ($Version) {

            Write-Step "Modo versión específica: $Version"

            $cmp = Compare-Version $installed.InstalledVersion $Version

            if ($cmp -eq 0) {
                Write-Ok "Ya está en la versión objetivo."
            }
            elseif ($cmp -lt 0) {
                Install-WinGetPackage -Id $chosen.Id -Version $Version -Mode Silent -Force
                Write-Ok "Actualizado a $Version"
            }
            elseif ($cmp -gt 0) {
                if ($ForceVersion) {
                    Write-Warn "Downgrade forzado"

                    Uninstall-WinGetPackage -Id $chosen.Id -Force
                    Install-WinGetPackage -Id $chosen.Id -Version $Version -Mode Silent -Force

                    Write-Ok "Downgrade completado."
                }
                else {
                    Write-Warn "Versión superior instalada. Use -ForceVersion para bajar versión."
                }
            }
        }
        else {
            Write-Step "Modo latest"

            $available = Find-WinGetPackage -Id $chosen.Id | Select-Object -First 1

            if ($available -and $available.Version -ne $installed.InstalledVersion) {
                Write-Step "Actualizando..."
                Update-WinGetPackage -Id $chosen.Id -Mode Silent -Force
                Write-Ok "Actualizado."
            }
            else {
                Write-Ok "Ya actualizado."
            }
        }
    }
    else {
        Write-Step "Instalando..."

        if ($Version) {
            Install-WinGetPackage -Id $chosen.Id -Version $Version -Mode Silent -Force
        }
        else {
            Install-WinGetPackage -Id $chosen.Id -Mode Silent -Force
        }

        Write-Ok "Instalado."
    }
}

Write-Host ""
Write-Ok "Script finalizado correctamente."