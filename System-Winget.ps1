#Requires -RunAsAdministrator
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SoftwareName,

    [Parameter(Mandatory = $false)]
    [string]$Version = "",

    [Parameter(Mandatory = $false)]
    [string]$ForceVersion = "",

    [Parameter(Mandatory = $false)]
    [string]$Uninstall = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ─────────────────────────────────────────────
# Helpers de output
# ─────────────────────────────────────────────
function Write-Step { param([string]$Msg) Write-Host "`n[*] $Msg" -ForegroundColor Cyan   }
function Write-Ok   { param([string]$Msg) Write-Host "[+] $Msg"  -ForegroundColor Green  }
function Write-Warn { param([string]$Msg) Write-Host "[!] $Msg"  -ForegroundColor Yellow }
function Write-Fail { param([string]$Msg) Write-Host "[-] $Msg"  -ForegroundColor Red    }

# ─────────────────────────────────────────────
# Parseo de parámetros
# ─────────────────────────────────────────────
$SoftwareList = @($SoftwareName -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
$count        = @($SoftwareList).Count

# Version: debe tener 0 entradas (latest para todo) o exactamente $count entradas
$VersionList = if ($Version.Trim()) {
    @($Version -split ',' | ForEach-Object { $_.Trim() })
} else { @() }

if (@($VersionList).Count -ne 0 -and @($VersionList).Count -ne $count) {
    Write-Fail "-Version debe estar vacío (latest para todo) o contener exactamente $count entradas (una por software). Encontrado: $(@($VersionList).Count)"
    exit 1
}

# ForceVersion: un único valor se aplica a todos; array → debe tener $count entradas
function Resolve-BoolParam {
    param([string]$Raw, [int]$Count, [string]$ParamName)

    if (-not $Raw.Trim()) { return @($false) * $Count }

    $parts = @($Raw -split ',' | ForEach-Object { $_.Trim() })

    if (@($parts).Count -eq 1) {
        $val = $parts[0] -match '^(1|true)$'
        return @($val) * $Count
    }

    if (@($parts).Count -ne $Count) {
        Write-Fail "-$ParamName debe tener 1 valor (global) o $Count valores (uno por software). Encontrado: $(@($parts).Count)"
        exit 1
    }

    return @($parts | ForEach-Object { $_ -match '^(1|true)$' })
}

$ForceList     = @(Resolve-BoolParam -Raw $ForceVersion -Count $count -ParamName 'ForceVersion')
$UninstallList = @(Resolve-BoolParam -Raw $Uninstall   -Count $count -ParamName 'Uninstall')

# ─────────────────────────────────────────────
# WinGet — instalación del módulo si falta
# ─────────────────────────────────────────────
Write-Step "Verificando WinGet..."

$wingetModule = 'Microsoft.WinGet.Client'

if (-not (Get-Module -ListAvailable -Name $wingetModule)) {
    Write-Warn "Módulo no encontrado, instalando..."
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Install-PackageProvider -Name NuGet -Force -Scope AllUsers | Out-Null
        Set-PSRepository PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
        Install-Module -Name $wingetModule -Force -Scope AllUsers | Out-Null
        Import-Module $wingetModule -Force
        Repair-WinGetPackageManager -Force -Latest
        Write-Ok "WinGet instalado."
    } catch {
        Write-Fail "Error instalando WinGet: $_"
        exit 1
    }
} else {
    Import-Module $wingetModule -Force
    Write-Ok "WinGet OK."
}

# ─────────────────────────────────────────────
# Helpers de paquetes
# ─────────────────────────────────────────────
function Test-IsMsStore {
    param($p)
    return ($p.Source -eq 'msstore' -or $p.Id -match '^\d{9,}$')
}

function Compare-Version {
    param($v1, $v2)
    try   { return ([version]$v1).CompareTo([version]$v2) }
    catch { Write-Warn "No se pudo comparar '$v1' vs '$v2', se omite validación estricta"; return $null }
}

function Test-ConsoleAvailable {
    try { $null = [Console]::KeyAvailable; return $true } catch { return $false }
}

# Auto-selección: exacto primero, luego primer non-store, luego primero de la lista
function Get-AutoChoice {
    param([array]$Packages, [string]$Query)
    $nonStore = @($Packages | Where-Object { -not (Test-IsMsStore $_) })
    $pool     = if (@($nonStore).Count -gt 0) { $nonStore } else { $Packages }
    $exact    = @($pool | Where-Object { $_.Name -eq $Query -or $_.Id -eq $Query })
    if (@($exact).Count -gt 0) { return $exact[0] } else { return $pool[0] }
}

# Selección de paquete con timeout interactivo
function Select-Package {
    param([array]$Packages, [string]$Query)

    $pkgCount = @($Packages).Count

    if ($pkgCount -eq 1) {
        Write-Ok "Único paquete: $($Packages[0].Name)"
        return $Packages[0]
    }

    Write-Warn "$pkgCount paquetes encontrados:"
    for ($j = 0; $j -lt $pkgCount; $j++) {
        $store = if (Test-IsMsStore $Packages[$j]) { " [STORE]" } else { "" }
        Write-Host ("  [{0}] {1}  ({2}){3}" -f $j, $Packages[$j].Name, $Packages[$j].Id, $store)
    }

    if (-not (Test-ConsoleAvailable)) {
        Write-Warn "Sin consola interactiva → auto-selección"
        $auto = Get-AutoChoice -Packages $Packages -Query $Query
        Write-Ok "Auto-seleccionado: $($auto.Name)"
        return $auto
    }

    Write-Host ""
    Write-Warn "10 segundos para elegir (Enter = auto-selección)"
    Write-Host -NoNewline "Opción: "

    $buffer = ""
    $start  = Get-Date

    while ((Get-Date) -lt $start.AddSeconds(10)) {
        try {
            if ([Console]::KeyAvailable) {
                $key = [Console]::ReadKey($true)
                if ($key.Key -eq 'Enter') { break }
                if ($key.Key -eq 'Backspace' -and $buffer.Length -gt 0) {
                    $buffer = $buffer[0..($buffer.Length - 2)] -join ''
                    Write-Host "`b `b" -NoNewline
                } elseif ($key.KeyChar -match '\d') {
                    $buffer += $key.KeyChar
                    Write-Host $key.KeyChar -NoNewline
                }
            }
        } catch { break }
        Start-Sleep -Milliseconds 100
    }

    Write-Host ""

    if ($buffer -match '^\d+$' -and [int]$buffer -lt $pkgCount) {
        $chosen = $Packages[[int]$buffer]
        Write-Ok "Elegido manualmente: $($chosen.Name)"
        return $chosen
    }

    $auto = Get-AutoChoice -Packages $Packages -Query $Query
    Write-Ok "Auto-seleccionado: $($auto.Name)"
    return $auto
}

# ─────────────────────────────────────────────
# LOOP PRINCIPAL
# ─────────────────────────────────────────────
for ($i = 0; $i -lt $count; $i++) {

    $sw           = $SoftwareList[$i]
    $targetVer    = if (@($VersionList).Count -eq $count) { $VersionList[$i] } else { "" }
    $forceDown    = $ForceList[$i]
    $doUninstall  = $UninstallList[$i]

    Write-Step "[$($i+1)/$count] Procesando: '$sw'"

    # ── Búsqueda ──────────────────────────────
    $results = @(Find-WinGetPackage -Query $sw)

    if (@($results).Count -eq 0) {
        Write-Fail "Sin resultados para '$sw'. Se omite."
        continue
    }

    $chosen = Select-Package -Packages $results -Query $sw

    # ── Desinstalación ────────────────────────
    if ($doUninstall) {
        $installed = Get-WinGetPackage -Id $chosen.Id -ErrorAction SilentlyContinue
        if ($installed) {
            Write-Step "Desinstalando '$($chosen.Name)'..."
            Uninstall-WinGetPackage -Id $chosen.Id -Force
            Write-Ok "Desinstalado."
        } else {
            Write-Warn "'$($chosen.Name)' no estaba instalado."
        }
        continue
    }

    # ── Instalación / actualización ───────────
    $installed = Get-WinGetPackage -Id $chosen.Id -ErrorAction SilentlyContinue

    if ($installed) {
        Write-Ok "Versión instalada: $($installed.InstalledVersion)"

        if ($targetVer) {
            # ── Modo versión específica ──
            $cmp = Compare-Version $installed.InstalledVersion $targetVer

            switch ($cmp) {
                0 { Write-Ok "Ya está en la versión objetivo ($targetVer)." }

                { $_ -lt 0 } {
                    Write-Step "Actualizando a $targetVer..."
                    Install-WinGetPackage -Id $chosen.Id -Version $targetVer -Mode Silent -Force
                    Write-Ok "Actualizado a $targetVer."
                }

                { $_ -gt 0 } {
                    if ($forceDown) {
                        Write-Warn "Downgrade forzado → $targetVer"
                        Uninstall-WinGetPackage -Id $chosen.Id -Force
                        Install-WinGetPackage -Id $chosen.Id -Version $targetVer -Mode Silent -Force
                        Write-Ok "Downgrade completado."
                    } else {
                        Write-Warn "Versión instalada ($($installed.InstalledVersion)) > objetivo ($targetVer). Use -ForceVersion para hacer downgrade."
                    }
                }
            }
        } else {
            # ── Modo latest ──
            $available = Find-WinGetPackage -Id $chosen.Id | Select-Object -First 1
            if ($available -and $available.Version -ne $installed.InstalledVersion) {
                Write-Step "Actualizando a latest ($($available.Version))..."
                Update-WinGetPackage -Id $chosen.Id -Mode Silent -Force
                Write-Ok "Actualizado."
            } else {
                Write-Ok "Ya en la última versión."
            }
        }
    } else {
        Write-Step "Instalando '$($chosen.Name)'$(if ($targetVer) { " v$targetVer" })..."
        if ($targetVer) {
            Install-WinGetPackage -Id $chosen.Id -Version $targetVer -Mode Silent -Force
        } else {
            Install-WinGetPackage -Id $chosen.Id -Mode Silent -Force
        }
        Write-Ok "Instalado correctamente."
    }
}

Write-Host ""
Write-Ok "Script finalizado."