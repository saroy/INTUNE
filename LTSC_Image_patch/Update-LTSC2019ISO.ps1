#Requires -Version 5.1
#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$RootPath = "C:\Users\pe152503\OneDrive - Pierre Fabre SA\Documents\scripts\LTSC2019",

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$TargetEditionPattern = "Enterprise LTSC 2019",

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OutputIsoName = "Windows10_Enterprise_LTSC2019_x64_Patched.iso",

    # Dossier de travail local, hors OneDrive et sans espace dans le chemin :
    # DISM monte l'image dedans et oscdimg ne gère pas les espaces dans -bootdata.
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$WorkPath = "C:\LTSC2019_Work"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

# ============================================================================
# CHEMINS
# ============================================================================

$IsoFolder = Join-Path -Path $RootPath -ChildPath "ISO"
$UpdatesFolder = Join-Path -Path $RootPath -ChildPath "Updates"
$OutputFolder = Join-Path -Path $RootPath -ChildPath "Output"
$WorkFolder = $WorkPath
$MediaFolder = Join-Path -Path $WorkFolder -ChildPath "Media"
$MountFolder = Join-Path -Path $WorkFolder -ChildPath "Mount"
$ScratchFolder = Join-Path -Path $WorkFolder -ChildPath "Scratch"
$BackupFolder = Join-Path -Path $WorkFolder -ChildPath "Backup"
$WinReFolder = Join-Path -Path $WorkFolder -ChildPath "WinRE"
$WinReMountFolder = Join-Path -Path $WorkFolder -ChildPath "MountWinRE"
$BootMountFolder = Join-Path -Path $WorkFolder -ChildPath "MountBoot"
$LogsFolder = Join-Path -Path $RootPath -ChildPath "Logs"

$OutputIsoPath = Join-Path -Path $OutputFolder -ChildPath $OutputIsoName

$TimeStamp = Get-Date -Format "yyyyMMdd-HHmmss"

$TranscriptPath = Join-Path `
    -Path $LogsFolder `
    -ChildPath "LTSC2019-Servicing-$TimeStamp.log"

$DismLogPath = Join-Path `
    -Path $LogsFolder `
    -ChildPath "DISM-$TimeStamp.log"

$PackageListPath = Join-Path `
    -Path $LogsFolder `
    -ChildPath "Packages-$TimeStamp.txt"

$MountedIso = $false
$MountedWindowsImage = $false
$MountedWinRE = $false
$MountedBootImage = $false
$IsoPath = $null
$TranscriptStarted = $false

$IntegratedPackages = New-Object System.Collections.Generic.List[string]

# ============================================================================
# FONCTIONS D'AFFICHAGE
# ============================================================================

function Write-Step {
    param(
        [Parameter(Mandatory)]
        [string]$Message
    )

    Write-Host ""
    Write-Host "======================================================================" -ForegroundColor Cyan
    Write-Host $Message -ForegroundColor Cyan
    Write-Host "======================================================================" -ForegroundColor Cyan
}

function Write-Info {
    param(
        [Parameter(Mandatory)]
        [string]$Message
    )

    Write-Host "[INFO] $Message" -ForegroundColor Gray
}

function Write-Success {
    param(
        [Parameter(Mandatory)]
        [string]$Message
    )

    Write-Host "[OK] $Message" -ForegroundColor Green
}

function Write-WarningMessage {
    param(
        [Parameter(Mandatory)]
        [string]$Message
    )

    Write-Host "[ATTENTION] $Message" -ForegroundColor Yellow
}

# ============================================================================
# FONCTIONS TECHNIQUES
# ============================================================================

function Test-IsAdministrator {
    $CurrentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()

    $Principal = New-Object `
        -TypeName Security.Principal.WindowsPrincipal `
        -ArgumentList $CurrentIdentity

    return $Principal.IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator
    )
}

function Invoke-Dism {
    param(
        [Parameter(Mandatory)]
        [string[]]$Arguments,

        [Parameter()]
        [switch]$IgnoreError
    )

    Write-Host ""
    Write-Host "DISM.exe $($Arguments -join ' ')" -ForegroundColor DarkGray
    Write-Host ""

    # La sortie de DISM est envoyée à l'écran (Out-Host) et non dans le pipeline :
    # sinon la fonction renverrait tout le texte de DISM en plus du code de retour.
    & "$env:SystemRoot\System32\dism.exe" @Arguments | Out-Host

    $ExitCode = $LASTEXITCODE

    if ($ExitCode -ne 0) {
        if ($IgnoreError) {
            Write-WarningMessage "DISM a retourné le code $ExitCode. Le traitement continue."
            return $ExitCode
        }

        throw "Échec de DISM. Code de retour : $ExitCode"
    }

    return $ExitCode
}

function Get-OscdimgPath {
    $CandidatePaths = @(
        "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe",
        "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\x86\Oscdimg\oscdimg.exe",
        "$env:ProgramFiles\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe",
        "$env:ProgramFiles\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\x86\Oscdimg\oscdimg.exe"
    )

    foreach ($CandidatePath in $CandidatePaths) {
        if (Test-Path -LiteralPath $CandidatePath) {
            return $CandidatePath
        }
    }

    $OscdimgCommand = Get-Command `
        -Name "oscdimg.exe" `
        -ErrorAction SilentlyContinue

    if ($null -ne $OscdimgCommand) {
        return $OscdimgCommand.Source
    }

    throw @"
Oscdimg.exe est introuvable.

Installe le Windows ADK avec le composant Deployment Tools,
puis relance ce script.
"@
}

function Get-RootDriveFreeSpace {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $FullPath = [System.IO.Path]::GetFullPath($Path)
    $DriveRoot = [System.IO.Path]::GetPathRoot($FullPath)

    if ([string]::IsNullOrWhiteSpace($DriveRoot)) {
        throw "Impossible de déterminer le lecteur associé au chemin : $Path"
    }

    $Drive = New-Object -TypeName System.IO.DriveInfo -ArgumentList $DriveRoot

    return [PSCustomObject]@{
        Name      = $Drive.Name.TrimEnd('\').TrimEnd(':')
        FreeBytes = [int64]$Drive.AvailableFreeSpace
    }
}

function Initialize-WorkingFolders {
    Write-Info "Nettoyage des éventuels montages DISM abandonnés."

    foreach ($PreviousMountDir in @($MountFolder, $WinReMountFolder, $BootMountFolder)) {
        Dismount-OrphanRegistryHives -MountDir $PreviousMountDir
    }

    Invoke-Dism -Arguments @(
        "/English",
        "/Cleanup-Mountpoints"
    ) -IgnoreError | Out-Null

    if (Test-Path -LiteralPath $WorkFolder) {
        Write-Info "Suppression de l'ancien dossier de travail."

        Get-ChildItem `
            -LiteralPath $WorkFolder `
            -Recurse `
            -Force `
            -ErrorAction SilentlyContinue |
            ForEach-Object {
                try {
                    $_.Attributes = [System.IO.FileAttributes]::Normal
                }
                catch {
                    # L'élément sera traité par Remove-Item.
                }
            }

        Remove-Item `
            -LiteralPath $WorkFolder `
            -Recurse `
            -Force `
            -ErrorAction Stop
    }

    $RequiredFolders = @(
        $IsoFolder,
        $UpdatesFolder,
        $OutputFolder,
        $WorkFolder,
        $MediaFolder,
        $MountFolder,
        $ScratchFolder,
        $BackupFolder,
        $WinReFolder,
        $WinReMountFolder,
        $BootMountFolder,
        $LogsFolder
    )

    foreach ($Folder in $RequiredFolders) {
        if (-not (Test-Path -LiteralPath $Folder)) {
            New-Item `
                -Path $Folder `
                -ItemType Directory `
                -Force |
                Out-Null
        }
    }
}

function Copy-IsoContent {
    param(
        [Parameter(Mandatory)]
        [string]$SourceDrive,

        [Parameter(Mandatory)]
        [string]$Destination
    )

    $RobocopyArguments = @(
        $SourceDrive,
        $Destination,
        "/E",
        "/COPY:DAT",
        "/DCOPY:DAT",
        "/R:2",
        "/W:2",
        "/XJ",
        "/NFL",
        "/NDL",
        "/NP"
    )

    Write-Info "Copie du contenu de l'ISO avec Robocopy."

    & robocopy.exe @RobocopyArguments

    $RobocopyExitCode = $LASTEXITCODE

    # Pour Robocopy, les codes de 0 à 7 indiquent une réussite.
    if ($RobocopyExitCode -ge 8) {
        throw "Échec de Robocopy. Code de retour : $RobocopyExitCode"
    }

    Write-Success "Contenu de l'ISO copié dans : $Destination"
}

function Remove-ReadOnlyAttributes {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    Write-Info "Suppression des attributs Lecture seule."

    Get-ChildItem `
        -LiteralPath $Path `
        -Recurse `
        -Force `
        -ErrorAction Stop |
        ForEach-Object {
            try {
                if ($_.Attributes -band [System.IO.FileAttributes]::ReadOnly) {
                    $_.Attributes = $_.Attributes -bxor [System.IO.FileAttributes]::ReadOnly
                }
            }
            catch {
                Write-WarningMessage "Attribut impossible à modifier : $($_.FullName)"
            }
        }
}

function Select-TargetWindowsImage {
    param(
        [Parameter(Mandatory)]
        [object[]]$Images,

        [Parameter(Mandatory)]
        [string]$EditionPattern
    )

    Write-Host ""
    Write-Host "Images détectées dans le média :" -ForegroundColor White

    foreach ($Image in $Images) {
        Write-Host (
            "  Index {0} : {1} | Architecture : {2} | Version : {3}" -f
            $Image.ImageIndex,
            $Image.ImageName,
            $Image.Architecture,
            $Image.Version
        )
    }

    $MatchingImages = @(
        $Images | Where-Object {
            $_.ImageName -like "*$EditionPattern*"
        }
    )

    # L'ISO française nomme l'édition "Windows 10 Entreprise LTSC", sans l'année
    # (le build 17763 est contrôlé à l'étape 8). L'édition N est exclue :
    # "Entreprise N LTSC" ne correspond pas au motif.
    if ($MatchingImages.Count -eq 0) {
        $MatchingImages = @(
            $Images | Where-Object {
                $_.ImageName -match "(Enterprise|Entreprise) LTSC"
            }
        )
    }

    if ($MatchingImages.Count -eq 0) {
        $DetectedEditions = $Images |
            ForEach-Object {
                "Index $($_.ImageIndex) : $($_.ImageName)"
            }

        throw @"
L'édition "$EditionPattern" n'a pas été trouvée.

Éditions détectées :
$($DetectedEditions -join [Environment]::NewLine)
"@
    }

    if ($MatchingImages.Count -gt 1) {
        $X64Match = $MatchingImages |
            Where-Object {
                $_.Architecture -in @("x64", "amd64", 9)
            } |
            Select-Object -First 1

        if ($null -ne $X64Match) {
            return $X64Match
        }

        Write-WarningMessage "Plusieurs éditions correspondent. La première sera utilisée."
    }

    return $MatchingImages | Select-Object -First 1
}

function Export-ImageToWim {
    param(
        [Parameter(Mandatory)]
        [string]$SourceImage,

        [Parameter(Mandatory)]
        [int]$SourceIndex,

        [Parameter(Mandatory)]
        [string]$DestinationImage,

        [Parameter(Mandatory)]
        [string]$DestinationName
    )

    if (Test-Path -LiteralPath $DestinationImage) {
        Remove-Item `
            -LiteralPath $DestinationImage `
            -Force `
            -ErrorAction Stop
    }

    Invoke-Dism -Arguments @(
        "/English",
        "/Export-Image",
        "/SourceImageFile:$SourceImage",
        "/SourceIndex:$SourceIndex",
        "/DestinationImageFile:$DestinationImage",
        "/DestinationName:$DestinationName",
        "/Compress:max",
        "/CheckIntegrity",
        "/ScratchDir:$ScratchFolder",
        "/LogPath:$DismLogPath"
    ) | Out-Null

    if (-not (Test-Path -LiteralPath $DestinationImage)) {
        throw "Le fichier WIM n'a pas été créé : $DestinationImage"
    }
}

function Get-UpdatePackageCategory {
    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$Package
    )

    # 0 = SSU autonome, 1 = cumulative Windows (LCU), 2 = .NET, 3 = autre.
    if ($Package.Name -match "ssu|servicing") { return 0 }
    if ($Package.Name -match "ndp|dotnet|netfx") { return 2 }
    if ($Package.Name -match "^windows10\.0-kb\d+-x64[_.]") { return 1 }

    return 3
}

function Add-UpdatePackages {
    param(
        [Parameter(Mandatory)]
        [string]$MountDir,

        [Parameter(Mandatory)]
        [System.IO.FileInfo[]]$Packages,

        [Parameter()]
        [switch]$IgnoreError
    )

    foreach ($Package in $Packages) {
        Write-Info "Intégration : $($Package.Name)"

        $ExitCode = Invoke-Dism -Arguments @(
            "/English",
            "/Image:$MountDir",
            "/Add-Package",
            "/PackagePath:$($Package.FullName)",
            "/ScratchDir:$ScratchFolder",
            "/LogPath:$DismLogPath"
        ) -IgnoreError:$IgnoreError

        if ($ExitCode -eq 0) {
            Write-Success "Package intégré : $($Package.Name)"
        }
    }
}

function Invoke-ComponentCleanup {
    param(
        [Parameter(Mandatory)]
        [string]$MountDir
    )

    Invoke-Dism -Arguments @(
        "/English",
        "/Image:$MountDir",
        "/Cleanup-Image",
        "/StartComponentCleanup",
        "/ResetBase",
        "/ScratchDir:$ScratchFolder",
        "/LogPath:$DismLogPath"
    ) -IgnoreError | Out-Null
}

function Dismount-OrphanRegistryHives {
    param(
        [Parameter(Mandatory)]
        [string]$MountDir
    )

    # DISM/CBS charge les ruches de l'image sous HKLM\{bf1a281b-...}C:/chemin/...
    # Si elles restent chargées, leurs fichiers sont verrouillés et le démontage
    # échoue en 0xC1420117.
    $HivePrefix = "*{bf1a281b-ad7b-4476-ac95-f47682990ce7}" + $MountDir.Replace('\', '/').TrimEnd('/') + "/*"

    # reg.exe écrit sur stderr : sous PowerShell 5.1 cela ne doit pas être bloquant.
    $PreviousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = "Continue"

    try {
        $LoadedHives = @(
            & reg.exe query HKLM 2>&1 |
                ForEach-Object { $_.ToString().Trim() } |
                Where-Object { $_ -like $HivePrefix }
        )

        foreach ($Hive in $LoadedHives) {
            Write-WarningMessage "Ruche encore chargée, déchargement : $Hive"

            & reg.exe unload $Hive 2>&1 | Out-Null

            if ($LASTEXITCODE -ne 0) {
                Write-WarningMessage "Impossible de décharger : $Hive"
            }
        }
    }
    finally {
        $ErrorActionPreference = $PreviousErrorActionPreference
    }
}

function Dismount-ImageCommit {
    param(
        [Parameter(Mandatory)]
        [string]$MountDir,

        [Parameter()]
        [int]$MaxAttempts = 6,

        [Parameter()]
        [int]$RetryDelaySeconds = 60
    )

    # 0xC1420117 : le dossier de montage n'a pas pu être libéré car des fichiers y sont
    # encore ouverts (antivirus/EDR, Explorateur...). On attend puis on recommence.
    $UnmountBusyCode = -1052638953

    for ($Attempt = 1; $Attempt -le $MaxAttempts; $Attempt++) {
        Dismount-OrphanRegistryHives -MountDir $MountDir

        $ExitCode = Invoke-Dism -Arguments @(
            "/English",
            "/Unmount-Image",
            "/MountDir:$MountDir",
            "/Commit",
            "/ScratchDir:$ScratchFolder",
            "/LogPath:$DismLogPath"
        ) -IgnoreError

        if ($ExitCode -eq 0) {
            return
        }

        if ($ExitCode -ne $UnmountBusyCode) {
            throw "Échec du démontage de $MountDir. Code de retour DISM : $ExitCode"
        }

        Write-WarningMessage (
            "Fichiers encore ouverts dans {0} (tentative {1}/{2}). Nouvel essai dans {3} s." -f
            $MountDir,
            $Attempt,
            $MaxAttempts,
            $RetryDelaySeconds
        )

        Start-Sleep -Seconds $RetryDelaySeconds
    }

    throw @"
Impossible de démonter $MountDir : des fichiers y restent ouverts (code 0xC1420117).
Ferme les fenêtres de l'Explorateur sur ce dossier et vérifie l'antivirus, puis exécute :
  dism /Unmount-Image /MountDir:$MountDir /Commit
"@
}

function Dismount-ImageDiscard {
    param(
        [Parameter(Mandatory)]
        [string]$MountDir
    )

    Write-WarningMessage "Image encore montée dans $MountDir. Démontage sans validation."

    try {
        Dismount-OrphanRegistryHives -MountDir $MountDir

        Invoke-Dism -Arguments @(
            "/English",
            "/Unmount-Image",
            "/MountDir:$MountDir",
            "/Discard"
        ) -IgnoreError | Out-Null
    }
    catch {
        Write-WarningMessage "Le démontage d'urgence de $MountDir a échoué."
    }
}

# ============================================================================
# CONTRÔLES INITIAUX
# ============================================================================

if (-not (Test-IsAdministrator)) {
    throw "Le script doit être exécuté dans Windows PowerShell en tant qu'administrateur."
}

if (-not (Test-Path -LiteralPath $RootPath)) {
    New-Item `
        -Path $RootPath `
        -ItemType Directory `
        -Force |
        Out-Null
}

if (-not (Test-Path -LiteralPath $LogsFolder)) {
    New-Item `
        -Path $LogsFolder `
        -ItemType Directory `
        -Force |
        Out-Null
}

Start-Transcript `
    -Path $TranscriptPath `
    -Force |
    Out-Null

$TranscriptStarted = $true

try {
    # ========================================================================
    # ÉTAPE 1 : INITIALISATION
    # ========================================================================

    Write-Step "1. Initialisation"

    Write-Info "Dossier racine : $RootPath"
    Write-Info "Dossier de travail : $WorkFolder"
    Write-Info "Édition recherchée : $TargetEditionPattern"
    Write-Info "ISO de sortie : $OutputIsoPath"

    Initialize-WorkingFolders

    $OscdimgPath = Get-OscdimgPath
    Write-Success "Oscdimg détecté : $OscdimgPath"

    $RootDriveInformation = Get-RootDriveFreeSpace -Path $WorkFolder
    $FreeSpace = $RootDriveInformation.FreeBytes
    $RootDriveName = $RootDriveInformation.Name

    if ($FreeSpace -lt 25GB) {
        Write-WarningMessage (
            "Espace libre inférieur à 25 Go sur le lecteur {0}: {1:N2} Go." -f
            $RootDriveName,
            ($FreeSpace / 1GB)
        )
    }
    else {
        Write-Success (
            "Espace libre disponible sur le lecteur {0}: {1:N2} Go." -f
            $RootDriveName,
            ($FreeSpace / 1GB)
        )
    }

    # ========================================================================
    # ÉTAPE 2 : RECHERCHE DE L'ISO
    # ========================================================================

    Write-Step "2. Recherche de l'ISO source"

    $IsoFiles = @(
        Get-ChildItem `
            -LiteralPath $IsoFolder `
            -Filter "*.iso" `
            -File |
            Sort-Object -Property Name
    )

    if ($IsoFiles.Count -eq 0) {
        throw "Aucune ISO trouvée dans le dossier : $IsoFolder"
    }

    if ($IsoFiles.Count -gt 1) {
        $IsoList = $IsoFiles |
            ForEach-Object {
                " - $($_.FullName)"
            }

        throw @"
Plusieurs ISO sont présentes dans le dossier ISO.

Conserve une seule ISO :
$($IsoList -join [Environment]::NewLine)
"@
    }

    $IsoPath = $IsoFiles[0].FullName

    Write-Success "ISO détectée : $IsoPath"

    # ========================================================================
    # ÉTAPE 3 : RECHERCHE DES PACKAGES
    # ========================================================================

    Write-Step "3. Vérification des packages de mise à jour"

    $UpdatePackages = @(
        Get-ChildItem `
            -LiteralPath $UpdatesFolder `
            -Recurse `
            -File |
            Where-Object {
                $_.Extension.ToLowerInvariant() -in @(".msu", ".cab")
            } |
            Sort-Object -Property @(
                # Ordre : SSU autonome, puis cumulative Windows (LCU), puis .NET, puis le reste.
                @{ Expression = { Get-UpdatePackageCategory -Package $_ } },
                "Name"
            )
    )

    # Une cumulative combinée contient sa SSU (SSU-*.cab), mais DISM l'applique dans
    # la même session que la LCU, avec l'ancienne pile de l'image : la LCU échoue alors
    # en 0x800F0823 (CBS_E_NEW_SERVICING_STACK_REQUIRED). La SSU est donc extraite
    # et installée séparément, avant la cumulative.
    $SsuFolder = Join-Path -Path $WorkFolder -ChildPath "SSU"

    $ExtractedSsuPackages = New-Object System.Collections.Generic.List[System.IO.FileInfo]

    foreach ($CumulativePackage in @(
        $UpdatePackages | Where-Object {
            $_.Extension -eq ".msu" -and (Get-UpdatePackageCategory -Package $_) -eq 1
        }
    )) {
        $SsuDestination = Join-Path -Path $SsuFolder -ChildPath $CumulativePackage.BaseName

        New-Item -Path $SsuDestination -ItemType Directory -Force | Out-Null

        Write-Info "Recherche d'une SSU dans : $($CumulativePackage.Name)"

        & "$env:SystemRoot\System32\expand.exe" "-F:SSU-*.cab" $CumulativePackage.FullName $SsuDestination | Out-Null

        if ($LASTEXITCODE -ne 0) {
            throw "Échec de l'extraction de la SSU depuis : $($CumulativePackage.Name)"
        }

        $FoundSsu = @(
            Get-ChildItem -LiteralPath $SsuDestination -Filter "SSU-*.cab" -File
        )

        foreach ($SsuCab in $FoundSsu) {
            $ExtractedSsuPackages.Add($SsuCab)
            Write-Success "SSU extraite : $($SsuCab.Name)"
        }
    }

    if ($ExtractedSsuPackages.Count -gt 0) {
        $UpdatePackages = @($ExtractedSsuPackages) + $UpdatePackages
    }

    # WinRE et WinPE ne reçoivent que la SSU et la cumulative Windows (pas .NET).
    $PePackages = @(
        $UpdatePackages | Where-Object {
            (Get-UpdatePackageCategory -Package $_) -in @(0, 1)
        }
    )

    if ($UpdatePackages.Count -eq 0) {
        throw @"
Aucun package MSU ou CAB trouvé dans :
$UpdatesFolder

Place les mises à jour Windows 10 version 1809 x64 dans ce dossier.
"@
    }

    Write-Host ""
    Write-Host "Packages à intégrer, dans cet ordre :" -ForegroundColor White

    $PackageNumber = 0

    foreach ($UpdatePackage in $UpdatePackages) {
        $PackageNumber++

        Write-Host (
            "  {0}. {1} ({2:N2} Mo)" -f
            $PackageNumber,
            $UpdatePackage.Name,
            ($UpdatePackage.Length / 1MB)
        )
    }

    # ========================================================================
    # ÉTAPE 4 : MONTAGE ET COPIE DE L'ISO
    # ========================================================================

    Write-Step "4. Montage et copie de l'ISO"

    $DiskImage = Mount-DiskImage `
        -ImagePath $IsoPath `
        -PassThru `
        -ErrorAction Stop

    $MountedIso = $true

    Start-Sleep -Seconds 3

    $Volume = $DiskImage |
        Get-Volume |
        Where-Object {
            $null -ne $_.DriveLetter
        } |
        Select-Object -First 1

    if ($null -eq $Volume) {
        throw "Impossible de déterminer la lettre du lecteur de l'ISO montée."
    }

    $IsoDriveLetter = "$($Volume.DriveLetter):\"

    Write-Success "ISO montée sur : $IsoDriveLetter"

    Copy-IsoContent `
        -SourceDrive $IsoDriveLetter `
        -Destination $MediaFolder

    Remove-ReadOnlyAttributes -Path $MediaFolder

    Dismount-DiskImage `
        -ImagePath $IsoPath `
        -ErrorAction Stop

    $MountedIso = $false

    Write-Success "ISO source démontée."

    # ========================================================================
    # ÉTAPE 5 : DÉTECTION DE INSTALL.WIM OU INSTALL.ESD
    # ========================================================================

    Write-Step "5. Détection de l'image Windows"

    $InstallWimPath = Join-Path `
        -Path $MediaFolder `
        -ChildPath "sources\install.wim"

    $InstallEsdPath = Join-Path `
        -Path $MediaFolder `
        -ChildPath "sources\install.esd"

    if (Test-Path -LiteralPath $InstallWimPath) {
        $SourceInstallImagePath = $InstallWimPath
        $SourceImageType = "WIM"
    }
    elseif (Test-Path -LiteralPath $InstallEsdPath) {
        $SourceInstallImagePath = $InstallEsdPath
        $SourceImageType = "ESD"
    }
    else {
        throw "Aucun fichier install.wim ou install.esd trouvé dans le dossier Sources."
    }

    Write-Success "Image détectée : $SourceInstallImagePath"
    Write-Info "Format détecté : $SourceImageType"

    # Sans -Index, Get-WindowsImage ne renvoie ni Architecture ni Version :
    # on relit chaque index pour obtenir les informations détaillées.
    $Images = @(
        Get-WindowsImage `
            -ImagePath $SourceInstallImagePath `
            -ErrorAction Stop |
            ForEach-Object {
                Get-WindowsImage `
                    -ImagePath $SourceInstallImagePath `
                    -Index $_.ImageIndex `
                    -ErrorAction Stop
            }
    )

    if ($Images.Count -eq 0) {
        throw "Aucune édition Windows détectée dans : $SourceInstallImagePath"
    }

    $TargetImage = Select-TargetWindowsImage `
        -Images $Images `
        -EditionPattern $TargetEditionPattern

    $SourceIndex = [int]$TargetImage.ImageIndex
    $SelectedImageName = [string]$TargetImage.ImageName

    Write-Success "Édition sélectionnée : $SelectedImageName"
    Write-Success "Index source : $SourceIndex"

    # ========================================================================
    # ÉTAPE 6 : EXPORT VERS INSTALL.WIM MONO-INDEX
    # ========================================================================

    Write-Step "6. Création de install.wim avec l'édition LTSC 2019"

    $OriginalExtension = [System.IO.Path]::GetExtension(
        $SourceInstallImagePath
    )

    $OriginalBackupPath = Join-Path `
        -Path $BackupFolder `
        -ChildPath "install-original$OriginalExtension"

    Copy-Item `
        -LiteralPath $SourceInstallImagePath `
        -Destination $OriginalBackupPath `
        -Force `
        -ErrorAction Stop

    Write-Success "Sauvegarde créée : $OriginalBackupPath"

    if (Test-Path -LiteralPath $InstallWimPath) {
        Remove-Item `
            -LiteralPath $InstallWimPath `
            -Force `
            -ErrorAction Stop
    }

    if (Test-Path -LiteralPath $InstallEsdPath) {
        Remove-Item `
            -LiteralPath $InstallEsdPath `
            -Force `
            -ErrorAction Stop
    }

    Export-ImageToWim `
        -SourceImage $OriginalBackupPath `
        -SourceIndex $SourceIndex `
        -DestinationImage $InstallWimPath `
        -DestinationName $SelectedImageName

    Write-Success "install.wim créé avec un seul index."

    # ========================================================================
    # ÉTAPE 7 : MONTAGE DE INSTALL.WIM
    # ========================================================================

    Write-Step "7. Montage de install.wim"

    Invoke-Dism -Arguments @(
        "/English",
        "/Mount-Image",
        "/ImageFile:$InstallWimPath",
        "/Index:1",
        "/MountDir:$MountFolder",
        "/CheckIntegrity",
        "/Optimize",
        "/ScratchDir:$ScratchFolder",
        "/LogPath:$DismLogPath"
    ) | Out-Null

    $MountedWindowsImage = $true

    Write-Success "Image montée dans : $MountFolder"

    # ========================================================================
    # ÉTAPE 8 : VÉRIFICATION DU BUILD
    # ========================================================================

    Write-Step "8. Vérification de Windows 10 LTSC 2019"

    # Lecture du build dans les métadonnées du WIM plutôt que par reg.exe load :
    # sous PowerShell 5.1, la sortie d'erreur de reg.exe devient une erreur
    # bloquante avec $ErrorActionPreference = "Stop".
    $MountedImageInfo = Get-WindowsImage `
        -ImagePath $InstallWimPath `
        -Index 1 `
        -ErrorAction Stop

    $ImageVersion = [version]$MountedImageInfo.Version
    $CurrentBuild = [string]$ImageVersion.Build

    Write-Info "Produit : $($MountedImageInfo.ImageName)"
    Write-Info "Build avant intégration : $($MountedImageInfo.Version)"

    if ($CurrentBuild -ne "17763") {
        throw @"
L'image ne correspond pas à Windows 10 version 1809.

Build attendu : 17763
Build détecté : $CurrentBuild

Le traitement est arrêté pour éviter d'intégrer des correctifs incompatibles.
"@
    }

    Write-Success "L'image correspond à Windows 10 version 1809 / LTSC 2019."

    # ========================================================================
    # ÉTAPE 8b : MISE À JOUR DE WINRE
    # ========================================================================

    Write-Step "8b. Mise à jour de WinRE (Winre.wim)"

    $WinRePathInImage = Join-Path `
        -Path $MountFolder `
        -ChildPath "Windows\System32\Recovery\Winre.wim"

    if (-not (Test-Path -LiteralPath $WinRePathInImage)) {
        Write-WarningMessage "Winre.wim introuvable dans l'image. Étape ignorée."
    }
    elseif ($PePackages.Count -eq 0) {
        Write-WarningMessage "Aucune cumulative Windows ou SSU à intégrer dans WinRE. Étape ignorée."
    }
    else {
        $WinReItem = Get-Item -LiteralPath $WinRePathInImage -Force
        $WinReAttributes = $WinReItem.Attributes

        $WinReWorkPath = Join-Path -Path $WinReFolder -ChildPath "winre.wim"
        $WinReOptimizedPath = Join-Path -Path $WinReFolder -ChildPath "winre-optimized.wim"

        Copy-Item `
            -LiteralPath $WinRePathInImage `
            -Destination $WinReWorkPath `
            -Force `
            -ErrorAction Stop

        (Get-Item -LiteralPath $WinReWorkPath -Force).Attributes = [System.IO.FileAttributes]::Normal

        Invoke-Dism -Arguments @(
            "/English",
            "/Mount-Image",
            "/ImageFile:$WinReWorkPath",
            "/Index:1",
            "/MountDir:$WinReMountFolder",
            "/ScratchDir:$ScratchFolder",
            "/LogPath:$DismLogPath"
        ) | Out-Null

        $MountedWinRE = $true

        # Avec une cumulative combinée SSU+LCU, WinRE peut refuser la partie LCU
        # (problème connu documenté par Microsoft) : les erreurs sont tolérées.
        Add-UpdatePackages `
            -MountDir $WinReMountFolder `
            -Packages $PePackages `
            -IgnoreError

        Invoke-ComponentCleanup -MountDir $WinReMountFolder

        Dismount-ImageCommit -MountDir $WinReMountFolder

        $MountedWinRE = $false

        $WinReName = [string](Get-WindowsImage -ImagePath $WinReWorkPath -Index 1).ImageName

        Export-ImageToWim `
            -SourceImage $WinReWorkPath `
            -SourceIndex 1 `
            -DestinationImage $WinReOptimizedPath `
            -DestinationName $WinReName

        # Winre.wim est masqué/système : Copy-Item ne peut pas l'écraser sans retirer ces attributs.
        $WinReItem.Attributes = [System.IO.FileAttributes]::Normal

        Copy-Item `
            -LiteralPath $WinReOptimizedPath `
            -Destination $WinRePathInImage `
            -Force `
            -ErrorAction Stop

        (Get-Item -LiteralPath $WinRePathInImage -Force).Attributes = $WinReAttributes

        Write-Success "WinRE mis à jour et réinjecté dans l'image."
    }

    # ========================================================================
    # ÉTAPE 9 : INJECTION DES PACKAGES
    # ========================================================================

    Write-Step "9. Injection des mises à jour"

    foreach ($UpdatePackage in $UpdatePackages) {
        Write-Host ""
        Write-Host "--------------------------------------------------------------------" -ForegroundColor Yellow
        Write-Host "Intégration : $($UpdatePackage.Name)" -ForegroundColor Yellow
        Write-Host "--------------------------------------------------------------------" -ForegroundColor Yellow

        Invoke-Dism -Arguments @(
            "/English",
            "/Image:$MountFolder",
            "/Add-Package",
            "/PackagePath:$($UpdatePackage.FullName)",
            "/ScratchDir:$ScratchFolder",
            "/LogPath:$DismLogPath"
        ) | Out-Null

        $IntegratedPackages.Add($UpdatePackage.FullName)

        Write-Success "Package intégré : $($UpdatePackage.Name)"
    }

    # ========================================================================
    # ÉTAPE 10 : NETTOYAGE DU MAGASIN DE COMPOSANTS
    # ========================================================================

    Write-Step "10. Nettoyage du magasin de composants"

    $CleanupExitCode = Invoke-Dism -Arguments @(
        "/English",
        "/Image:$MountFolder",
        "/Cleanup-Image",
        "/StartComponentCleanup",
        "/ScratchDir:$ScratchFolder",
        "/LogPath:$DismLogPath"
    ) -IgnoreError

    if ($CleanupExitCode -eq 0) {
        Write-Success "Nettoyage du magasin de composants terminé."
    }
    else {
        Write-WarningMessage "Le nettoyage n'a pas pu être entièrement exécuté."
    }

    # ========================================================================
    # ÉTAPE 11 : LISTE DES PACKAGES
    # ========================================================================

    Write-Step "11. Vérification des packages intégrés"

    $PackageOutput = & "$env:SystemRoot\System32\dism.exe" `
        "/English" `
        "/Image:$MountFolder" `
        "/Get-Packages" `
        "/Format:Table"

    $PackageListExitCode = $LASTEXITCODE

    $PackageOutput |
        Out-File `
            -FilePath $PackageListPath `
            -Encoding UTF8 `
            -Force

    if ($PackageListExitCode -ne 0) {
        throw "Impossible de récupérer la liste des packages."
    }

    Write-Success "Liste enregistrée dans : $PackageListPath"

    # ========================================================================
    # ÉTAPE 12 : VALIDATION ET DÉMONTAGE
    # ========================================================================

    Write-Step "12. Validation et démontage de install.wim"

    Dismount-ImageCommit -MountDir $MountFolder

    $MountedWindowsImage = $false

    Write-Success "Modifications enregistrées dans install.wim."

    # ========================================================================
    # ÉTAPE 13 : OPTIMISATION DE INSTALL.WIM
    # ========================================================================

    Write-Step "13. Optimisation de install.wim"

    $OptimizedWimPath = Join-Path `
        -Path (Split-Path -Parent $InstallWimPath) `
        -ChildPath "install-optimized.wim"

    if (Test-Path -LiteralPath $OptimizedWimPath) {
        Remove-Item `
            -LiteralPath $OptimizedWimPath `
            -Force `
            -ErrorAction Stop
    }

    Export-ImageToWim `
        -SourceImage $InstallWimPath `
        -SourceIndex 1 `
        -DestinationImage $OptimizedWimPath `
        -DestinationName $SelectedImageName

    Remove-Item `
        -LiteralPath $InstallWimPath `
        -Force `
        -ErrorAction Stop

    Rename-Item `
        -LiteralPath $OptimizedWimPath `
        -NewName "install.wim" `
        -ErrorAction Stop

    Write-Success "install.wim optimisé."

    # ========================================================================
    # ÉTAPE 14 : VÉRIFICATION FINALE
    # ========================================================================

    Write-Step "14. Vérification finale de install.wim"

    $FinalImages = @(
        Get-WindowsImage `
            -ImagePath $InstallWimPath `
            -ErrorAction Stop
    )

    if ($FinalImages.Count -ne 1) {
        throw "L'image finale devrait contenir exactement un index."
    }

    $FinalImage = Get-WindowsImage `
        -ImagePath $InstallWimPath `
        -Index 1 `
        -ErrorAction Stop

    $FinalWimFile = Get-Item -LiteralPath $InstallWimPath

    Write-Info "Nom final : $($FinalImage.ImageName)"
    Write-Info "Index final : $($FinalImage.ImageIndex)"
    Write-Info "Version finale : $($FinalImage.Version)"

    Write-Info (
        "Taille de install.wim : {0:N2} Go" -f
        ($FinalWimFile.Length / 1GB)
    )

    # ========================================================================
    # ÉTAPE 14b : MISE À JOUR DE BOOT.WIM
    # ========================================================================

    Write-Step "14b. Mise à jour de boot.wim (WinPE et Setup)"

    $BootWimPath = Join-Path `
        -Path $MediaFolder `
        -ChildPath "sources\boot.wim"

    if (-not (Test-Path -LiteralPath $BootWimPath)) {
        Write-WarningMessage "boot.wim introuvable. Étape ignorée."
    }
    elseif ($PePackages.Count -eq 0) {
        Write-WarningMessage "Aucune cumulative Windows ou SSU à intégrer dans boot.wim. Étape ignorée."
    }
    else {
        $BootImages = @(
            Get-WindowsImage `
                -ImagePath $BootWimPath `
                -ErrorAction Stop
        )

        foreach ($BootImage in $BootImages) {
            Write-Info "boot.wim index $($BootImage.ImageIndex) : $($BootImage.ImageName)"

            Invoke-Dism -Arguments @(
                "/English",
                "/Mount-Image",
                "/ImageFile:$BootWimPath",
                "/Index:$($BootImage.ImageIndex)",
                "/MountDir:$BootMountFolder",
                "/ScratchDir:$ScratchFolder",
                "/LogPath:$DismLogPath"
            ) | Out-Null

            $MountedBootImage = $true

            Add-UpdatePackages `
                -MountDir $BootMountFolder `
                -Packages $PePackages `
                -IgnoreError

            Invoke-ComponentCleanup -MountDir $BootMountFolder

            # L'index Setup contient setup.exe : on récupère la version à jour pour le média.
            $UpdatedSetupPath = Join-Path `
                -Path $BootMountFolder `
                -ChildPath "sources\setup.exe"

            if (Test-Path -LiteralPath $UpdatedSetupPath) {
                Copy-Item `
                    -LiteralPath $UpdatedSetupPath `
                    -Destination (Join-Path -Path $MediaFolder -ChildPath "sources\setup.exe") `
                    -Force `
                    -ErrorAction Stop

                Write-Success "setup.exe du média mis à jour."
            }

            Dismount-ImageCommit -MountDir $BootMountFolder

            $MountedBootImage = $false
        }

        $OptimizedBootWimPath = Join-Path `
            -Path (Split-Path -Parent $BootWimPath) `
            -ChildPath "boot-optimized.wim"

        if (Test-Path -LiteralPath $OptimizedBootWimPath) {
            Remove-Item `
                -LiteralPath $OptimizedBootWimPath `
                -Force `
                -ErrorAction Stop
        }

        foreach ($BootImage in $BootImages) {
            Invoke-Dism -Arguments @(
                "/English",
                "/Export-Image",
                "/SourceImageFile:$BootWimPath",
                "/SourceIndex:$($BootImage.ImageIndex)",
                "/DestinationImageFile:$OptimizedBootWimPath",
                "/DestinationName:$($BootImage.ImageName)",
                "/Compress:max",
                "/CheckIntegrity",
                "/ScratchDir:$ScratchFolder",
                "/LogPath:$DismLogPath"
            ) | Out-Null
        }

        Remove-Item `
            -LiteralPath $BootWimPath `
            -Force `
            -ErrorAction Stop

        Rename-Item `
            -LiteralPath $OptimizedBootWimPath `
            -NewName "boot.wim" `
            -ErrorAction Stop

        Write-Success "boot.wim mis à jour ($($BootImages.Count) index)."
    }

    # ========================================================================
    # ÉTAPE 15 : VÉRIFICATION DES FICHIERS DE DÉMARRAGE
    # ========================================================================

    Write-Step "15. Vérification des fichiers de démarrage"

    $BiosBootFile = Join-Path `
        -Path $MediaFolder `
        -ChildPath "boot\etfsboot.com"

    $UefiBootFile = Join-Path `
        -Path $MediaFolder `
        -ChildPath "efi\microsoft\boot\efisys.bin"

    if (-not (Test-Path -LiteralPath $BiosBootFile)) {
        throw "Fichier de démarrage BIOS introuvable : $BiosBootFile"
    }

    if (-not (Test-Path -LiteralPath $UefiBootFile)) {
        throw "Fichier de démarrage UEFI introuvable : $UefiBootFile"
    }

    Write-Success "Fichier BIOS détecté : $BiosBootFile"
    Write-Success "Fichier UEFI détecté : $UefiBootFile"

    # ========================================================================
    # ÉTAPE 16 : CRÉATION DE L'ISO
    # ========================================================================

    Write-Step "16. Création de l'ISO bootable"

    if (Test-Path -LiteralPath $OutputIsoPath) {
        Remove-Item `
            -LiteralPath $OutputIsoPath `
            -Force `
            -ErrorAction Stop
    }

    $VolumeLabel = "W10_LTSC2019"

    $OscdimgArguments = @(
        "-m",
        "-o",
        "-u2",
        "-udfver102",
        "-l$VolumeLabel",
        "-bootdata:2#p0,e,b$BiosBootFile#pEF,e,b$UefiBootFile",
        $MediaFolder,
        $OutputIsoPath
    )

    Write-Host ""
    Write-Host "$OscdimgPath $($OscdimgArguments -join ' ')" -ForegroundColor DarkGray
    Write-Host ""

    # Oscdimg écrit sa progression sur stderr : sous PowerShell 5.1, cela ne doit
    # pas être traité comme une erreur bloquante. Seul le code de retour compte.
    $PreviousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = "Continue"

    try {
        & $OscdimgPath @OscdimgArguments 2>&1 |
            ForEach-Object { Write-Host $_.ToString() }

        $OscdimgExitCode = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $PreviousErrorActionPreference
    }

    if ($OscdimgExitCode -ne 0) {
        throw "Échec de la création de l'ISO. Code Oscdimg : $OscdimgExitCode"
    }

    if (-not (Test-Path -LiteralPath $OutputIsoPath)) {
        throw "L'ISO de sortie n'a pas été créée."
    }

    $OutputIsoFile = Get-Item -LiteralPath $OutputIsoPath

    Write-Success "ISO créée : $OutputIsoPath"

    Write-Success (
        "Taille de l'ISO : {0:N2} Go" -f
        ($OutputIsoFile.Length / 1GB)
    )

    # ========================================================================
    # ÉTAPE 17 : CALCUL DU HASH
    # ========================================================================

    Write-Step "17. Calcul du hash SHA-256"

    $IsoHash = Get-FileHash `
        -LiteralPath $OutputIsoPath `
        -Algorithm SHA256

    $HashFilePath = "$OutputIsoPath.sha256.txt"

    @(
        "Fichier : $OutputIsoPath"
        "SHA256  : $($IsoHash.Hash)"
        "Date    : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    ) |
        Out-File `
            -FilePath $HashFilePath `
            -Encoding ASCII `
            -Force

    Write-Success "SHA-256 : $($IsoHash.Hash)"
    Write-Success "Hash enregistré dans : $HashFilePath"

    # ========================================================================
    # RÉSULTAT
    # ========================================================================

    Write-Step "TRAITEMENT TERMINÉ AVEC SUCCÈS"

    Write-Host "ISO source :" -ForegroundColor White
    Write-Host "  $IsoPath"
    Write-Host ""

    Write-Host "Édition intégrée :" -ForegroundColor White
    Write-Host "  $SelectedImageName"
    Write-Host ""

    Write-Host "Packages intégrés :" -ForegroundColor White

    foreach ($IntegratedPackage in $IntegratedPackages) {
        Write-Host "  - $IntegratedPackage"
    }

    Write-Host ""
    Write-Host "ISO finale :" -ForegroundColor White
    Write-Host "  $OutputIsoPath" -ForegroundColor Green
    Write-Host ""

    Write-Host "Journal principal :" -ForegroundColor White
    Write-Host "  $TranscriptPath"
    Write-Host ""

    Write-Host "Journal DISM :" -ForegroundColor White
    Write-Host "  $DismLogPath"
    Write-Host ""

    Write-Host "SHA-256 :" -ForegroundColor White
    Write-Host "  $($IsoHash.Hash)"
    Write-Host ""

    Write-WarningMessage "Teste l'ISO dans une machine virtuelle avant utilisation en production."
}
catch {
    Write-Host ""
    Write-Host "======================================================================" -ForegroundColor Red
    Write-Host "ERREUR" -ForegroundColor Red
    Write-Host "======================================================================" -ForegroundColor Red

    Write-Host $_.Exception.Message -ForegroundColor Red

    if (-not [string]::IsNullOrWhiteSpace($_.ScriptStackTrace)) {
        Write-Host ""
        Write-Host "Détail :" -ForegroundColor DarkRed
        Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
    }

    throw
}
finally {
    Write-Host ""
    Write-Info "Exécution du nettoyage final."

    # WinRE est stocké dans install.wim : on le démonte avant celui-ci.
    if ($MountedWinRE) {
        Dismount-ImageDiscard -MountDir $WinReMountFolder
    }

    if ($MountedBootImage) {
        Dismount-ImageDiscard -MountDir $BootMountFolder
    }

    if (
        $MountedWindowsImage -and
        (Test-Path -LiteralPath $MountFolder)
    ) {
        Dismount-ImageDiscard -MountDir $MountFolder
    }

    if (
        $MountedIso -and
        -not [string]::IsNullOrWhiteSpace($IsoPath)
    ) {
        Write-WarningMessage "L'ISO source est encore montée. Tentative de démontage."

        try {
            Dismount-DiskImage `
                -ImagePath $IsoPath `
                -ErrorAction SilentlyContinue
        }
        catch {
            Write-WarningMessage "Le démontage de l'ISO source a échoué."
        }
    }

    if ($TranscriptStarted) {
        Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
    }
}