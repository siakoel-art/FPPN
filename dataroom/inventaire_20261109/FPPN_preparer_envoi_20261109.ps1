# FPPN - Préparation de l'envoi à Adenka des fichiers nouveaux ou modifiés de la dataroom (09/11/2026)
#
# Ce que fait le script :
#   1. il lit la liste FPPN_a_transmettre_20261109.csv (fichiers nouveaux ou modifiés, priorités P1 et P2) ;
#   2. il range ces fichiers dans des archives .zip de 24 Mo au plus, lot par lot, en gardant les sous-dossiers ;
#   3. un fichier de plus de 24 Mo est mis seul dans une archive, puis cette archive est coupée en morceaux de 24 Mo
#      (fichiers « ..._morceau01sur03.zip », etc.). Ces morceaux ne s'ouvrent pas seuls : Adenka les recolle ;
#   4. il écrit contenu_envoi.csv (archive, fichier, taille, empreinte) et signale les fichiers introuvables.
# Aucun fichier d'origine n'est modifié ni déplacé.
#
# Lots :
#   Lot A : P1, tableurs, textes et présentations (environ 420 Mo avant compression) : à envoyer en premier
#   Lot B : P1, PDF et images (environ 890 Mo)
#   Lot C : P2, relevés bancaires, budgets, CNPS, CNAM, ITS, paie (environ 1 480 Mo)
#
# Mode d'emploi :
#   - placer ce script et FPPN_a_transmettre_20261109.csv dans le même dossier (par exemple D:\) ;
#   - dans PowerShell :  powershell -ExecutionPolicy Bypass -File "D:\FPPN_preparer_envoi_20261109.ps1"
#   - pour un seul lot :  powershell -ExecutionPolicy Bypass -File "D:\FPPN_preparer_envoi_20261109.ps1" -Lots "Lot A"

param(
    [string]$Liste = (Join-Path $PSScriptRoot "FPPN_a_transmettre_20261109.csv"),
    [string]$Sortie = "D:\FPPN_Envoi_20261109",
    [string[]]$Lots = @("Lot A", "Lot B", "Lot C"),
    [int]$TailleMo = 24
)

$ErrorActionPreference = "Stop"
$Racines = @{
    "03-Adenka-Audit"                     = "D:\03-Adenka-Audit"
    "05-Adenka Audit Juridique et Fiscal" = "D:\05-Adenka Audit Juridique et Fiscal"
}
Add-Type -AssemblyName System.IO.Compression
$limite = [int64]$TailleMo * 1MB

function Long([string]$p) { if ($p.StartsWith("\\?\")) { $p } else { "\\?\" + $p } }

function Ecrire-Zip([string]$chemin, $elements) {
    $fs = [System.IO.File]::Open((Long $chemin), [System.IO.FileMode]::Create)
    $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($e in $elements) {
            $entree = $zip.CreateEntry($e.Entree, [System.IO.Compression.CompressionLevel]::Optimal)
            $src = [System.IO.File]::OpenRead((Long $e.Source))
            $dst = $entree.Open()
            try { $src.CopyTo($dst) } finally { $dst.Dispose(); $src.Dispose() }
        }
    } finally { $zip.Dispose(); $fs.Dispose() }
}

function Couper([string]$chemin) {
    $taille = (New-Object System.IO.FileInfo($chemin)).Length
    if ($taille -le $limite) { return @($chemin) }
    $nb = [int][math]::Ceiling($taille / $limite)
    $base = $chemin.Substring(0, $chemin.Length - 4)
    $parts = @()
    $in = [System.IO.File]::OpenRead($chemin)
    $buf = New-Object byte[] (1MB)
    try {
        for ($i = 1; $i -le $nb; $i++) {
            $part = "{0}_morceau{1:D2}sur{2:D2}.zip" -f $base, $i, $nb
            $out = [System.IO.File]::Create($part)
            $ecrit = 0
            try {
                while ($ecrit -lt $limite) {
                    $n = $in.Read($buf, 0, [int][math]::Min($buf.Length, $limite - $ecrit))
                    if ($n -le 0) { break }
                    $out.Write($buf, 0, $n); $ecrit += $n
                }
            } finally { $out.Dispose() }
            $parts += $part
        }
    } finally { $in.Dispose() }
    Remove-Item -LiteralPath $chemin
    return $parts
}

if (-not (Test-Path -LiteralPath $Liste)) { throw "Liste introuvable : $Liste" }
foreach ($k in $Racines.Keys) { if (-not (Test-Path -LiteralPath $Racines[$k])) { throw "Dossier introuvable : $($Racines[$k])" } }
New-Item -ItemType Directory -Force -Path $Sortie | Out-Null

$liste = Import-Csv -LiteralPath $Liste -Delimiter ";" -Encoding UTF8
$contenu = New-Object System.Collections.Generic.List[object]
$manquants = New-Object System.Collections.Generic.List[string]

foreach ($lot in $Lots) {
    $fichiers = @($liste | Where-Object { $_.lot -eq $lot })
    if ($fichiers.Count -eq 0) { continue }
    $code = $lot -replace " ", ""
    $num = 0; $groupe = @(); $cumul = 0; $k = 0
    $vider = {
        if ($groupe.Count -gt 0) {
            $script:num++
            $zipNom = Join-Path $Sortie ("FPPN_{0}_{1:D3}.zip" -f $code, $script:num)
            Ecrire-Zip $zipNom $groupe
            $noms = Couper $zipNom
            foreach ($g in $groupe) {
                $contenu.Add([pscustomobject]@{ lot = $lot; archive = (($noms | ForEach-Object { Split-Path $_ -Leaf }) -join " + ");
                    fichier = $g.Entree; taille_octets = $g.Taille; sha256 = $g.Sha })
            }
        }
    }
    foreach ($f in $fichiers) {
        $k++
        Write-Progress -Activity "Préparation du $lot" -Status "$k / $($fichiers.Count) : $($f.chemin_relatif)" -PercentComplete ([int](100 * $k / $fichiers.Count))
        $src = Join-Path $Racines[$f.dossier_dataroom] $f.chemin_relatif
        if (-not [System.IO.File]::Exists((Long $src))) { $manquants.Add("$($f.dossier_dataroom)\$($f.chemin_relatif)"); continue }
        $t = [int64]$f.taille_octets
        $el = [pscustomobject]@{ Source = $src; Entree = ("$($f.dossier_dataroom)\$($f.chemin_relatif)" -replace "\\", "/"); Taille = $t; Sha = $f.sha256 }
        if ($t -gt $limite) {
            $sauve = $groupe; $groupe = @($el); . $vider; $groupe = $sauve
        } else {
            if ($cumul + $t -gt $limite) { . $vider; $groupe = @(); $cumul = 0 }
            $groupe += $el; $cumul += $t
        }
    }
    . $vider; $groupe = @(); $cumul = 0
    Write-Host "$lot : $($fichiers.Count) fichiers, $num archive(s) avant découpage"
}

$contenu | Export-Csv -LiteralPath (Join-Path $Sortie "contenu_envoi.csv") -Delimiter ";" -NoTypeInformation -Encoding UTF8
if ($manquants.Count -gt 0) {
    $manquants | Set-Content -LiteralPath (Join-Path $Sortie "fichiers_introuvables.txt") -Encoding UTF8
    Write-Host "Attention : $($manquants.Count) fichier(s) introuvable(s), voir fichiers_introuvables.txt"
}
$tous = Get-ChildItem -LiteralPath $Sortie -Filter *.zip
Write-Host ""
Write-Host "Archives prêtes dans $Sortie : $($tous.Count) fichiers, $([math]::Round((($tous | Measure-Object Length -Sum).Sum) / 1MB, 1)) Mo"
Write-Host "À envoyer : toutes les archives .zip du lot, puis contenu_envoi.csv."
