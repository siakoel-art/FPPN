# FPPN - Inventaire de la dataroom et extraction des nouveaux fichiers (Adenka, 09/11/2026)
#
# Ce que fait le script :
#   1. il inventorie tous les fichiers des dossiers indiqués (chemin, taille, date, empreinte SHA-256) ;
#   2. il compare chaque fichier à la liste des fichiers déjà reçus par Adenka (FPPN_fichiers_deja_recus.csv),
#      sur le contenu exact (empreinte), puis sur le nom ;
#   3. il copie les fichiers nouveaux ou modifiés dans un nouveau dossier, en gardant les sous-dossiers d'origine ;
#   4. il écrit l'inventaire détaillé (CSV lisible dans Excel), le journal des fichiers copiés et un résumé.
# Aucun fichier d'origine n'est modifié ni déplacé.
#
# Mode d'emploi :
#   - placer ce script et FPPN_fichiers_deja_recus.csv dans le même dossier ;
#   - clic droit sur le script > « Exécuter avec PowerShell »,
#     ou dans PowerShell :  powershell -ExecutionPolicy Bypass -File .\FPPN_inventaire_dataroom.ps1

param(
    [string[]]$Dossiers = @("D:\05-Adenka Audit Juridique et Fiscal", "D:\03-Adenka-Audit"),
    [string]$Reference = (Join-Path $PSScriptRoot "FPPN_fichiers_deja_recus.csv"),
    [string]$Sortie = "D:\FPPN_Nouveaux_fichiers_20261109"
)

$ErrorActionPreference = "Stop"
function Long([string]$p) { if ($p.StartsWith("\\?\")) { $p } else { "\\?\" + $p } }
function Empreinte([string]$p) {
    $flux = [System.IO.File]::OpenRead((Long $p))
    try { $algo = [System.Security.Cryptography.SHA256]::Create(); ([System.BitConverter]::ToString($algo.ComputeHash($flux)) -replace "-", "").ToLower() }
    finally { $flux.Dispose() }
}

foreach ($d in $Dossiers) { if (-not (Test-Path -LiteralPath $d)) { throw "Dossier introuvable : $d" } }
if (-not (Test-Path -LiteralPath $Reference)) { throw "Liste de référence introuvable : $Reference" }

Write-Host "Lecture de la liste des fichiers déjà reçus..."
$ref = Import-Csv -LiteralPath $Reference -Delimiter ";" -Encoding UTF8
$parSha = @{}
$parNom = @{}
foreach ($r in $ref) {
    $s = $r.sha256.ToLower()
    if (-not $parSha.ContainsKey($s)) { $parSha[$s] = $r }
    $n = $r.nom_fichier.ToLower()
    if (-not $parNom.ContainsKey($n)) { $parNom[$n] = $r }
}

New-Item -ItemType Directory -Force -Path $Sortie | Out-Null
$inventaire = New-Object System.Collections.Generic.List[object]
$vus = @{}
$copies = 0

foreach ($racine in $Dossiers) {
    $nomRacine = Split-Path $racine -Leaf
    $fichiers = Get-ChildItem -LiteralPath $racine -Recurse -File -Force -ErrorAction SilentlyContinue
    $total = $fichiers.Count
    $k = 0
    foreach ($f in $fichiers) {
        $k++
        Write-Progress -Activity "Inventaire de $nomRacine" -Status "$k / $total : $($f.Name)" -PercentComplete ([int](100 * $k / [math]::Max($total, 1)))
        $rel = $f.FullName.Substring($racine.TrimEnd("\").Length).TrimStart("\")
        $sha = ""
        $erreur = ""
        try { $sha = Empreinte $f.FullName } catch { $erreur = "Lecture impossible : " + $_.Exception.Message }
        $statut = ""; $envoi = ""; $cheminRef = ""; $copie = ""
        if ($erreur) {
            $statut = "Illisible"
        } elseif ($parSha.ContainsKey($sha)) {
            $statut = "Déjà reçu (contenu identique)"; $envoi = $parSha[$sha].envoi; $cheminRef = $parSha[$sha].chemin_origine
        } elseif ($parNom.ContainsKey($f.Name.ToLower())) {
            $statut = "Même nom qu'un fichier déjà reçu, contenu différent"; $envoi = $parNom[$f.Name.ToLower()].envoi; $cheminRef = $parNom[$f.Name.ToLower()].chemin_origine
        } else {
            $statut = "Nouveau"
        }
        if ($sha -and $vus.ContainsKey($sha)) {
            $doublon = "Doublon de : " + $vus[$sha]
        } else {
            $doublon = ""
            if ($sha) { $vus[$sha] = "$nomRacine\$rel" }
        }
        if (($statut -eq "Nouveau" -or $statut -like "Même nom*") -and -not $doublon) {
            $cible = Join-Path (Join-Path $Sortie $nomRacine) $rel
            [System.IO.Directory]::CreateDirectory((Long (Split-Path $cible -Parent))) | Out-Null
            [System.IO.File]::Copy((Long $f.FullName), (Long $cible), $true)
            $copie = "$nomRacine\$rel"
            $copies++
        }
        $inventaire.Add([pscustomobject]@{
            dossier_dataroom    = $nomRacine
            sous_dossier        = (Split-Path $rel -Parent)
            fichier             = $f.Name
            extension           = $f.Extension.ToLower()
            taille_octets       = $f.Length
            date_modification   = $f.LastWriteTime.ToString("dd/MM/yyyy HH:mm")
            sha256              = $sha
            statut              = $statut
            deja_recu_dans      = $envoi
            chemin_deja_recu    = $cheminRef
            doublon_dataroom    = $doublon
            copie_dans_nouveaux = $copie
            remarque            = $erreur
        })
    }
}

$inv = Join-Path $Sortie "FPPN_inventaire_dataroom_20261109.csv"
$inventaire | Export-Csv -LiteralPath $inv -Delimiter ";" -NoTypeInformation -Encoding UTF8
$journal = Join-Path $Sortie "journal_nouveaux_fichiers.csv"
$inventaire | Where-Object { $_.copie_dans_nouveaux } | Select-Object @{n = "chemin_relatif"; e = { $_.copie_dans_nouveaux } }, taille_octets, sha256 |
    Export-Csv -LiteralPath $journal -Delimiter ";" -NoTypeInformation -Encoding UTF8

$res = Join-Path $Sortie "resume.txt"
$lignes = @("Inventaire de la dataroom FPPN du " + (Get-Date -Format "dd/MM/yyyy HH:mm"), "")
foreach ($racine in $Dossiers) {
    $n = Split-Path $racine -Leaf
    $sel = $inventaire | Where-Object { $_.dossier_dataroom -eq $n }
    $lignes += "$n : $($sel.Count) fichiers, $([math]::Round((($sel | Measure-Object taille_octets -Sum).Sum) / 1MB, 1)) Mo"
    foreach ($g in ($sel | Group-Object statut)) { $lignes += "   - $($g.Name) : $($g.Count)" }
    $lignes += "   - doublons internes : $(($sel | Where-Object { $_.doublon_dataroom }).Count)"
}
$lignes += "", "Fichiers copiés dans $Sortie : $copies"
$lignes | Set-Content -LiteralPath $res -Encoding UTF8
$lignes | ForEach-Object { Write-Host $_ }
Write-Host ""
Write-Host "Inventaire : $inv"
Write-Host "À transmettre à Adenka : le dossier $Sortie (compressé, en plusieurs parties s'il dépasse la taille autorisée)."
