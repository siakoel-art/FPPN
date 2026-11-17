# Mesure, pour une liste de diapositives, la hauteur réelle de chaque forme (texte rendu par PowerPoint) et les chevauchements.
# Utilisation : powershell -File mesure_v7.ps1 <fichier.pptx> <liste_diapos.txt> <sortie.json>
param([string]$f, [string]$liste, [string]$sortie)
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($f, $true, $false, $false)
$nums = (Get-Content $liste) -split '\s+' | Where-Object { $_ -ne '' } | ForEach-Object { [int]$_ }
$res = @()
foreach ($n in $nums) {
  $s = $pres.Slides.Item($n)
  foreach ($sh in $s.Shapes) {
    if ($sh.Width -gt 900 -and $sh.Height -gt 500) { continue }
    $bas = $sh.Top + $sh.Height
    $txt = ''
    if ($sh.HasTextFrame -and $sh.TextFrame2.HasText) {
      $tf = $sh.TextFrame2
      $bas = [Math]::Max($sh.Top + $tf.MarginTop + $tf.TextRange.BoundHeight + $tf.MarginBottom, $sh.Top)
      $txt = $tf.TextRange.Text
      if ($txt.Length -gt 40) { $txt = $txt.Substring(0, 40) }
    }
    $res += [pscustomobject]@{ n = $n; nom = $sh.Name; haut = [Math]::Round($sh.Top, 1); bas = [Math]::Round($bas, 1);
      cadre = [Math]::Round($sh.Top + $sh.Height, 1); g = [Math]::Round($sh.Left, 1); d = [Math]::Round($sh.Left + $sh.Width, 1); t = $txt }
  }
}
$pres.Close()
$res | ConvertTo-Json -Depth 3 | Out-File -Encoding utf8 $sortie
