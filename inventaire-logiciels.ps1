$DossierSortie = "$env:USERPROFILE\Documents\InventaireLogiciels"
if (-not (Test-Path $DossierSortie)) {
    New-Item -ItemType Directory -Path $DossierSortie -Force | Out-Null
}

$Date = Get-Date
$Horodatage = $Date.ToString("yyyy-MM-dd_HH-mm")
$FichierCsv  = Join-Path $DossierSortie "inventaire_$Horodatage.csv"
$FichierHtml = Join-Path $DossierSortie "inventaire_$Horodatage.html"
$DateAffichee = $Date.ToString("dd/MM/yyyy HH:mm")

Write-Host "Collecte des logiciels installes..." -ForegroundColor Cyan

$clesUninstall = @(
    @{ Chemin = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*";              Source = "Systeme (64 bits)" },
    @{ Chemin = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"; Source = "Systeme (32 bits)" },
    @{ Chemin = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*";              Source = "Utilisateur" }
)

$logiciels = @()

foreach ($cle in $clesUninstall) {
    Write-Host "  - Registre : $($cle.Source)" -ForegroundColor Gray
    $entrees = Get-ItemProperty -Path $cle.Chemin -ErrorAction SilentlyContinue
    foreach ($e in $entrees) {
        if (-not $e.DisplayName) { continue }
        if ($e.SystemComponent -eq 1) { continue }
        if ($e.ParentKeyName) { continue }

        $taille = if ($e.EstimatedSize) { [math]::Round($e.EstimatedSize / 1024, 1) } else { $null }
        $dateInstall = ""
        if ($e.InstallDate -and $e.InstallDate -match '^\d{8}$') {
            try {
                $d = [datetime]::ParseExact($e.InstallDate, "yyyyMMdd", $null)
                $dateInstall = $d.ToString("yyyy-MM-dd")
            } catch {}
        }

        $logiciels += [PSCustomObject]@{
            Nom          = $e.DisplayName
            Version      = if ($e.DisplayVersion) { $e.DisplayVersion } else { "" }
            Editeur      = if ($e.Publisher) { $e.Publisher } else { "" }
            DateInstall  = $dateInstall
            TailleMo     = $taille
            Source       = $cle.Source
            Emplacement  = if ($e.InstallLocation) { $e.InstallLocation } else { "" }
            Desinstaller = if ($e.UninstallString) { $e.UninstallString } else { "" }
        }
    }
}

Write-Host "  - Applications Microsoft Store" -ForegroundColor Gray
try {
    $appx = Get-AppxPackage -ErrorAction Stop | Where-Object { -not $_.IsFramework -and -not $_.IsResourcePackage }
    foreach ($a in $appx) {
        $logiciels += [PSCustomObject]@{
            Nom          = $a.Name
            Version      = $a.Version
            Editeur      = ($a.Publisher -replace '^CN=', '' -split ',')[0]
            DateInstall  = ""
            TailleMo     = $null
            Source       = "Microsoft Store"
            Emplacement  = $a.InstallLocation
            Desinstaller = ""
        }
    }
} catch {
    Write-Host "    (impossible de lister les applis Store)" -ForegroundColor Yellow
}

$logiciels = $logiciels | Sort-Object Nom, Version -Unique | Sort-Object Nom

Write-Host ""
Write-Host "$($logiciels.Count) logiciels detectes." -ForegroundColor Green

Write-Host "Export CSV..." -ForegroundColor Cyan
$logiciels | Export-Csv -Path $FichierCsv -Delimiter ';' -Encoding UTF8 -NoTypeInformation

Write-Host "Generation du rapport HTML..." -ForegroundColor Cyan

$stats = @{
    Total          = $logiciels.Count
    Systeme64      = ($logiciels | Where-Object { $_.Source -eq "Systeme (64 bits)" }).Count
    Systeme32      = ($logiciels | Where-Object { $_.Source -eq "Systeme (32 bits)" }).Count
    Utilisateur    = ($logiciels | Where-Object { $_.Source -eq "Utilisateur" }).Count
    Store          = ($logiciels | Where-Object { $_.Source -eq "Microsoft Store" }).Count
    TailleTotaleGo = [math]::Round((($logiciels | Where-Object { $_.TailleMo } | Measure-Object -Property TailleMo -Sum).Sum) / 1024, 1)
}

$lignesHtml = ""
foreach ($l in $logiciels) {
    $nom     = ($l.Nom     -replace '<','&lt;' -replace '>','&gt;')
    $version = ($l.Version -replace '<','&lt;' -replace '>','&gt;')
    $editeur = ($l.Editeur -replace '<','&lt;' -replace '>','&gt;')
    $taille  = if ($l.TailleMo) { "$($l.TailleMo) Mo" } else { "-" }
    $date    = if ($l.DateInstall) { $l.DateInstall } else { "-" }
    $source  = $l.Source

    $classeSrc = switch ($source) {
        "Systeme (64 bits)" { "src64" }
        "Systeme (32 bits)" { "src32" }
        "Utilisateur"       { "srcuser" }
        "Microsoft Store"   { "srcstore" }
        default             { "" }
    }

    $lignesHtml += "<tr><td><strong>$nom</strong></td><td class='mono'>$version</td><td>$editeur</td><td class='num'>$taille</td><td class='num'>$date</td><td><span class='badge $classeSrc'>$source</span></td></tr>`n"
}

$pcNom = $env:COMPUTERNAME
$utilisateur = $env:USERNAME

$html = @"
<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<title>Inventaire logiciels - $pcNom</title>
<style>
  * { box-sizing: border-box; }
  body { font-family: 'Segoe UI', system-ui, sans-serif; background: #f4f6f8; color: #222; margin: 0; padding: 32px; }
  .wrap { max-width: 1400px; margin: 0 auto; }
  header { background: linear-gradient(135deg, #0ea5e9, #6366f1); color: white; padding: 32px; border-radius: 16px; margin-bottom: 24px; box-shadow: 0 4px 20px rgba(14,165,233,0.25); }
  header h1 { margin: 0 0 8px 0; font-size: 28px; }
  header .meta { opacity: 0.9; font-size: 14px; }
  .stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 16px; margin-bottom: 24px; }
  .stat { background: white; border-radius: 12px; padding: 18px; box-shadow: 0 2px 8px rgba(0,0,0,0.05); }
  .stat-label { font-size: 11px; text-transform: uppercase; color: #6b7280; letter-spacing: 0.5px; }
  .stat-value { font-size: 28px; font-weight: 700; margin-top: 4px; color: #111; }
  .toolbar { background: white; padding: 16px; border-radius: 12px; margin-bottom: 16px; display: flex; gap: 12px; align-items: center; box-shadow: 0 2px 8px rgba(0,0,0,0.05); flex-wrap: wrap; }
  .toolbar input { flex: 1; min-width: 240px; padding: 10px 14px; border: 1px solid #d1d5db; border-radius: 8px; font-size: 14px; }
  .toolbar select { padding: 10px 14px; border: 1px solid #d1d5db; border-radius: 8px; font-size: 14px; background: white; }
  .count { color: #6b7280; font-size: 13px; }
  .table-wrap { background: white; border-radius: 12px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.05); }
  table { width: 100%; border-collapse: collapse; }
  th { background: #f3f4f6; padding: 12px 14px; text-align: left; font-size: 13px; color: #374151; border-bottom: 1px solid #e5e7eb; cursor: pointer; user-select: none; position: sticky; top: 0; }
  th:hover { background: #e5e7eb; }
  td { padding: 10px 14px; border-bottom: 1px solid #f3f4f6; font-size: 14px; }
  tr:hover td { background: #fafbfc; }
  .mono { font-family: 'Consolas', 'Courier New', monospace; color: #555; font-size: 13px; }
  .num { white-space: nowrap; }
  .badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: 600; }
  .src64 { background: #dbeafe; color: #1e40af; }
  .src32 { background: #fef3c7; color: #92400e; }
  .srcuser { background: #dcfce7; color: #166534; }
  .srcstore { background: #fce7f3; color: #9f1239; }
  footer { text-align: center; padding: 24px; color: #9ca3af; font-size: 12px; }
</style>
</head>
<body>
<div class="wrap">
  <header>
    <h1>Inventaire des logiciels</h1>
    <div class="meta">$pcNom - $utilisateur - $DateAffichee</div>
  </header>

  <div class="stats">
    <div class="stat"><div class="stat-label">Total</div><div class="stat-value">$($stats.Total)</div></div>
    <div class="stat"><div class="stat-label">Systeme 64 bits</div><div class="stat-value">$($stats.Systeme64)</div></div>
    <div class="stat"><div class="stat-label">Systeme 32 bits</div><div class="stat-value">$($stats.Systeme32)</div></div>
    <div class="stat"><div class="stat-label">Utilisateur</div><div class="stat-value">$($stats.Utilisateur)</div></div>
    <div class="stat"><div class="stat-label">Microsoft Store</div><div class="stat-value">$($stats.Store)</div></div>
    <div class="stat"><div class="stat-label">Taille totale</div><div class="stat-value">$($stats.TailleTotaleGo) Go</div></div>
  </div>

  <div class="toolbar">
    <input type="text" id="filtre" placeholder="Rechercher un logiciel, un editeur...">
    <select id="sourceFiltre">
      <option value="">Toutes les sources</option>
      <option value="Systeme (64 bits)">Systeme (64 bits)</option>
      <option value="Systeme (32 bits)">Systeme (32 bits)</option>
      <option value="Utilisateur">Utilisateur</option>
      <option value="Microsoft Store">Microsoft Store</option>
    </select>
    <div class="count" id="compteur"></div>
  </div>

  <div class="table-wrap">
    <table id="tbl">
      <thead>
        <tr>
          <th data-col="0">Nom</th>
          <th data-col="1">Version</th>
          <th data-col="2">Editeur</th>
          <th data-col="3">Taille</th>
          <th data-col="4">Installe le</th>
          <th data-col="5">Source</th>
        </tr>
      </thead>
      <tbody>
$lignesHtml
      </tbody>
    </table>
  </div>

  <footer>Inventaire genere automatiquement - $DateAffichee - CSV : $FichierCsv</footer>
</div>

<script>
  const input = document.getElementById('filtre');
  const sourceSel = document.getElementById('sourceFiltre');
  const compteur = document.getElementById('compteur');
  const lignes = Array.from(document.querySelectorAll('#tbl tbody tr'));

  function filtrer() {
    const q = input.value.toLowerCase();
    const src = sourceSel.value;
    let visibles = 0;
    for (const tr of lignes) {
      const texte = tr.innerText.toLowerCase();
      const source = tr.children[5].innerText;
      const match = (q === '' || texte.includes(q)) && (src === '' || source === src);
      tr.style.display = match ? '' : 'none';
      if (match) visibles++;
    }
    compteur.textContent = visibles + ' / ' + lignes.length + ' logiciels affiches';
  }

  input.addEventListener('input', filtrer);
  sourceSel.addEventListener('change', filtrer);
  filtrer();

  document.querySelectorAll('th').forEach(th => {
    let asc = true;
    th.addEventListener('click', () => {
      const col = parseInt(th.dataset.col);
      const tbody = document.querySelector('#tbl tbody');
      const rows = Array.from(tbody.querySelectorAll('tr'));
      rows.sort((a, b) => {
        const av = a.children[col].innerText.trim().toLowerCase();
        const bv = b.children[col].innerText.trim().toLowerCase();
        const an = parseFloat(av);
        const bn = parseFloat(bv);
        if (!isNaN(an) && !isNaN(bn)) return asc ? an - bn : bn - an;
        return asc ? av.localeCompare(bv) : bv.localeCompare(av);
      });
      asc = !asc;
      rows.forEach(r => tbody.appendChild(r));
    });
  });
</script>
</body>
</html>
"@

$html | Out-File -FilePath $FichierHtml -Encoding UTF8

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "   Inventaire genere avec succes" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host "   CSV  : $FichierCsv" -ForegroundColor Yellow
Write-Host "   HTML : $FichierHtml" -ForegroundColor Yellow
Write-Host ""

Start-Process $FichierHtml
