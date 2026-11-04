<#
.SYNOPSIS
    Crée les cinq catégories de priorité (avec leur couleur) dans les
    boîtes aux lettres ciblées, via Microsoft Graph PowerShell.

.DESCRIPTION
    Catégories : « Priorité · Urgent » (rouge), « Priorité · P1 » (orange),
    « Priorité · P2 » (jaune), « Priorité · P3 » (vert), « CONFIDENTIEL » (mauve).
    Les noms sont réservés au système de priorité : les catégories
    personnelles des employés ne sont jamais modifiées.
    - Une catégorie existante n'est pas dupliquée ; sa couleur est
      corrigée au besoin.
    - Idempotent : peut être relancé sans effet indésirable.
    - -WhatIf affiche ce qui serait fait sans rien modifier.

    Autorisations :
    - Votre propre boîte : connexion déléguée avec MailboxSettings.ReadWrite.
    - Boîtes d'autres utilisateurs : application inscrite dans Entra ID
      avec l'autorisation d'application MailboxSettings.ReadWrite
      (et User.Read.All pour -ToutesLesBoites), consentement administrateur.
      Se connecter avant de lancer le script (voir GUIDE-TI.md).

    Module requis : Microsoft.Graph.Users
      Install-Module Microsoft.Graph.Users -Scope CurrentUser

.PARAMETER Boites
    Adresses des boîtes ciblées.

.PARAMETER ToutesLesBoites
    Cible toutes les boîtes des utilisateurs membres actifs.

.EXAMPLE
    .\Creer-Categories.ps1 -Boites eric@bonneville.ca -WhatIf
.EXAMPLE
    .\Creer-Categories.ps1 -ToutesLesBoites
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Liste')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Liste')]
    [string[]]$Boites,

    [Parameter(Mandatory, ParameterSetName = 'Toutes')]
    [switch]$ToutesLesBoites
)

$ErrorActionPreference = 'Stop'

# Couleurs Graph : preset0 rouge, preset1 orange, preset3 jaune,
# preset4 vert, preset8 mauve
$Categories = [ordered]@{
    'Priorité · Urgent' = 'preset0'
    'Priorité · P1'     = 'preset1'
    'Priorité · P2'     = 'preset3'
    'Priorité · P3'     = 'preset4'
    'CONFIDENTIEL'      = 'preset8'
}

# ---------- Journal ----------
$dossierJournal = Join-Path $PSScriptRoot 'journaux'
New-Item -ItemType Directory -Path $dossierJournal -Force -WhatIf:$false | Out-Null
$fichierJournal = Join-Path $dossierJournal ("Creer-Categories_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))

function Write-Journal {
    param([string]$Message, [string]$Niveau = 'INFO')
    $ligne = '{0:yyyy-MM-dd HH:mm:ss} [{1}] {2}' -f (Get-Date), $Niveau, $Message
    Write-Host $ligne
    Add-Content -Path $fichierJournal -Value $ligne -Encoding UTF8 -WhatIf:$false
}

# ---------- Connexion ----------
Import-Module Microsoft.Graph.Users

if (-not (Get-MgContext)) {
    $portees = @('MailboxSettings.ReadWrite')
    if ($ToutesLesBoites) { $portees += 'User.Read.All' }
    Write-Journal "Connexion à Microsoft Graph (portées : $($portees -join ', '))"
    Connect-MgGraph -Scopes $portees -NoWelcome
}

# ---------- Boîtes ciblées ----------
if ($ToutesLesBoites) {
    Write-Journal 'Recherche des utilisateurs membres actifs...'
    $Boites = Get-MgUser -All -Filter 'accountEnabled eq true' -Property Mail, UserType |
        Where-Object { $_.Mail -and $_.UserType -eq 'Member' } |
        Select-Object -ExpandProperty Mail
}
Write-Journal ("{0} boîte(s) ciblée(s)" -f $Boites.Count)

# ---------- Traitement ----------
$erreurs = 0
foreach ($boite in $Boites) {
    try {
        $existantes = @(Get-MgUserOutlookMasterCategory -UserId $boite -All)
        foreach ($nom in $Categories.Keys) {
            $couleur = $Categories[$nom]
            $cat = $existantes | Where-Object { $_.DisplayName -eq $nom } | Select-Object -First 1

            if (-not $cat) {
                if ($PSCmdlet.ShouldProcess($boite, "Créer la catégorie $nom ($couleur)")) {
                    New-MgUserOutlookMasterCategory -UserId $boite -DisplayName $nom -Color $couleur | Out-Null
                    Write-Journal "$boite : catégorie $nom créée ($couleur)"
                }
            }
            elseif ($cat.Color -ne $couleur) {
                if ($PSCmdlet.ShouldProcess($boite, "Corriger la couleur de $nom ($($cat.Color) -> $couleur)")) {
                    Update-MgUserOutlookMasterCategory -UserId $boite -OutlookCategoryId $cat.Id -Color $couleur | Out-Null
                    Write-Journal "$boite : couleur de $nom corrigée ($($cat.Color) -> $couleur)"
                }
            }
            else {
                Write-Journal "$boite : catégorie $nom déjà correcte"
            }
        }
    }
    catch {
        $erreurs++
        Write-Journal "$boite : ERREUR - $($_.Exception.Message)" 'ERREUR'
    }
}

Write-Journal ("Terminé. Erreurs : {0}. Journal : {1}" -f $erreurs, $fichierJournal)
