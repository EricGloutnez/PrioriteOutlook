<#
.SYNOPSIS
    Retire les règles de réception « Priorité - ... » créées par
    Creer-ReglesReception.ps1.

.DESCRIPTION
    Seules les règles dont le nom commence par « Priorité - » et se
    termine par un niveau connu sont retirées. Les catégories et les
    courriels déjà classés ne sont pas touchés.
    -WhatIf affiche ce qui serait fait sans rien modifier.

    Module requis : ExchangeOnlineManagement

.PARAMETER Boites
    Adresses des boîtes ciblées.

.PARAMETER ToutesLesBoites
    Cible toutes les boîtes de type UserMailbox.

.EXAMPLE
    .\Supprimer-ReglesReception.ps1 -Boites eric@bonneville.ca -WhatIf
.EXAMPLE
    .\Supprimer-ReglesReception.ps1 -ToutesLesBoites
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Liste')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Liste')]
    [string[]]$Boites,

    [Parameter(Mandatory, ParameterSetName = 'Toutes')]
    [switch]$ToutesLesBoites
)

$ErrorActionPreference = 'Stop'

$NomsRegles = @('URGENT', 'P1', 'P2', 'P3', 'CONFIDENTIEL') | ForEach-Object { "Priorité - $_" }

# ---------- Journal ----------
$dossierJournal = Join-Path $PSScriptRoot 'journaux'
New-Item -ItemType Directory -Path $dossierJournal -Force -WhatIf:$false | Out-Null
$fichierJournal = Join-Path $dossierJournal ("Supprimer-ReglesReception_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))

function Write-Journal {
    param([string]$Message, [string]$Niveau = 'INFO')
    $ligne = '{0:yyyy-MM-dd HH:mm:ss} [{1}] {2}' -f (Get-Date), $Niveau, $Message
    Write-Host $ligne
    Add-Content -Path $fichierJournal -Value $ligne -Encoding UTF8 -WhatIf:$false
}

# ---------- Connexion ----------
Import-Module ExchangeOnlineManagement
if (-not (Get-ConnectionInformation | Where-Object { $_.State -eq 'Connected' })) {
    Write-Journal 'Connexion à Exchange Online...'
    Connect-ExchangeOnline -ShowBanner:$false
}

# ---------- Boîtes ciblées ----------
if ($ToutesLesBoites) {
    Write-Journal 'Recherche des boîtes utilisateurs...'
    $Boites = Get-EXOMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited |
        Select-Object -ExpandProperty PrimarySmtpAddress
}
Write-Journal ("{0} boîte(s) ciblée(s)" -f $Boites.Count)

# ---------- Traitement ----------
$erreurs = 0
foreach ($boite in $Boites) {
    try {
        $aRetirer = @(Get-InboxRule -Mailbox $boite | Where-Object { $NomsRegles -contains $_.Name })
        if ($aRetirer.Count -eq 0) {
            Write-Journal "$boite : aucune règle de priorité"
            continue
        }
        foreach ($regle in $aRetirer) {
            if ($PSCmdlet.ShouldProcess($boite, "Retirer la règle « $($regle.Name) »")) {
                Remove-InboxRule -Identity $regle.Identity -Confirm:$false
                Write-Journal "$boite : règle « $($regle.Name) » retirée"
            }
        }
    }
    catch {
        $erreurs++
        Write-Journal "$boite : ERREUR - $($_.Exception.Message)" 'ERREUR'
    }
}

Write-Journal ("Terminé. Erreurs : {0}. Journal : {1}" -f $erreurs, $fichierJournal)
