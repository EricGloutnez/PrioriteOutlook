<#
.SYNOPSIS
    Crée, dans chaque boîte ciblée, une règle de réception par niveau de
    priorité qui applique la catégorie de couleur correspondante.

.DESCRIPTION
    Règles créées (nom → mots recherchés dans l'objet → catégorie) :
      Priorité - URGENT        [URGENT]        _URGENT]        URGENT
      Priorité - P1            [P1]            _P1]            P1
      Priorité - P2            [P2]            _P2]            P2
      Priorité - P3            [P3]            _P3]            P3
      Priorité - CONFIDENTIEL  [CONFIDENTIEL]  [CONFIDENTIEL_  CONFIDENTIEL

    Les deux formes couvrent [P1] et [M_12345_P1] / [R_12345_P1], ainsi que
    [CONFIDENTIEL] et [CONFIDENTIEL_M_12345].
    À VALIDER À L'ESSAI : la façon dont SubjectContainsWords traite les
    crochets et le trait de soulignement (voir README, « À valider »).

    - Idempotent : une règle existante du même nom est mise à jour.
    - -WhatIf affiche ce qui serait fait sans rien modifier.
    - Exécuter Creer-Categories.ps1 d'abord, pour que les catégories
      existent avec leur couleur.

    Avertissement Exchange : modifier les règles par PowerShell peut
    supprimer les règles « client seulement » désactivées dans Outlook.
    Voir GUIDE-TI.md.

    Module requis : ExchangeOnlineManagement
      Install-Module ExchangeOnlineManagement -Scope CurrentUser

.PARAMETER Boites
    Adresses des boîtes ciblées.

.PARAMETER ToutesLesBoites
    Cible toutes les boîtes de type UserMailbox.

.EXAMPLE
    .\Creer-ReglesReception.ps1 -Boites eric@bonneville.ca -WhatIf
.EXAMPLE
    .\Creer-ReglesReception.ps1 -ToutesLesBoites
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Liste')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Liste')]
    [string[]]$Boites,

    [Parameter(Mandatory, ParameterSetName = 'Toutes')]
    [switch]$ToutesLesBoites
)

$ErrorActionPreference = 'Stop'

$Niveaux = @('URGENT', 'P1', 'P2', 'P3', 'CONFIDENTIEL')
$PrefixeRegle = 'Priorité - '

# ---------- Journal ----------
$dossierJournal = Join-Path $PSScriptRoot 'journaux'
New-Item -ItemType Directory -Path $dossierJournal -Force -WhatIf:$false | Out-Null
$fichierJournal = Join-Path $dossierJournal ("Creer-ReglesReception_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))

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
        $regles = @(Get-InboxRule -Mailbox $boite)
        foreach ($niveau in $Niveaux) {
            $nom = "$PrefixeRegle$niveau"
            $mots = if ($niveau -eq 'CONFIDENTIEL') { @('[CONFIDENTIEL]', '[CONFIDENTIEL_') }
                    else { @("[$niveau]", "_$niveau]") }
            $existante = $regles | Where-Object { $_.Name -eq $nom } | Select-Object -First 1

            if ($existante) {
                if ($PSCmdlet.ShouldProcess($boite, "Mettre à jour la règle « $nom »")) {
                    Set-InboxRule -Identity $existante.Identity `
                        -SubjectContainsWords $mots `
                        -ApplyCategory $niveau `
                        -StopProcessingRules $false `
                        -Force -Confirm:$false
                    Write-Journal "$boite : règle « $nom » mise à jour"
                }
            }
            else {
                if ($PSCmdlet.ShouldProcess($boite, "Créer la règle « $nom »")) {
                    New-InboxRule -Mailbox $boite -Name $nom `
                        -SubjectContainsWords $mots `
                        -ApplyCategory $niveau `
                        -StopProcessingRules $false `
                        -Force | Out-Null
                    Write-Journal "$boite : règle « $nom » créée"
                }
            }
        }
    }
    catch {
        $erreurs++
        Write-Journal "$boite : ERREUR - $($_.Exception.Message)" 'ERREUR'
    }
}

Write-Journal ("Terminé. Erreurs : {0}. Journal : {1}" -f $erreurs, $fichierJournal)
