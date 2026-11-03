<#
.SYNOPSIS
    Active ou désactive, dans l'Outlook classique du poste local, la mise
    en forme conditionnelle des courriels selon leur balise de priorité.

.DESCRIPTION
    Ajoute une règle de mise en forme par niveau dans l'affichage actuel
    de la boîte de réception : texte en gras et en couleur quand l'objet
    contient la balise.

      URGENT        rouge
      P1            bordeaux (Outlook n'offre pas l'orange ici)
      P2            olive    (le jaune vif serait illisible)
      P3            vert
      CONFIDENTIEL  mauve

    - Outlook classique pour Windows seulement : le nouvel Outlook et
      les applications mobiles n'offrent pas cette fonction.
    - S'applique à l'affichage actuel de la boîte de réception de la
      boîte par défaut du profil.
    - Fermer puis rouvrir Outlook si l'affichage ne se met pas à jour.
    - -WhatIf affiche ce qui serait fait sans rien modifier.

.PARAMETER Activer
    Ajoute (ou remplace) les règles de mise en forme.

.PARAMETER Desactiver
    Retire les règles de mise en forme ajoutées par ce script.

.EXAMPLE
    .\MiseEnFormeConditionnelle.ps1 -Activer -WhatIf
.EXAMPLE
    .\MiseEnFormeConditionnelle.ps1 -Desactiver
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Activer')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Activer')]
    [switch]$Activer,

    [Parameter(Mandatory, ParameterSetName = 'Desactiver')]
    [switch]$Desactiver
)

$ErrorActionPreference = 'Stop'

# Valeurs de l'énumération OlColor d'Outlook
$Formats = [ordered]@{
    'URGENT'       = 10   # olColorRed
    'P1'           = 2    # olColorMaroon
    'P2'           = 4    # olColorOlive
    'P3'           = 3    # olColorGreen
    'CONFIDENTIEL' = 6    # olColorPurple
}
$PrefixeRegle = 'Priorité - '

# ---------- Journal ----------
$dossierJournal = Join-Path $PSScriptRoot 'journaux'
New-Item -ItemType Directory -Path $dossierJournal -Force -WhatIf:$false | Out-Null
$fichierJournal = Join-Path $dossierJournal ("MiseEnFormeConditionnelle_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))

function Write-Journal {
    param([string]$Message, [string]$Niveau = 'INFO')
    $ligne = '{0:yyyy-MM-dd HH:mm:ss} [{1}] {2}' -f (Get-Date), $Niveau, $Message
    Write-Host $ligne
    Add-Content -Path $fichierJournal -Value $ligne -Encoding UTF8 -WhatIf:$false
}

# ---------- Outlook ----------
try {
    $outlook = New-Object -ComObject Outlook.Application
} catch {
    throw "Outlook classique est introuvable sur ce poste : $($_.Exception.Message)"
}
$ns     = $outlook.GetNamespace('MAPI')
$inbox  = $ns.GetDefaultFolder(6)          # olFolderInbox
$vue    = $inbox.CurrentView

if ($vue.ViewType -ne 0) {                 # 0 = olTableView
    throw "L'affichage actuel de la boîte de réception n'est pas un affichage en tableau ; impossible d'y ajouter la mise en forme."
}
$regles = $vue.AutoFormatRules
Write-Journal "Affichage : « $($vue.Name) » ($($regles.Count) règle(s) de mise en forme existante(s))"

# Retirer les règles déjà créées par ce script (en partant de la fin).
for ($i = $regles.Count; $i -ge 1; $i--) {
    $nomExistant = $regles.Item($i).Name
    if ($nomExistant -like "$PrefixeRegle*") {
        if ($PSCmdlet.ShouldProcess($nomExistant, 'Retirer la règle de mise en forme')) {
            $regles.Remove($i)
            Write-Journal "Règle « $nomExistant » retirée"
        }
    }
}

if ($Activer) {
    foreach ($niveau in $Formats.Keys) {
        $nom = "$PrefixeRegle$niveau"
        if ($PSCmdlet.ShouldProcess($nom, 'Ajouter la règle de mise en forme')) {
            $r = $regles.Add($nom)
            # Filtre DASL : couvre [P1] et [M_12345_P1], ou [CONFIDENTIEL] et
            # [CONFIDENTIEL_M_12345]. À VALIDER À L'ESSAI : crochets dans LIKE.
            $sujet = '"urn:schemas:httpmail:subject"'
            $r.Filter = if ($niveau -eq 'CONFIDENTIEL') {
                "$sujet LIKE '%[CONFIDENTIEL%'"
            } else {
                "$sujet LIKE '%[$niveau]%' OR $sujet LIKE '%_$niveau]%'"
            }
            $r.Font.Bold  = $true
            $r.Font.Color = $Formats[$niveau]
            $r.Enabled    = $true
            Write-Journal "Règle « $nom » ajoutée"
        }
    }
}

if (-not $WhatIfPreference) {
    $regles.Save()
    $vue.Save()
    $vue.Apply()
    Write-Journal 'Affichage enregistré et appliqué'
}

Write-Journal "Terminé. Journal : $fichierJournal"
